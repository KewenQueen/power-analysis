const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

const json = (body: unknown, status = 200) => new Response(JSON.stringify(body), {
  status,
  headers: { ...corsHeaders, 'Content-Type': 'application/json; charset=utf-8' },
});

const apiBase = Deno.env.get('LARK_API_BASE') || 'https://open.feishu.cn/open-apis';
let cachedToken = '';
let tokenExpiresAt = 0;

async function getTenantToken() {
  if (cachedToken && Date.now() < tokenExpiresAt) return cachedToken;
  const appId = Deno.env.get('FEISHU_APP_ID');
  const appSecret = Deno.env.get('FEISHU_APP_SECRET');
  if (!appId || !appSecret) throw new Error('云导出服务尚未配置飞书应用凭证');
  const response = await fetch(`${apiBase}/auth/v3/tenant_access_token/internal`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ app_id: appId, app_secret: appSecret }),
  });
  const result = await response.json();
  if (!response.ok || result.code !== 0 || !result.tenant_access_token) {
    throw new Error(result.msg || '获取飞书访问令牌失败');
  }
  cachedToken = result.tenant_access_token;
  tokenExpiresAt = Date.now() + Math.max(60, Number(result.expire || 7200) - 300) * 1000;
  return cachedToken;
}

async function larkRequest(path: string, init: RequestInit = {}) {
  const token = await getTenantToken();
  const response = await fetch(`${apiBase}${path}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(init.headers || {}),
    },
  });
  const result = await response.json().catch(() => ({}));
  if (!response.ok || result.code !== 0) throw new Error(result.msg || `飞书接口请求失败（${response.status}）`);
  return result.data || {};
}

function parseToken(value: string, type: 'sheet' | 'folder') {
  const text = String(value || '').trim();
  if (!text) return '';
  if (!text.includes('/')) return text;
  const pattern = type === 'sheet' ? /\/(?:sheets|spreadsheets)\/([A-Za-z0-9_-]+)/ : /\/folder\/([A-Za-z0-9_-]+)/;
  return text.match(pattern)?.[1] || '';
}

function columnName(index: number) {
  let value = index;
  let result = '';
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function parseStartCell(cell: string) {
  const match = String(cell || 'A1').toUpperCase().match(/^([A-Z]+)([1-9]\d*)$/);
  if (!match) throw new Error('写入起点格式不正确');
  let column = 0;
  for (const char of match[1]) column = column * 26 + char.charCodeAt(0) - 64;
  return { column, row: Number(match[2]) };
}

async function listSheets(spreadsheetToken: string) {
  const data = await larkRequest(`/sheets/v3/spreadsheets/${spreadsheetToken}/sheets/query`);
  return Array.isArray(data.sheets) ? data.sheets : [];
}

async function addSheet(spreadsheetToken: string, title: string) {
  const data = await larkRequest(`/sheets/v2/spreadsheets/${spreadsheetToken}/sheets_batch_update`, {
    method: 'POST',
    body: JSON.stringify({ requests: [{ addSheet: { properties: { title } } }] }),
  });
  const replies = data.replies || [];
  const properties = replies[0]?.addSheet?.properties;
  if (properties?.sheetId) return properties.sheetId;
  const sheets = await listSheets(spreadsheetToken);
  const created = sheets.find((sheet: Record<string, unknown>) => sheet.title === title);
  if (!created) throw new Error('新建 Sheet 后未能读取其标识');
  return created.sheet_id || created.sheetId;
}

async function writeValues(spreadsheetToken: string, sheetId: string, startCell: string, values: unknown[][]) {
  if (!values.length || !values[0]?.length) throw new Error('没有可写入的数据');
  const width = Math.max(...values.map((row) => row.length));
  const normalized = values.map((row) => Array.from({ length: width }, (_, index) => row[index] ?? ''));
  const start = parseStartCell(startCell);
  const endCell = `${columnName(start.column + width - 1)}${start.row + normalized.length - 1}`;
  const range = `${sheetId}!${String(startCell).toUpperCase()}:${endCell}`;
  await larkRequest(`/sheets/v2/spreadsheets/${spreadsheetToken}/values/${encodeURIComponent(range)}`, {
    method: 'PUT',
    body: JSON.stringify({ valueRange: { range, values: normalized } }),
  });
}

Deno.serve(async (request) => {
  if (request.method === 'OPTIONS') return new Response('ok', { headers: corsHeaders });
  if (request.method !== 'POST') return json({ ok: false, error: '仅支持 POST 请求' }, 405);
  try {
    const payload = await request.json();
    const values = Array.isArray(payload.values) ? payload.values : [];
    if (!values.length) throw new Error('导出数据为空');
    if (values.length * Math.max(...values.map((row: unknown[]) => Array.isArray(row) ? row.length : 0)) > 20000) {
      throw new Error('单次云导出最多支持 20,000 个单元格');
    }

    let spreadsheetToken = '';
    let documentUrl = '';
    let createdDocument = false;
    if (payload.action === 'create') {
      const title = String(payload.title || '').trim();
      if (!title) throw new Error('请输入云文档名称');
      const folderToken = parseToken(payload.folder || '', 'folder');
      if (payload.folder && !folderToken) throw new Error('保存区域链接或 folder token 格式不正确');
      const data = await larkRequest('/sheets/v3/spreadsheets', {
        method: 'POST',
        body: JSON.stringify({ title, ...(folderToken ? { folder_token: folderToken } : {}) }),
      });
      const spreadsheet = data.spreadsheet || data;
      spreadsheetToken = spreadsheet.spreadsheet_token;
      documentUrl = spreadsheet.url || `https://www.feishu.cn/sheets/${spreadsheetToken}`;
      createdDocument = true;
    } else if (payload.action === 'write') {
      spreadsheetToken = parseToken(payload.documentUrl || '', 'sheet');
      if (!spreadsheetToken) throw new Error('仅支持有效的飞书电子表格链接');
      documentUrl = String(payload.documentUrl);
    } else {
      throw new Error('不支持的云导出操作');
    }

    let sheetId = '';
    const sheetName = String(payload.sheetName || '').trim();
    if (!sheetName) throw new Error('请输入 Sheet 名称');
    if (createdDocument) {
      const sheets = await listSheets(spreadsheetToken);
      const firstSheet = sheets[0];
      if (!firstSheet) throw new Error('新建云文档后未找到默认 Sheet');
      sheetId = String(firstSheet.sheet_id || firstSheet.sheetId);
    } else if (payload.sheetMode === 'existing') {
      const sheets = await listSheets(spreadsheetToken);
      const matched = sheets.find((sheet: Record<string, unknown>) => sheet.title === sheetName);
      if (!matched) throw new Error(`未找到 Sheet「${sheetName}」`);
      sheetId = String(matched.sheet_id || matched.sheetId);
    } else {
      sheetId = await addSheet(spreadsheetToken, sheetName);
    }
    await writeValues(spreadsheetToken, sheetId, payload.startCell || 'A1', values);
    return json({ ok: true, url: documentUrl, spreadsheetToken, sheetId });
  } catch (error) {
    console.error(error);
    return json({ ok: false, error: error instanceof Error ? error.message : '云文档导出失败' }, 400);
  }
});
