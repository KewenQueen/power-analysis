// 飞书电子表格读取代理（Supabase Edge Function）
//
// 用途：纯前端应用无法安全保存飞书 App Secret，浏览器直连 open.feishu.cn 也会被 CORS 拦截。
// 因此由本函数用「应用凭证 → tenant_access_token」代取表格数据，返回二维数组给前端解析。
//
// 需要配置的环境变量（Supabase → Edge Functions → Secrets，或 `supabase secrets set`）：
//   FEISHU_APP_ID      飞书开放平台应用 App ID
//   FEISHU_APP_SECRET  飞书开放平台应用 App Secret
// 可选：
//   LARK_API_BASE      默认 https://open.feishu.cn/open-apis（Lark 国际版可改为 https://open.larksuite.com/open-apis）
//   FEISHU_PROXY_TOKEN 设置后，前端必须在请求头 x-proxy-token 中带上同样的值（用于加一层简单的访问口令）
//
// 部署：supabase functions deploy feishu-sheet-proxy --project-ref <project-ref>
// 说明：读取的是「应用被授权访问」的文档，需要在飞书文档里把该应用加为协作者（或文档对组织内可见且应用有 sheets 权限）。

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, apikey, content-type, x-client-info, x-proxy-token',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
};

const API_BASE = (Deno.env.get('LARK_API_BASE') || 'https://open.feishu.cn/open-apis').replace(/\/+$/, '');
const MAX_ROWS = 4000;
const MAX_COLUMNS = 60;

type Json = Record<string, unknown>;

const json = (body: Json, status = 200) =>
  new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, 'Content-Type': 'application/json; charset=utf-8' },
  });

let cachedToken = '';
let tokenExpiresAt = 0;

class ProxyError extends Error {
  code: string;
  hint: string;
  status: number;
  constructor(message: string, code = 'PROXY_ERROR', hint = '', status = 400) {
    super(message);
    this.code = code;
    this.hint = hint;
    this.status = status;
  }
}

/** 把飞书返回的错误码翻译成可执行的中文提示 */
function describeFeishuError(code: number | string, message: string, stage: string): ProxyError {
  const numeric = Number(code);
  if (numeric === 99991663 || numeric === 99991661 || numeric === 99991664) {
    return new ProxyError(
      `飞书应用凭证无效（${message || code}）`,
      'FEISHU_CREDENTIAL_INVALID',
      '请检查 FEISHU_APP_ID / FEISHU_APP_SECRET 是否正确，以及应用是否已发布可用版本。',
      502,
    );
  }
  if (numeric === 1061004 || numeric === 1310213 || numeric === 91403 || numeric === 1770001 || numeric === 1254302) {
    return new ProxyError(
      `飞书未授权访问该文档（${message || code}）`,
      'FEISHU_FORBIDDEN',
      '请在飞书文档中点「···→ 更多 → 添加文档协作者」，把该应用（机器人）加为可阅读协作者；或让文档对组织内成员可见并给应用开通 sheets:spreadsheet:readonly 权限。',
      403,
    );
  }
  if (numeric === 1310211 || numeric === 1310212 || numeric === 90215 || numeric === 90202) {
    return new ProxyError(
      `表格 token 或 sheet 不存在（${message || code}）`,
      'FEISHU_NOT_FOUND',
      '请确认链接是「电子表格」而不是多维表格/文档，并确认 Sheet 名称拼写正确。',
      404,
    );
  }
  return new ProxyError(
    `${stage}失败：${message || code}`,
    'FEISHU_API_ERROR',
    '可稍后重试；若持续失败请检查飞书开放平台应用状态与权限。',
    502,
  );
}

async function getTenantToken(): Promise<string> {
  if (cachedToken && Date.now() < tokenExpiresAt) return cachedToken;
  const appId = Deno.env.get('FEISHU_APP_ID');
  const appSecret = Deno.env.get('FEISHU_APP_SECRET');
  if (!appId || !appSecret) {
    throw new ProxyError(
      '云文档代理尚未配置飞书应用凭证',
      'PROXY_NOT_CONFIGURED',
      '请为 Edge Function 配置环境变量 FEISHU_APP_ID 与 FEISHU_APP_SECRET 后重新部署。',
      503,
    );
  }
  const response = await fetch(`${API_BASE}/auth/v3/tenant_access_token/internal`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json; charset=utf-8' },
    body: JSON.stringify({ app_id: appId, app_secret: appSecret }),
  });
  const result = await response.json().catch(() => ({} as Json)) as Json;
  const token = result.tenant_access_token as string | undefined;
  if (!response.ok || !token) {
    throw describeFeishuError(result.code as number, String(result.msg || ''), '获取飞书访问令牌');
  }
  cachedToken = token;
  const expiresIn = Number(result.expire) || 7200;
  tokenExpiresAt = Date.now() + Math.max(expiresIn - 120, 60) * 1000;
  return token;
}

async function feishuGet(token: string, path: string, stage: string): Promise<Json> {
  const response = await fetch(`${API_BASE}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const result = await response.json().catch(() => ({} as Json)) as Json;
  if (!response.ok || Number(result.code) !== 0) {
    throw describeFeishuError(result.code as number, String(result.msg || ''), stage);
  }
  return (result.data || {}) as Json;
}

type SheetMeta = {
  sheetId: string;
  title: string;
  index: number;
  hidden: boolean;
  rowCount: number;
  columnCount: number;
};

async function listSheets(token: string, spreadsheetToken: string): Promise<SheetMeta[]> {
  const data = await feishuGet(
    token,
    `/sheets/v3/spreadsheets/${encodeURIComponent(spreadsheetToken)}/sheets/query`,
    '读取工作表列表',
  );
  const sheets = Array.isArray(data.sheets) ? data.sheets as Json[] : [];
  return sheets.map((sheet, position) => {
    const grid = (sheet.grid_properties || {}) as Json;
    return {
      sheetId: String(sheet.sheet_id || ''),
      title: String(sheet.title || `Sheet${position + 1}`),
      index: Number(sheet.index ?? position),
      hidden: Boolean(sheet.hidden),
      rowCount: Number(grid.row_count || 0),
      columnCount: Number(grid.column_count || 0),
    };
  }).sort((a, b) => a.index - b.index);
}

function columnLetter(index: number): string {
  let n = Math.max(1, index);
  let label = '';
  while (n > 0) {
    n -= 1;
    label = String.fromCharCode(65 + (n % 26)) + label;
    n = Math.floor(n / 26);
  }
  return label;
}

async function readValues(token: string, spreadsheetToken: string, range: string): Promise<unknown[][]> {
  const data = await feishuGet(
    token,
    `/sheets/v2/spreadsheets/${encodeURIComponent(spreadsheetToken)}/values/${encodeURIComponent(range)}?valueRenderOption=ToString`,
    '读取表格数据',
  );
  const valueRange = (data.valueRange || {}) as Json;
  const values = valueRange.values;
  if (!Array.isArray(values)) return [];
  return values.map((row) => (Array.isArray(row) ? row.map((cell) => {
    if (cell === null || cell === undefined) return '';
    if (typeof cell === 'object') {
      const rich = cell as Json;
      if (typeof rich.text === 'string') return rich.text;
      if (Array.isArray(rich.segments)) {
        return (rich.segments as Json[]).map((seg) => String(seg.text || '')).join('');
      }
      return '';
    }
    return String(cell);
  }) : []));
}

/** 支持完整链接 / 带查询参数的链接 / 裸 token */
function parseSpreadsheetToken(raw: unknown): string {
  const value = String(raw || '').trim();
  if (!value) return '';
  const fromUrl = value.match(/\/(?:sheets|spreadsheets)\/([A-Za-z0-9]{6,})/);
  if (fromUrl) return fromUrl[1];
  if (/^[A-Za-z0-9]{6,}$/.test(value)) return value;
  return '';
}

function pickSheet(sheets: SheetMeta[], wantedId: string, wantedName: string): SheetMeta | null {
  if (wantedId) {
    const byId = sheets.find((sheet) => sheet.sheetId === wantedId);
    if (byId) return byId;
  }
  if (wantedName) {
    const exact = sheets.find((sheet) => sheet.title === wantedName);
    if (exact) return exact;
    const loose = sheets.find((sheet) => sheet.title.includes(wantedName) || wantedName.includes(sheet.title));
    if (loose) return loose;
  }
  const visible = sheets.filter((sheet) => !sheet.hidden);
  return visible[0] || sheets[0] || null;
}

function checkProxyToken(request: Request) {
  const expected = Deno.env.get('FEISHU_PROXY_TOKEN');
  if (!expected) return;
  const provided = request.headers.get('x-proxy-token') || '';
  if (provided !== expected) {
    throw new ProxyError('云文档代理访问口令不正确', 'PROXY_TOKEN_INVALID', '请检查请求头 x-proxy-token 是否与部署时配置的口令一致。', 401);
  }
}

Deno.serve(async (request: Request) => {
  if (request.method === 'OPTIONS') return new Response('ok', { headers: corsHeaders });

  const configured = Boolean(Deno.env.get('FEISHU_APP_ID') && Deno.env.get('FEISHU_APP_SECRET'));

  if (request.method === 'GET') {
    return json({
      ok: true,
      service: 'feishu-sheet-proxy',
      configured,
      hint: configured
        ? 'POST { "action": "read", "url": "<飞书表格链接>" } 即可取回表格数据'
        : '尚未配置 FEISHU_APP_ID / FEISHU_APP_SECRET，请先设置环境变量并重新部署',
    });
  }

  if (request.method !== 'POST') {
    return json({ ok: false, error: '只支持 GET / POST', code: 'METHOD_NOT_ALLOWED' }, 405);
  }

  try {
    checkProxyToken(request);
    const body = await request.json().catch(() => ({} as Json)) as Json;
    const spreadsheetToken = parseSpreadsheetToken(body.url ?? body.spreadsheetToken ?? body.documentUrl);
    if (!spreadsheetToken) {
      throw new ProxyError(
        '未能从链接中识别表格 token',
        'INVALID_URL',
        '请粘贴形如 https://xxx.larkoffice.com/sheets/XXXXXXXX 的飞书「电子表格」链接。',
        400,
      );
    }

    const token = await getTenantToken();
    const sheets = await listSheets(token, spreadsheetToken);

    if (String(body.action || 'read') === 'meta') {
      return json({ ok: true, spreadsheetToken, sheets });
    }

    const sheet = pickSheet(sheets, String(body.sheetId || ''), String(body.sheetName || ''));
    if (!sheet) {
      throw new ProxyError('该表格中没有任何工作表', 'EMPTY_SPREADSHEET', '请确认链接指向的是含数据的电子表格。', 404);
    }

    const rowLimit = Math.min(Math.max(Number(body.maxRows) || MAX_ROWS, 1), MAX_ROWS);
    const columnLimit = Math.min(Math.max(Number(body.maxColumns) || MAX_COLUMNS, 1), MAX_COLUMNS);
    const lastRow = Math.min(sheet.rowCount || rowLimit, rowLimit);
    const lastColumn = Math.min(sheet.columnCount || columnLimit, columnLimit);
    const range = typeof body.range === 'string' && body.range.trim()
      ? String(body.range).trim()
      : `${sheet.sheetId}!A1:${columnLetter(lastColumn)}${lastRow}`;

    const values = await readValues(token, spreadsheetToken, range);

    return json({
      ok: true,
      spreadsheetToken,
      sheet: {
        sheetId: sheet.sheetId,
        title: sheet.title,
        index: sheet.index,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
      },
      sheets,
      range,
      values,
    });
  } catch (error) {
    if (error instanceof ProxyError) {
      return json({ ok: false, error: error.message, code: error.code, hint: error.hint }, error.status);
    }
    console.error(error);
    return json({
      ok: false,
      error: `云文档代理异常：${(error as Error)?.message || '未知错误'}`,
      code: 'PROXY_EXCEPTION',
      hint: '请稍后重试；若持续失败请查看 Edge Function 日志。',
    }, 500);
  }
});
