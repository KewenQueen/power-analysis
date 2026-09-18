/* global Chart, Papa, XlsxPopulate */

// ---------- Supabase data API ----------
async function getSupabaseClient() {
  await window.powerAuth.ready;
  const client = window.powerAuth.client;
  if (!client || !window.powerAuth.getUser()) throw new Error('请先登录账号');
  return client;
}

function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  let binary = '';
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

function base64ToArrayBuffer(value) {
  const binary = atob(value || '');
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

function normalizeTemplateRow(rec) {
  const payload = rec.data || {};
  return {
    id: rec.id,
    name: rec.name,
    project: rec.project || rec.name,
    fileName: payload.file_name || `${rec.name}.xlsx`,
    sizeBytes: Number(payload.size_bytes || 0),
    version: Number(payload.version || 1),
    createdAt: rec.created_at,
    updatedAt: payload.updated_at || rec.created_at,
    createdBy: rec.user_id || '',
    updatedBy: payload.updated_by || '',
    data: payload,
  };
}

const templateApi = {
  async list() {
    const client = await getSupabaseClient();
    const { data, error } = await client.from('templates').select('*').order('created_at', { ascending: false });
    if (error) throw error;
    return (data || []).map(normalizeTemplateRow);
  },
  async get(id) {
    const client = await getSupabaseClient();
    const { data, error } = await client.from('templates').select('*').eq('id', id).single();
    if (error) throw error;
    return normalizeTemplateRow(data);
  },
  async download(id) {
    const record = await this.get(id);
    const base64 = record.data && record.data.base64;
    if (!base64) throw new Error('模板文件内容为空');
    return { buffer: base64ToArrayBuffer(base64), version: record.version, updatedAt: record.updatedAt, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' };
  },
  async upload({ file, name, fileName }) {
    if (!window.powerAuth.isAdmin()) throw new Error('只有管理员可以上传模板');
    const client = await getSupabaseClient();
    const user = window.powerAuth.getUser();
    const buffer = await file.arrayBuffer();
    const payload = { base64: arrayBufferToBase64(buffer), file_name: fileName || file.name, size_bytes: buffer.byteLength, version: 1, updated_at: new Date().toISOString(), updated_by: user.email };
    const { data, error } = await client.from('templates').insert({ user_id: user.id, name, project: name, data: payload }).select().single();
    if (error) throw error;
    return normalizeTemplateRow(data);
  },
  async update(id, { file, name, fileName } = {}) {
    if (!window.powerAuth.isAdmin()) throw new Error('只有管理员可以修改模板');
    const client = await getSupabaseClient();
    const current = await this.get(id);
    const payload = { ...(current.data || {}) };
    if (file) {
      const buffer = await file.arrayBuffer();
      payload.base64 = arrayBufferToBase64(buffer);
      payload.size_bytes = buffer.byteLength;
      payload.file_name = fileName || file.name;
      payload.version = Number(payload.version || 1) + 1;
    }
    payload.updated_at = new Date().toISOString();
    payload.updated_by = window.powerAuth.getUser().email;
    const changes = { data: payload };
    if (name !== undefined && name !== null) {
      changes.name = name;
      changes.project = name;
    }
    const { data, error } = await client.from('templates').update(changes).eq('id', id).select().single();
    if (error) throw error;
    return normalizeTemplateRow(data);
  },
  async remove(id) {
    if (!window.powerAuth.isAdmin()) throw new Error('只有管理员可以删除模板');
    const client = await getSupabaseClient();
    const { error } = await client.from('templates').delete().eq('id', id);
    if (error) throw error;
    return { ok: true };
  },
};

function clearTransientCaches() {
  try {
    const HISTORY_KEY = 'qepm_operation_history';
    const DEFAULT_COLUMNS_KEY = 'qepm_default_columns';
    const keepKeys = new Set([HISTORY_KEY, DEFAULT_COLUMNS_KEY]);
    const keepValues = {};
    keepKeys.forEach((key) => {
      const value = localStorage.getItem(key);
      if (value !== null) keepValues[key] = value;
    });
    const removeKeys = [];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && !keepKeys.has(key)) removeKeys.push(key);
    }
    removeKeys.forEach((key) => localStorage.removeItem(key));
    Object.keys(keepValues).forEach((key) => localStorage.setItem(key, keepValues[key]));
    sessionStorage.clear();
    // 保留 IndexedDB 模板库 qepm_template_store 与历史缓存 qepm_history_store：
    // 模板与历史结果都需要跨刷新保留，避免模板丢失或历史回溯能力失效。
    if (typeof window !== 'undefined' && window.__QEP_APP_READY__ && typeof resetWorkspaceState === 'function') {
      resetWorkspaceState();
    }
  } catch (e) {
    console.warn('clearTransientCaches failed', e);
  }
}
clearTransientCaches();

const TARGET_SHEET_NAME = '6-QEPM原始数据';
const PREVIEW_SHEET_NAME = '3_功耗采集数据表';
const DB_NAME = 'qepm_template_store';
const DB_VERSION = 2;
const STORE_NAME = 'templates';
const HISTORY_STORE_NAME = 'qepm_history_store';
const WRITE_BATCH_SIZE = 2000;
const UI_YIELD_EVERY_BATCHES = 2;

// Strictly matches a pure numeric string (optional sign, integer/decimal, scientific notation).
const NUMERIC_RE = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;

// ---- 预览结果列配置 ----
// 预览需要展示「CSV 写入模板后」的处理结果值；但导出的 workbook 仍要保留模板公式。
// xlsx-populate 读取公式格时拿到的是模板缓存值，所以这里保留一层“仅预览使用”的 JS 侧结果覆盖，
// 不回写 workbook，只用于 render preview。
// 说明：6-QEPM原始数据 的表头位于第 3 行（csvRows 下标 2），数据从第 5 行（下标 4）开始，
// 对应模板里 5-测试原始数据 的 AVERAGE / HLOOKUP 公式结构。
const QEPM_HEADER_ROW_INDEX = 2;
const QEPM_DATA_START_INDEX = 4;

// 「3_功耗采集数据表」常见列位（1 基）。真实模板列位仍会在运行时按公式动态识别。
const COL_L_CHANNEL = 12;         // fallback: 采样通道编号（QEPM）
const COL_M_VAVG = 13;            // fallback: Vavg(V)
const COL_N_IAVG = 14;            // fallback: Iavg(mA)
const COL_O_PAVG = 15;            // fallback: Pavg(mW)
const COL_P_DCDC_NAME = 16;       // fallback: DCDC
const COL_Q_DCDC_EFF = 17;        // fallback: DCDC Efficiency
const COL_R_LDO_NAME = 18;        // fallback: LDO
const COL_S_LDO_EFF = 19;         // fallback: LDO Efficiency
const COL_T_TOTAL_EFF = 20;       // fallback: Total Efficiency
const COL_U_PWR_VBATT = 21;       // fallback: Power @VBATT
const COL_V_INCLUDE_SUM = 22;     // fallback: Include in Sum?
const COL_W_POWER_IN_SUM = 23;    // fallback: Power in Sum

// Admin identity is controlled by POWER_ANALYSIS_CONFIG.adminEmails.
const ADMIN_STORAGE_KEY = 'qepm_admin_session';
const DEFAULT_COLUMNS_STORAGE_KEY = 'qepm_default_columns';
const FALLBACK_DEFAULT_COLUMNS = 'A-D,L-O';

// Target header keywords for auto-detecting default display columns
const AUTO_DETECT_HEADERS = [
  { keywords: ['模块'], exclude: ['子模块', 'tiny', '小模块'] },
  { keywords: ['子模块'] },
  { keywords: ['tiny模块', 'tiny'] },
  { keywords: ['power domain', 'power_domain', 'powerdomain'] },
  { keywords: ['采样通道编号', 'QEPM', '通道编号'], exclude: [] },
  { keywords: ['Vavg', 'V_avg', 'V avg', 'Vavg(V)'] },
  { keywords: ['Iavg', 'I_avg', 'I avg', 'Iavg(mA)'] },
  { keywords: ['Pavg', 'P_avg', 'P avg', 'Pavg(mW)'] },
];

// Column header keywords that identify "采样结果" value columns (should not be truncated)
const VALUE_COL_KEYWORDS = ['vavg', 'iavg', 'pavg', 'v_avg', 'i_avg', 'p_avg'];

// Operation history
const HISTORY_STORAGE_KEY = 'qepm_operation_history';
const HISTORY_MAX = 200;
const HISTORY_CACHE_MAX = 50;
const HISTORY_CSV_CACHE_LIMIT_BYTES = 20 * 1024 * 1024;
const HISTORY_CSV_ROW_SOFT_LIMIT = 5000;

// Chart plotting cap (avoid too many points causing lag)
const CHART_MAX_POINTS = 1500;

// Waveform metric types (for channel click → chart)
const METRIC_CURRENT  = 'current';
const METRIC_VOLTAGE  = 'voltage';
const METRIC_POWER    = 'power';
const METRIC_LABELS   = { [METRIC_CURRENT]: '电流 (I)', [METRIC_VOLTAGE]: '电压 (V)', [METRIC_POWER]: '功率 (P)' };
const METRIC_UNITS    = { [METRIC_CURRENT]: 'mA', [METRIC_VOLTAGE]: 'V', [METRIC_POWER]: 'mW' };
const METRIC_REGEX    = {
  [METRIC_CURRENT]: /(_i$|_i_|_curr|current|电流|iavg|\(ma\)|_ma$)/i,
  [METRIC_VOLTAGE]: /(_v$|_v_|_volt|voltage|电压|vavg|\(v\)|_vbat)/i,
  [METRIC_POWER]:   /(_p$|_pwr|power|功率|pavg|\(mw\))/i,
};

// Template pagination
const TEMPLATE_PAGE_SIZE = 5;

// Compare snapshots
const COMPARE_MAX = 10;

// Batch upload cap
const BATCH_MAX = 8;

// Compare column semantics for the 功耗采集数据表 (0-based sheet column indices):
//  - columns [0, 13)  => A..M common headers (M = 采样通道编号 QEPM)
//  - columns 13,14,15 => N / O / P value columns (P = Pavg(mW) = Power)
const COMPARE_COMMON_END = 13;      // columns < 13 are common titles
const COMPARE_VALUE_COLS = [13, 14, 15]; // N, O, P
const CHANNEL_COL = 12;             // M column, QEPM 采样通道编号 (fallback)
const POWER_COL_DEFAULT = 15;       // P column, Pavg(mW)

// Palette for datasets / compare groups
const DS_COLORS = ['#4f46e5', '#ec4899', '#0ea5e9', '#f59e0b', '#10b981', '#ef4444'];

const state = {
  // ---- Datasets (batch upload) ----
  // each: { id, fileName, csvRows, csvHeaderRowIndex, csvDataStartRow,
  //         processedRows, processedMerges, workbookBlob, projectId, projectName }
  datasets: [],
  activeDatasetId: null,

  // ---- Active-dataset mirrors (kept in sync for legacy code paths) ----
  csvRows: [],
  processedRows: [],
  processedMerges: [],
  fileName: '',
  workbookBlob: null,

  activeProject: null,
  isAdmin: false,

  // CSV column-search & chart states
  csvSearchKeyword: '',
  csvSearchMode: 'manual', // manual | channel
  csvMatchedCols: [],
  csvChartSelectedCols: [],
  csvHeaderRowIndex: -1,
  csvDataStartRow: -1,

  // Processed preview column filter (0-based)
  processedColIndices: null,

  // Compare common-column selections (0-based, < COMPARE_COMMON_END).
  // Global default for all projects when no per-project override.
  // null = default: show QEPM + Vavg (auto-detected per snapshot).
  compareCommonColIndices: null,
  // Per-project overrides: { [projectId]: number[] | null }.
  // When a project entry is:
  //   - an array: use that explicit common-column list for snapshots of the project
  //   - null: force the project back to default behaviour (global setting or QEPM+Vavg)
  //   - missing: fall back to the global compareCommonColIndices, then QEPM+Vavg
  compareCommonColIndicesByProject: {},

  // Cross-project custom channel selection: { [groupId]: { [projectKey]: { snapId, channel, colIndex } } }
  crossProjectChannelSelection: {},
  crossProjectChartInstances: {},

  // Chart instance
  chartInstance: null,
  // Template rename modal context
  templateRenameContext: null,
  // Waveform metric selection (channel click chart)
  channelMetricType: METRIC_CURRENT, // default: show current
  channelMetricCols: { [METRIC_CURRENT]: [], [METRIC_VOLTAGE]: [], [METRIC_POWER]: [] },
  channelAllMatchedCols: [], // all matched columns for the clicked channel
  // Cross-project waveform metric type (per group)
  crossProjectMetricType: {}, // { [groupId]: METRIC_CURRENT | METRIC_VOLTAGE | METRIC_POWER }

  // Template management views + pagination
  templateItems: [],
  templateView: 'list',
  templateDetailId: null,
  templatePage: 1,
  // Whether the current signed-in user is an admin according to the backend.
  // The local admin toggle (state.isAdmin) still controls UI visibility, but
  // any write operation ultimately depends on the backend accepting the call.
  serverIsAdmin: false,
  serverUser: null,
  templatesLoadError: null,
  historyItems: [],

  // Compare groups (in-memory). Each group is an independent comparison table.
  // { id, name, snapshots: [], highlightDiff: false }
  compareGroups: [{ id: 'grp_init', name: '对比组 1', snapshots: [], highlightDiff: false }],

  // Multi-project split preview: track active dataset per project column
  // { [projectId]: datasetId }
  projectActiveDsIds: {},

  // Self-select preview: dataset ids explicitly HIDDEN from the result preview.
  // Empty set = show all. Using a "hidden" set means newly exported datasets
  // are shown by default.
  previewHiddenIds: new Set(),
};

const els = {
  templateInput: document.getElementById('templateInput'),
  fileInput: document.getElementById('fileInput'),
  fileInputNew: document.getElementById('fileInputNew'),
  projectNameInput: document.getElementById('projectNameInput'),
  projectSelect: document.getElementById('projectSelect'),
  projectSelectHint: document.getElementById('projectSelectHint'),
  downloadTemplateBtn: document.getElementById('downloadTemplateBtn'),
  mainStatus: document.getElementById('mainStatus'),
  activeTemplateChip: document.getElementById('activeTemplateChip'),
  activeProject: document.getElementById('activeProject'),
  fileName: document.getElementById('fileName'),
  exportBtn: document.getElementById('exportBtn'),
  previewBtn: document.getElementById('previewBtn'),
  downloadBtn: document.getElementById('downloadBtn'),
  downloadAllBtn: document.getElementById('downloadAllBtn'),
  resultActions: document.getElementById('resultActions'),
  downloadTip: document.getElementById('downloadTip'),
  tableHead: document.getElementById('tableHead'),
  tableBody: document.getElementById('tableBody'),
  processedHead: document.getElementById('processedHead'),
  processedBody: document.getElementById('processedBody'),
  processedSingleWrap: document.getElementById('processedSingleWrap'),
  processedSplitContainer: document.getElementById('processedSplitContainer'),
  datasetSelectPanel: document.getElementById('datasetSelectPanel'),
  processedSection: document.getElementById('processedSection'),
  csvPreviewSection: document.getElementById('csvPreviewSection'),
  previewStatus: document.getElementById('previewStatus'),
  processedStatus: document.getElementById('processedStatus'),
  datasetTabs: document.getElementById('datasetTabs'),
  templateList: document.getElementById('templateList'),
  templateListView: document.getElementById('templateListView'),
  templateDetailView: document.getElementById('templateDetailView'),
  templateDetailBack: document.getElementById('templateDetailBack'),
  templateDetailBody: document.getElementById('templateDetailBody'),
  templateUploadView: document.getElementById('templateUploadView'),
  templateUploadBack: document.getElementById('templateUploadBack'),
  templateNewBtn: document.getElementById('templateNewBtn'),
  templatePaginationBar: document.getElementById('templatePaginationBar'),
  templatePageInfo: document.getElementById('templatePageInfo'),
  templatePrevBtn: document.getElementById('templatePrevBtn'),
  templateNextBtn: document.getElementById('templateNextBtn'),
  // Admin
  adminModal: document.getElementById('adminModal'),
  templateRenameModal: document.getElementById('templateRenameModal'),
  templateRenameInput: document.getElementById('templateRenameInput'),
  templateRenameDesc: document.getElementById('templateRenameDesc'),
  templateRenameError: document.getElementById('templateRenameError'),
  templateRenameConfirmBtn: document.getElementById('templateRenameConfirmBtn'),
  templateRenameCancelBtn: document.getElementById('templateRenameCancelBtn'),
  adminPasswordInput: document.getElementById('adminPasswordInput'),
  adminLoginError: document.getElementById('adminLoginError'),
  adminLoginBtn: document.getElementById('adminLoginBtn'),
  adminCancelBtn: document.getElementById('adminCancelBtn'),
  adminBadge: document.getElementById('adminBadge'),
  adminToggleBtn: document.getElementById('adminToggleBtn'),
  // History
  historyList: document.getElementById('historyList'),
  clearHistoryBtn: document.getElementById('clearHistoryBtn'),
  clearHistoryCacheBtn: document.getElementById('clearHistoryCacheBtn'),
  // CSV upload display
  csvUploadArea: document.getElementById('csvUploadArea'),
  csvDropzone: document.getElementById('csvDropzone'),
  csvFileInfo: document.getElementById('csvFileInfo'),
  csvFileNameDisplay: document.getElementById('csvFileNameDisplay'),
  csvFileMetaDisplay: document.getElementById('csvFileMetaDisplay'),
  reuploadCsvBtn: document.getElementById('reuploadCsvBtn'),
  // CSV controls
  csvColSearch: document.getElementById('csvColSearch'),
  csvColSearchBtn: document.getElementById('csvColSearchBtn'),
  csvColResetBtn: document.getElementById('csvColResetBtn'),
  csvMatchInfo: document.getElementById('csvMatchInfo'),
  csvColPickerWrap: document.getElementById('csvColPickerWrap'),
  csvColCheckList: document.getElementById('csvColCheckList'),
  chartType: document.getElementById('chartType'),
  generateChartBtn: document.getElementById('generateChartBtn'),
  clearChartBtn: document.getElementById('clearChartBtn'),
  chartContainer: document.getElementById('chartContainer'),
  chartTitle: document.getElementById('chartTitle'),
  chartMeta: document.getElementById('chartMeta'),
  chartResetBtn: document.getElementById('chartResetBtn'),
  csvChart: document.getElementById('csvChart'),
  chartMetricSelector: document.getElementById('chartMetricSelector'),
  // Processed controls
  processedColRange: document.getElementById('processedColRange'),
  processedApplyFilter: document.getElementById('processedApplyFilter'),
  processedShowAllBtn: document.getElementById('processedShowAllBtn'),
  processedRestoreDefaultBtn: document.getElementById('processedRestoreDefaultBtn'),
  processedSetDefaultBtn: document.getElementById('processedSetDefaultBtn'),
  processedAddCompareBtn: document.getElementById('processedAddCompareBtn'),
  compareAllBtn: document.getElementById('compareAllBtn'),
  // Compare
  compareArea: document.getElementById('compareArea'),
  compareGroupsContainer: document.getElementById('compareGroupsContainer'),
  compareGroupsChip: document.getElementById('compareGroupsChip'),
  addCompareGroupBtn: document.getElementById('addCompareGroupBtn'),
  // Header tooltip
  headerTooltip: document.getElementById('headerTooltip'),
};

function syncPreviewSectionVisibility() {
  const hasCsv = state.csvRows.length > 0;
  const hasProcessed = state.processedRows.length > 0;
  const hasCompare = state.datasets.some((d) => d.processedRows && d.processedRows.length)
    || state.compareGroups.some((g) => g.snapshots.length > 0);
  const showDatasetTabs = state.datasets.length > 1;

  if (els.csvPreviewSection) els.csvPreviewSection.classList.toggle('hidden', !hasCsv);
  if (els.processedSection) els.processedSection.classList.toggle('hidden', !hasProcessed);
  if (els.compareArea) els.compareArea.classList.toggle('hidden', !hasCompare);
  if (els.datasetTabs) els.datasetTabs.classList.toggle('hidden', !showDatasetTabs);
}

function resetWorkspaceState() {
  clearChart();
  state.datasets = [];
  state.activeDatasetId = null;
  state.csvRows = [];
  state.processedRows = [];
  state.processedMerges = [];
  state.fileName = '';
  state.workbookBlob = null;
  state.activeProject = null;
  state.csvSearchKeyword = '';
  state.csvSearchMode = 'manual';
  state.csvMatchedCols = [];
  state.csvChartSelectedCols = [];
  state.csvHeaderRowIndex = -1;
  state.csvDataStartRow = -1;
  state.processedColIndices = null;
  state.compareCommonColIndices = null;
  state.compareCommonColIndicesByProject = {};
  Object.values(state.crossProjectChartInstances || {}).forEach((chart) => { try { chart.destroy(); } catch (e) { /* noop */ } });
  state.crossProjectChannelSelection = {};
  state.crossProjectChartInstances = {};
  resetCompareGroups();

  if (els.tableHead) els.tableHead.innerHTML = '';
  if (els.tableBody) els.tableBody.innerHTML = '';
  if (els.processedHead) els.processedHead.innerHTML = '';
  if (els.processedBody) els.processedBody.innerHTML = '';
  if (els.compareGroupsContainer) els.compareGroupsContainer.innerHTML = '';
  if (els.compareGroupsChip) els.compareGroupsChip.textContent = '';
  if (els.datasetTabs) els.datasetTabs.innerHTML = '';
  if (els.previewStatus) els.previewStatus.textContent = '暂无数据';
  if (els.processedStatus) els.processedStatus.textContent = '尚未生成';
  if (els.fileName) els.fileName.textContent = '未上传';
  if (els.activeProject) els.activeProject.textContent = '未选择';
  if (els.mainStatus) els.mainStatus.textContent = '未选模板';
  if (els.chartTitle) els.chartTitle.textContent = '数据图表';
  if (els.chartMeta) els.chartMeta.textContent = '';
  if (els.csvFileNameDisplay) els.csvFileNameDisplay.textContent = '未上传';
  if (els.csvFileMetaDisplay) els.csvFileMetaDisplay.textContent = '等待上传';
  if (els.csvColSearch) els.csvColSearch.value = '';
  if (els.csvColCheckList) els.csvColCheckList.innerHTML = '';
  if (els.csvColPickerWrap) els.csvColPickerWrap.classList.add('hidden');
  if (els.processedColRange) els.processedColRange.value = getDefaultProcessedColumns();
  if (els.fileInput) els.fileInput.value = '';
  if (els.fileInputNew) els.fileInputNew.value = '';

  renderPreview();
  updateMatchInfoAndPicker();
  updateCsvFileDisplay();
  updateSummary();
  syncPreviewSectionVisibility();
}

function getDefaultProcessedColumns() {
  const saved = localStorage.getItem(DEFAULT_COLUMNS_STORAGE_KEY);
  // 兼容旧版本默认值：之前误设为 A-D,M-P，这里迁移到用户最新要求的 A-D,L-O
  if (!saved || !saved.trim()) return FALLBACK_DEFAULT_COLUMNS;
  const trimmed = saved.trim();
  return trimmed === 'A-D,M-P' ? FALLBACK_DEFAULT_COLUMNS : trimmed;
}

function applyProcessedColumnExpression(expression) {
  state.processedColIndices = parseColRange(expression);
  renderProcessedPreview();
}

function restoreDefaultProcessedColumns() {
  // Prefer auto-detection from current data; fall back to saved/FALLBACK
  let defaultColumns = getDefaultProcessedColumns();
  const autoDetected = autoDetectDefaultColumns(state.processedRows);
  if (autoDetected && autoDetected.length) {
    defaultColumns = autoDetected.map((i) => indexToColLetter(i)).join(',');
  }
  els.processedColRange.value = defaultColumns;
  /* Also reset all column widths to auto (undo any manual drag adjustments) */
  resetAllColumnWidths();
  applyProcessedColumnExpression(defaultColumns);
}

/* Reset all tables back to auto column widths (undo freeze + manual drag) */
function resetAllColumnWidths() {
  document.querySelectorAll('table.resizable-table').forEach((table) => {
    delete table.dataset.frozen;
    table.style.removeProperty('min-width');
    table.style.removeProperty('width');
    table.querySelectorAll('th, td').forEach((cell) => {
      cell.style.removeProperty('width');
      cell.style.removeProperty('min-width');
      cell.style.removeProperty('max-width');
    });
    table.querySelectorAll('colgroup col').forEach((col) => {
      col.style.removeProperty('width');
    });
  });
}

function saveDefaultProcessedColumns() {
  const value = (els.processedColRange.value || '').trim() || FALLBACK_DEFAULT_COLUMNS;
  localStorage.setItem(DEFAULT_COLUMNS_STORAGE_KEY, value);
  els.processedColRange.value = value;
  applyProcessedColumnExpression(value);
  alert(`已更新默认展示列为：${value}`);
}

// ---------- Column range parser ----------
function registerChartZoomPlugin() {
  if (typeof Chart === 'undefined') return;
  const zoomPlugin = window.ChartZoom || window.chartjsPluginZoom;
  if (!zoomPlugin) {
    console.warn('chartjs-plugin-zoom is not available');
    return;
  }
  try {
    Chart.register(zoomPlugin);
  } catch {
    // 重复注册时直接忽略，保证刷新/重载场景稳定。
  }
}

function colLetterToIndex(letter) {
  letter = letter.toUpperCase();
  let idx = 0;
  for (let i = 0; i < letter.length; i++) {
    idx = idx * 26 + (letter.charCodeAt(i) - 64);
  }
  return idx - 1;
}
function indexToColLetter(idx) {
  let letter = '';
  idx += 1;
  while (idx > 0) {
    const rem = (idx - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    idx = Math.floor((idx - 1) / 26);
  }
  return letter;
}
function parseColRange(rangeStr) {
  if (!rangeStr || !rangeStr.trim()) return null;
  const parts = rangeStr.trim().split(/[,，;；\s]+/);
  const indices = [];
  for (const part of parts) {
    if (!part) continue;
    const rangeMatch = part.match(/^([A-Za-z]+)-([A-Za-z]+)$/) || part.match(/^(\d+)-(\d+)$/);
    if (rangeMatch) {
      let start, end;
      if (/^[A-Za-z]/.test(rangeMatch[1])) {
        start = colLetterToIndex(rangeMatch[1]);
        end = colLetterToIndex(rangeMatch[2]);
      } else {
        start = parseInt(rangeMatch[1], 10) - 1;
        end = parseInt(rangeMatch[2], 10) - 1;
      }
      if (start > end) [start, end] = [end, start];
      for (let i = start; i <= end; i++) {
        if (!indices.includes(i)) indices.push(i);
      }
    } else if (/^[A-Za-z]+$/.test(part)) {
      const idx = colLetterToIndex(part);
      if (!indices.includes(idx)) indices.push(idx);
    } else if (/^\d+$/.test(part)) {
      const idx = parseInt(part, 10) - 1;
      if (!indices.includes(idx) && idx >= 0) indices.push(idx);
    }
  }
  return indices.length > 0 ? indices.sort((a, b) => a - b) : null;
}

// ---------- IndexedDB helpers ----------
function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: 'id' });
      }
      if (!db.objectStoreNames.contains(HISTORY_STORE_NAME)) {
        db.createObjectStore(HISTORY_STORE_NAME, { keyPath: 'id' });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}
function runStoreRequest(store, method, ...args) {
  return new Promise((resolve, reject) => {
    const req = store[method](...args);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function dbGetAll() {
  const db = await openDb();
  const tx = db.transaction(STORE_NAME, 'readonly');
  return runStoreRequest(tx.objectStore(STORE_NAME), 'getAll');
}
async function dbGet(id) {
  const db = await openDb();
  const tx = db.transaction(STORE_NAME, 'readonly');
  return runStoreRequest(tx.objectStore(STORE_NAME), 'get', id).then((result) => result || null);
}
async function dbPut(record) {
  const db = await openDb();
  const tx = db.transaction(STORE_NAME, 'readwrite');
  return runStoreRequest(tx.objectStore(STORE_NAME), 'put', record);
}
async function dbDelete(id) {
  const db = await openDb();
  const tx = db.transaction(STORE_NAME, 'readwrite');
  await runStoreRequest(tx.objectStore(STORE_NAME), 'delete', id);
}
async function openHistoryDB() {
  return openDb();
}
async function getAllHistoryEntries() {
  const db = await openHistoryDB();
  const tx = db.transaction(HISTORY_STORE_NAME, 'readonly');
  const items = await runStoreRequest(tx.objectStore(HISTORY_STORE_NAME), 'getAll');
  return Array.isArray(items) ? items : [];
}
async function getHistoryEntry(id) {
  if (!id) return null;
  const db = await openHistoryDB();
  const tx = db.transaction(HISTORY_STORE_NAME, 'readonly');
  const item = await runStoreRequest(tx.objectStore(HISTORY_STORE_NAME), 'get', id);
  return item || null;
}
async function deleteHistoryEntryCache(id) {
  if (!id) return;
  const db = await openHistoryDB();
  const tx = db.transaction(HISTORY_STORE_NAME, 'readwrite');
  await runStoreRequest(tx.objectStore(HISTORY_STORE_NAME), 'delete', id);
}
async function clearAllHistoryEntries() {
  const db = await openHistoryDB();
  const tx = db.transaction(HISTORY_STORE_NAME, 'readwrite');
  await runStoreRequest(tx.objectStore(HISTORY_STORE_NAME), 'clear');
}
async function trimHistoryStore(limit = HISTORY_CACHE_MAX) {
  const items = await getAllHistoryEntries();
  if (items.length <= limit) return;
  items.sort((a, b) => (a.createdAt || 0) - (b.createdAt || 0));
  const extra = items.length - limit;
  for (let i = 0; i < extra; i += 1) {
    await deleteHistoryEntryCache(items[i].id);
  }
}
async function saveHistoryEntry(entry) {
  const normalized = {
    ...entry,
    datasets: Array.isArray(entry.datasets) ? entry.datasets.slice(0, BATCH_MAX) : [],
  };
  const db = await openHistoryDB();
  const tx = db.transaction(HISTORY_STORE_NAME, 'readwrite');
  await runStoreRequest(tx.objectStore(HISTORY_STORE_NAME), 'put', normalized);
  await trimHistoryStore(HISTORY_CACHE_MAX);
  return normalized.id;
}

// ---------- Utilities ----------
function normalizeCellValue(value) {
  if (value === undefined || value === null) return '';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return value;
  if (value instanceof Date) return value;
  // xlsx-populate RichText: cell.value() may return a RichText object.
  if (typeof value.text === 'function') return value.text();
  if (typeof value.text === 'string') return value.text;
  try { return String(value); } catch { return ''; }
}
function normalizeCell(value) {
  const normalized = normalizeCellValue(value);
  return normalized instanceof Date ? normalized.toISOString() : String(normalized);
}
function parseCsvText(text) {
  const result = Papa.parse(text, { skipEmptyLines: false });
  return result.data.map((row) => row.map(normalizeCell));
}
function getRowCount(rows) { return rows.length; }
function getColCount(rows) { return rows.reduce((max, row) => Math.max(max, row.length), 0); }

function formatSize(bytes) {
  if (!bytes && bytes !== 0) return '-';
  if (bytes >= 1024 * 1024) return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
  if (bytes >= 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${bytes} B`;
}
function formatTime(ts) {
  if (!ts) return '-';
  const d = new Date(ts);
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}
function formatDateTime(ts) {
  if (!ts) return '-';
  const d = new Date(ts);
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}
function estimateJsonSize(value) {
  try {
    return new Blob([JSON.stringify(value)]).size;
  } catch {
    return Number.POSITIVE_INFINITY;
  }
}
function scrollToEl(el, block = 'start') {
  if (el && typeof el.scrollIntoView === 'function') {
    el.scrollIntoView({ behavior: 'smooth', block });
  }
}

function escapeHtml(str) {
  return String(str === null || str === undefined ? '' : str).replace(/[&<>"']/g, (ch) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
  }[ch]));
}

// ---------- Number helpers ----------
function isNumericLike(v) {
  if (typeof v === 'number') return Number.isFinite(v);
  if (v === null || v === undefined) return false;
  const s = String(v).trim();
  if (s === '') return false;
  return NUMERIC_RE.test(s);
}
function toNumberOrNull(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  if (s === '' || !NUMERIC_RE.test(s)) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}
function formatNumeric2(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return Number.isFinite(v) ? v.toFixed(2) : '';
  const s = String(v);
  const trimmed = s.trim();
  if (trimmed === '') return '';
  if (NUMERIC_RE.test(trimmed)) {
    const n = Number(trimmed);
    if (Number.isFinite(n)) return n.toFixed(2);
  }
  return s;
}

function computeNumericColumnMap(rows, colIndices) {
  const map = new Map();
  for (const ci of colIndices) {
    let numericCount = 0;
    let nonEmpty = 0;
    for (const row of rows) {
      const v = ci < row.length ? row[ci] : '';
      if (v === '' || v === null || v === undefined) continue;
      nonEmpty += 1;
      if (isNumericLike(v)) numericCount += 1;
    }
    map.set(ci, nonEmpty >= 2 && numericCount / nonEmpty >= 0.7);
  }
  return map;
}

// ---------- Operation history ----------
function getHistory() {
  return Array.isArray(state.historyItems) ? state.historyItems : [];
}

async function loadHistoryFromCloud() {
  const client = await getSupabaseClient();
  const { data, error } = await client.from('history').select('*').order('created_at', { ascending: false }).limit(HISTORY_MAX);
  if (error) throw error;
  state.historyItems = (data || []).map((item) => ({
    id: item.id,
    time: Date.parse(item.created_at),
    action: '结果导出',
    fileName: item.filename || '',
    projectName: item.template_name || '',
    rowCount: Number(item.rows || 0),
    colCount: Number(item.cols || 0),
    datasetCount: 0,
    historyStoreId: item.id,
    historyCacheSummary: '云端已保存操作记录；本机保留可回溯结果缓存。',
  }));
  renderHistory();
}

async function addHistoryEntry(entry) {
  const client = await getSupabaseClient();
  const user = window.powerAuth.getUser();
  const payload = {
    id: entry.historyStoreId || entry.id || crypto.randomUUID(),
    user_id: user.id,
    filename: entry.fileName || '',
    template_name: entry.projectName || '',
    rows: Number(entry.rowCount || 0),
    cols: Number(entry.colCount || 0),
  };
  const { data, error } = await client.from('history').insert(payload).select().single();
  if (error) throw error;
  const record = {
    id: data.id,
    time: Date.parse(data.created_at),
    action: entry.action || '导出',
    fileName: data.filename,
    projectName: data.template_name,
    rowCount: data.rows,
    colCount: data.cols,
    datasetCount: entry.datasetCount || 0,
    historyStoreId: data.id,
    historyCacheSummary: entry.historyCacheSummary || '云端已保存操作记录。',
  };
  state.historyItems.unshift(record);
  if (state.historyItems.length > HISTORY_MAX) state.historyItems.length = HISTORY_MAX;
  renderHistory();
  return record;
}

async function deleteHistoryEntry(id) {
  const item = getHistory().find((record) => record.id === id);
  const client = await getSupabaseClient();
  const { error } = await client.from('history').delete().eq('id', id);
  if (error) throw error;
  state.historyItems = getHistory().filter((record) => record.id !== id);
  if (item && item.historyStoreId) {
    try { await deleteHistoryEntryCache(item.historyStoreId); }
    catch (error) { console.warn('deleteHistoryEntryCache failed', error); }
  }
  renderHistory();
}

async function clearHistory() {
  if (!getHistory().length) return;
  if (!confirm('确定要清空全部历史操作记录和缓存结果吗？该操作不可恢复。')) return;
  const client = await getSupabaseClient();
  const user = window.powerAuth.getUser();
  const { error } = await client.from('history').delete().eq('user_id', user.id);
  if (error) throw error;
  state.historyItems = [];
  try {
    await clearAllHistoryEntries();
  } catch (error) {
    console.warn('clearAllHistoryEntries failed', error);
  }
  renderHistory();
}
function renderHistory() {
  const list = getHistory();
  [els.clearHistoryBtn, els.clearHistoryCacheBtn].forEach((btn) => {
    if (!btn) return;
    btn.disabled = list.length === 0;
    btn.classList.toggle('opacity-50', list.length === 0);
    btn.classList.toggle('cursor-not-allowed', list.length === 0);
  });
  if (!list.length) {
    els.historyList.innerHTML = `
      <div class="empty-state rounded-2xl p-4 text-slate-500 bg-white border border-dashed border-slate-200 text-center">
        暂无历史操作记录
      </div>`;
    return;
  }
  els.historyList.innerHTML = list.map((item) => {
    const hasCache = !!item.historyStoreId;
    const cacheSummary = item.historyCacheSummary
      ? `<div class="mt-1 text-xs text-slate-400">${escapeHtml(item.historyCacheSummary)}</div>`
      : '';
    return `
    <div class="history-item rounded-2xl p-4 bg-white border border-slate-200" data-id="${item.id}">
      <div class="flex flex-col md:flex-row md:items-center md:justify-between gap-3">
        <div class="min-w-0">
          <div class="flex items-center gap-2 flex-wrap">
            <span class="material-symbols-outlined text-indigo-600">history</span>
            <div class="font-bold text-slate-900 truncate">${escapeHtml(item.fileName || '未命名 CSV')}</div>
            <span class="text-xs font-semibold px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700 border border-indigo-100">${escapeHtml(item.action || '导出')}</span>
            ${hasCache ? '<span class="text-[11px] font-semibold px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-100">已缓存结果</span>' : '<span class="text-[11px] font-semibold px-2 py-0.5 rounded-full bg-slate-100 text-slate-500 border border-slate-200">仅留痕</span>'}
          </div>
          <div class="mt-1 text-sm text-slate-500">
            ${formatDateTime(item.time)} · 模板：${escapeHtml(item.projectName || '-')} · ${Number(item.rowCount || 0).toLocaleString('zh-CN')} 行 × ${Number(item.colCount || 0).toLocaleString('zh-CN')} 列${item.datasetCount ? ` · ${item.datasetCount} 组数据` : ''}
          </div>
          ${cacheSummary}
        </div>
        <div class="flex gap-2 flex-wrap justify-end">
          <button class="template-btn template-btn-secondary" data-action="restore-history" data-id="${item.id}">回溯</button>
          <button class="template-btn template-btn-primary" data-action="download-history" data-id="${item.id}">下载</button>
          <button class="template-btn template-btn-danger" data-action="delete-history" data-id="${item.id}">删除</button>
        </div>
      </div>
    </div>`;
  }).join('');
}

function slugifyProjectName(name) { return `p_${name.replace(/\s+/g, '_')}`; }

// ---------- Dataset model ----------
function genDatasetId() { return `ds_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`; }
function getActiveDataset() { return state.datasets.find((d) => d.id === state.activeDatasetId) || null; }
function syncActiveDatasetToState() {
  const ds = getActiveDataset();
  if (!ds) {
    state.csvRows = [];
    state.fileName = '';
    state.csvHeaderRowIndex = -1;
    state.csvDataStartRow = -1;
    state.processedRows = [];
    state.processedMerges = [];
    state.workbookBlob = null;
    return;
  }
  state.csvRows = ds.csvRows;
  state.fileName = ds.fileName;
  state.csvHeaderRowIndex = ds.csvHeaderRowIndex;
  state.csvDataStartRow = ds.csvDataStartRow;
  state.processedRows = ds.processedRows;
  state.processedMerges = ds.processedMerges;
  state.workbookBlob = ds.workbookBlob;
}
function setActiveDataset(id) {
  state.activeDatasetId = id;
  syncActiveDatasetToState();
}

function hasReadyInput() {
  // A dataset is exportable when it has CSV rows AND a resolvable project template
  // — either its own per-file project or the global default active project.
  return state.datasets.some((d) => d.csvRows && d.csvRows.length > 0 && (d.projectId || state.activeProject));
}
function hasAnyProcessed() {
  return state.datasets.some((d) => d.processedRows && d.processedRows.length > 0);
}

// ---------- CSV header detection ----------
function detectHeaderRow(rows) {
  const limit = Math.min(rows.length, 20);
  let firstDataRow = -1;
  for (let i = 0; i < limit; i++) {
    const cell0 = (rows[i][0] || '').trim();
    if (cell0 !== '' && NUMERIC_RE.test(cell0)) { firstDataRow = i; break; }
  }
  if (firstDataRow < 0) return { headerRowIndex: 0, dataStartRow: 1 };

  let hdr = -1;
  for (let i = firstDataRow - 1; i >= 0; i--) {
    const cell0 = (rows[i][0] || '').trim();
    if (cell0 !== '' && !NUMERIC_RE.test(cell0)) { hdr = i; break; }
  }
  if (hdr < 0) {
    for (let i = firstDataRow - 1; i >= 0; i--) {
      const hasContent = rows[i].some((c) => (c || '').trim() !== '');
      if (hasContent) { hdr = i; break; }
    }
  }
  if (hdr < 0) hdr = 0;
  return { headerRowIndex: hdr, dataStartRow: firstDataRow };
}

// ---------- State refresh ----------
function updateSummary() {
  const rowCount = getRowCount(state.csvRows);
  const hasProcessed = state.processedRows.length > 0;
  const ready = hasReadyInput();
  const anyProcessed = hasAnyProcessed();
  const nDatasets = state.datasets.length;

  els.activeProject.textContent = state.activeProject ? state.activeProject.name : '未选择';
  els.activeTemplateChip.textContent = state.activeProject ? `当前模板：${state.activeProject.name}` : '未选择项目模板';
  els.fileName.textContent = nDatasets > 1 ? `${nDatasets} 个 CSV 文件` : (state.fileName || '未上传');
  els.previewStatus.textContent = rowCount ? `已载入 ${rowCount} 行` : '暂无数据';
  if (rowCount && nDatasets > 1) {
    const options = state.datasets.map((d) =>
      `<option value="${d.id}" ${d.id === state.activeDatasetId ? 'selected' : ''}>${escapeHtml(d.fileName)}</option>`
    ).join('');
    els.previewStatus.innerHTML = `<span class="current-ds-hint">当前预览：<select class="ds-switch-select" data-target="csv">${options}</select></span> · 已载入 ${rowCount} 行`;
  }
  if (hasProcessed && nDatasets > 1) {
    // Detect multi-project scenario for a more accurate status message
    const processedDs = state.datasets.filter((d) => d.processedRows && d.processedRows.length);
    const distinctPids = new Set(processedDs.map((d) => d.projectId || '__unassigned__'));
    if (distinctPids.size >= 2) {
      els.processedStatus.innerHTML = `<span class="current-ds-hint">已按项目分栏展示（${distinctPids.size} 个项目 · ${processedDs.length} 组数据）</span>`;
    } else {
      const options = state.datasets.map((d) =>
        `<option value="${d.id}" ${d.id === state.activeDatasetId ? 'selected' : ''}>${escapeHtml(d.fileName)}</option>`
      ).join('');
      els.processedStatus.innerHTML = `<span class="current-ds-hint">当前预览：<select class="ds-switch-select" data-target="processed">${options}</select></span> · 已加载 ${state.processedRows.length} 行`;
    }
  } else {
    els.processedStatus.textContent = hasProcessed ? `已加载 ${state.processedRows.length} 行` : '尚未生成';
  }

  // A project is considered selected if there's a global default OR any file has
  // its own per-file project assigned.
  const hasAnyProject = !!state.activeProject || state.datasets.some((d) => d.projectId);

  els.mainStatus.textContent = !hasAnyProject
    ? '未选模板'
    : !ready
      ? '待上传 CSV'
      : anyProcessed
        ? '已生成结果'
        : '就绪';

  els.exportBtn.disabled = !ready;
  if (els.processedAddCompareBtn) els.processedAddCompareBtn.disabled = !hasProcessed;
  if (els.compareAllBtn) {
    els.compareAllBtn.disabled = !anyProcessed || nDatasets < 2;
    els.compareAllBtn.classList.toggle('hidden', nDatasets < 2);
  }
  if (els.downloadAllBtn) {
    els.downloadAllBtn.classList.toggle('hidden', !(anyProcessed && nDatasets > 1));
  }

  // Detect distinct projects across uploaded datasets (for multi-project scenarios)
  const distinctProjectIds = new Set(
    state.datasets.filter((d) => d.projectId).map((d) => d.projectId)
  );
  const multiProject = distinctProjectIds.size > 1;
  const missingProject = state.datasets.some((d) => d.csvRows && d.csvRows.length && !d.projectId);

  if (!hasAnyProject) {
    els.downloadTip.textContent = '请先选择或上传一个项目模板（也可为每个文件单独选择）。';
  } else if (!ready) {
    const tipName = state.activeProject ? state.activeProject.name : '所选项目';
    els.downloadTip.textContent = `已选择项目模板：${tipName}。请继续上传 CSV（支持一次批量上传多个）。`;
  } else if (anyProcessed) {
    els.downloadTip.textContent = nDatasets > 1
      ? (multiProject
          ? `已生成 ${nDatasets} 组结果（涉及 ${distinctProjectIds.size} 个不同项目）。可切换上方数据组预览，或用"下载全部结果"批量下载。`
          : `已生成 ${nDatasets} 组结果。可切换上方数据组预览，或用"下载全部结果"批量下载。`)
      : '结果已生成。可点击下方"预览结果"或"下载 XLSX"。';
  } else if (missingProject) {
    els.downloadTip.textContent = '部分文件尚未指定项目模板，请在上方文件列表为每个文件选择项目后再导出。';
  } else {
    els.downloadTip.textContent = nDatasets > 1
      ? (multiProject
          ? `共 ${nDatasets} 组 CSV，涉及 ${distinctProjectIds.size} 个不同项目模板，点击"结果导出"按对应模板批量生成。`
          : `模板已选，共 ${nDatasets} 组 CSV 待处理，点击"结果导出"批量生成。`)
      : '模板和 CSV 已就绪，点击"结果导出"开始处理。';
  }

  if (anyProcessed) els.resultActions.classList.remove('hidden');
  else els.resultActions.classList.add('hidden');

  renderDatasetTabs();
  updateCsvFileDisplay();
  syncProjectSelect();
  syncPreviewSectionVisibility();
}

// Builds <option> list for a per-file project dropdown, marking the dataset's
// current project as selected.
function buildDatasetProjectOptions(selectedId) {
  const opts = ['<option value="">— 默认（跟随全局项目）—</option>'];
  for (const it of state.templateItems) {
    const sel = it.id === selectedId ? ' selected' : '';
    opts.push(`<option value="${escapeHtml(it.id)}"${sel}>${escapeHtml(it.name)}</option>`);
  }
  return opts.join('');
}

// Renders one uploaded-file line with its own project selector.
function renderCsvFileLine(d, i, opts = {}) {
  const { active = false, showIdx = true } = opts;
  const color = DS_COLORS[i % DS_COLORS.length];
  const idxBadge = showIdx
    ? `<span class="csv-file-idx" style="background:${color}">#${i + 1}</span>`
    : '';
  return `<div class="csv-file-line ${active ? 'csv-file-line-active' : ''}">
      ${idxBadge}
      <span class="csv-file-name" title="${escapeHtml(d.fileName)}">${escapeHtml(d.fileName)}</span>
      <select class="csv-file-project-select" data-action="ds-project" data-ds-id="${escapeHtml(d.id)}" title="为该文件选择项目模板">
        ${buildDatasetProjectOptions(d.projectId || '')}
      </select>
    </div>`;
}

function updateCsvFileDisplay() {
  const ds = getActiveDataset();
  const n = state.datasets.length;
  if (n > 0 && ds) {
    els.csvDropzone.classList.add('hidden');
    els.csvFileInfo.classList.remove('hidden');
    if (n > 1) {
      els.csvFileNameDisplay.style.whiteSpace = 'normal';
      const activeIdx = state.datasets.findIndex((d) => d.id === state.activeDatasetId);
      els.csvFileNameDisplay.innerHTML = state.datasets.map((d, i) =>
        renderCsvFileLine(d, i, { active: d.id === state.activeDatasetId, showIdx: true })
      ).join('');
      const rowCount = getRowCount(ds.csvRows);
      const colCount = getColCount(ds.csvRows);
      els.csvFileMetaDisplay.textContent = `共 ${n} 组数据 · 当前预览 #${activeIdx + 1}：${rowCount} 行 × ${colCount} 列 · 可为每个文件单独选择项目模板`;
    } else {
      els.csvFileNameDisplay.style.whiteSpace = 'normal';
      els.csvFileNameDisplay.innerHTML = renderCsvFileLine(ds, 0, { active: false, showIdx: false });
      const rowCount = getRowCount(ds.csvRows);
      const colCount = getColCount(ds.csvRows);
      els.csvFileMetaDisplay.textContent = `${rowCount} 行 × ${colCount} 列 · 可为该文件单独选择项目模板`;
    }
  } else {
    els.csvDropzone.classList.remove('hidden');
    els.csvFileInfo.classList.add('hidden');
  }
}

function syncProjectSelect() {
  if (!els.projectSelect) return;
  const activeId = state.activeProject ? state.activeProject.id : '';
  els.projectSelect.value = activeId || '';
  // Enable/disable template download button
  if (els.downloadTemplateBtn) {
    els.downloadTemplateBtn.disabled = !activeId;
  }
}

// ---------- Dataset tabs (multi-group preview) ----------
function renderDatasetTabs() {
  if (!els.datasetTabs) return;
  const ds = state.datasets;
  if (ds.length <= 1) {
    els.datasetTabs.classList.add('hidden');
    els.datasetTabs.innerHTML = '';
    return;
  }
  els.datasetTabs.classList.remove('hidden');
  const chips = ds.map((d, i) => {
    const active = d.id === state.activeDatasetId;
    const done = !!(d.processedRows && d.processedRows.length);
    const color = DS_COLORS[i % DS_COLORS.length];
    return `<div class="ds-tab ${active ? 'ds-tab-active' : ''}" data-action="pick-dataset" data-id="${d.id}" title="${escapeHtml(d.fileName)}">
      <span class="ds-dot" style="background:${done ? color : '#cbd5e1'}"></span>
      <span class="ds-name">${escapeHtml(d.fileName)}</span>
      ${done
        ? `<button class="ds-dl" data-action="dl-dataset" data-id="${d.id}" title="下载此组结果"><span class="material-symbols-outlined" style="font-size:14px">download</span></button>`
        : '<span class="ds-pending">待导出</span>'}
    </div>`;
  }).join('');
  els.datasetTabs.innerHTML = `<span class="ds-tabs-label">共 ${ds.length} 组数据：</span>${chips}<span class="ds-tabs-tip">💡 点击切换预览数据组</span>`;
}

// ---------- CSV preview: column search + rendering ----------
function performColumnSearch() {
  const kw = (els.csvColSearch.value || '').trim();
  state.csvSearchKeyword = kw;
  state.csvSearchMode = 'manual';
  if (!state.csvRows.length) {
    state.csvMatchedCols = [];
    updateMatchInfoAndPicker();
    renderCsvPreview();
    return;
  }
  if (state.csvHeaderRowIndex < 0) {
    const d = detectHeaderRow(state.csvRows);
    state.csvHeaderRowIndex = d.headerRowIndex;
    state.csvDataStartRow = d.dataStartRow;
  }
  const header = state.csvRows[state.csvHeaderRowIndex] || [];
  const maxCols = getColCount(state.csvRows);

  if (!kw) {
    state.csvMatchedCols = [];
    updateMatchInfoAndPicker();
    renderCsvPreview();
    return;
  }
  const kwLower = kw.toLowerCase();
  const matched = [];
  for (let c = 0; c < maxCols; c++) {
    const cell = (header[c] || '').toString().toLowerCase();
    if (cell && cell.includes(kwLower)) matched.push(c);
  }
  state.csvMatchedCols = matched;
  state.csvChartSelectedCols = matched.slice(0, 4);
  updateMatchInfoAndPicker();
  renderCsvPreview();
}

function resetColumnSearch() {
  els.csvColSearch.value = '';
  state.csvSearchKeyword = '';
  state.csvSearchMode = 'manual';
  state.csvMatchedCols = [];
  state.csvChartSelectedCols = [];
  clearChart();
  updateMatchInfoAndPicker();
  renderCsvPreview();
}

function updateMatchInfoAndPicker() {
  const kw = state.csvSearchKeyword;
  const matched = state.csvMatchedCols || [];
  if (!kw) {
    els.csvMatchInfo.textContent = '未搜索，默认显示前 40 列（超长列名点击表头查看完整名称）。';
    els.csvColPickerWrap.classList.add('hidden');
    return;
  }
  if (!matched.length) {
    els.csvMatchInfo.textContent = `关键字"${kw}"未匹配到任何列（表头行：${state.csvHeaderRowIndex + 1}）。`;
    els.csvColPickerWrap.classList.add('hidden');
    return;
  }
  const header = state.csvRows[state.csvHeaderRowIndex] || [];
  const selectedSet = new Set(state.csvChartSelectedCols);
  const activeMetric = state.channelMetricType || METRIC_CURRENT;
  const metricCols = state.channelMetricCols && state.channelMetricCols[activeMetric] ? state.channelMetricCols[activeMetric] : [];
  if (state.csvSearchMode === 'channel') {
    if (metricCols.length) {
      els.csvMatchInfo.textContent = `关键字"${kw}"匹配到 ${matched.length} 列；当前指标“${METRIC_LABELS[activeMetric]}”可用 ${metricCols.length} 列，已在图表中按该指标展示。`;
    } else {
      els.csvMatchInfo.textContent = `关键字"${kw}"匹配到 ${matched.length} 列，但当前指标“${METRIC_LABELS[activeMetric]}”未找到，所以不会拿其他列冒充该指标绘图。可切换到其他指标查看。`;
    }
  } else {
    els.csvMatchInfo.textContent = `关键字"${kw}"匹配到 ${matched.length} 列，已在表格中高亮。勾选下方列后可生成折线图/柱状图。`;
  }
  els.csvColPickerWrap.classList.remove('hidden');
  els.csvColCheckList.innerHTML = matched.map((c) => {
    const name = header[c] || `(空)`;
    const label = `${indexToColLetter(c)} · ${name}`;
    const isSel = selectedSet.has(c);
    return `<label class="col-chip ${isSel ? 'selected' : ''}" data-col="${c}" title="${escapeHtml(name)}">
      <input type="checkbox" ${isSel ? 'checked' : ''} data-col="${c}" />
      <span>${escapeHtml(label)}</span>
    </label>`;
  }).join('');
}

function buildTruncateHeaderHtml(colLetter, colIdx, headerText) {
  const label = headerText || '—';
  return `<div class="th-letter">${colLetter}<span class="th-idx">(${colIdx + 1})</span></div>
    <span class="th-truncate" data-full="${escapeHtml(label)}" data-letter="${colLetter}" data-idx="${colIdx + 1}" title="${escapeHtml(label)}">${escapeHtml(label)}</span>`;
}

function isProcessedCenteredMetricCol(ci) {
  return ci === COL_M_VAVG - 1 || ci === COL_N_IAVG - 1 || ci === COL_O_PAVG - 1;
}

function clampColumnWidth(width) {
  return Math.min(500, Math.max(40, width));
}

function setColumnWidth(th, width) {
  const nextWidth = clampColumnWidth(width);
  const table = th.closest('table');
  const colKey = th.dataset.resizeCol;
  if (!table || !colKey) return;
  /* Set width on the <col> element (authoritative for table-layout:fixed) */
  const col = table.querySelector(`colgroup col[data-col="${colKey}"]`);
  if (col) {
    col.style.width = `${nextWidth}px`;
  }
  /* Also sync th so resizer position stays correct */
  th.style.width = `${nextWidth}px`;
}

function updateResizeGuide(table, th, width) {
  const tableRect = table.getBoundingClientRect();
  const thRect = th.getBoundingClientRect();
  const guideLeft = thRect.left - tableRect.left + width;
  table.style.setProperty('--resize-guide-left', `${guideLeft}px`);
}

function measureColumnBestWidth(table, th) {
  const colKey = th.dataset.resizeCol;
  if (!colKey) return clampColumnWidth(th.getBoundingClientRect().width);
  const cells = table.querySelectorAll(`[data-resize-col="${colKey}"]`);
  if (!cells.length) return clampColumnWidth(th.getBoundingClientRect().width);

  const probe = document.createElement('span');
  probe.className = 'col-measure-probe';
  document.body.appendChild(probe);

  let maxWidth = 0;
  cells.forEach((cell) => {
    const text = (cell.textContent || '').trim();
    probe.textContent = text || '—';
    const style = window.getComputedStyle(cell);
    probe.style.font = style.font;
    probe.style.fontWeight = style.fontWeight;
    probe.style.fontVariantNumeric = style.fontVariantNumeric;
    probe.style.letterSpacing = style.letterSpacing;
    maxWidth = Math.max(maxWidth, probe.getBoundingClientRect().width);
  });

  document.body.removeChild(probe);
  const computed = window.getComputedStyle(th);
  const paddingX = (parseFloat(computed.paddingLeft) || 0) + (parseFloat(computed.paddingRight) || 0);
  return clampColumnWidth(Math.ceil(maxWidth + paddingX + 18));
}

function applyStickyHeaderOffsets(table) {
  if (!table) return;
  const rows = Array.from(table.querySelectorAll('thead tr'));
  let top = 0;
  rows.forEach((tr) => {
    const height = tr.getBoundingClientRect().height || tr.offsetHeight || 0;
    tr.querySelectorAll('th').forEach((th) => {
      th.style.top = `${top}px`;
      th.classList.add('sticky-head');
    });
    top += height;
  });
}

function applyStickyColumns(root = document) {
  /* M 列 sticky 已移除（与合并单元格冲突），此函数保留为空壳以避免调用处报错 */
  const scope = root && typeof root.querySelectorAll === 'function' ? root : document;
  scope.querySelectorAll('table.resizable-table').forEach((table) => {
    applyStickyHeaderOffsets(table);
  });
}

function initResizableTables(root = document) {
  const tables = root.querySelectorAll('table.resizable-table');
  tables.forEach((table) => {
    table.style.tableLayout = 'fixed';
    table.dataset.resized = '1';
    delete table.dataset.frozen;

    /* Ensure a <colgroup> exists with one <col> per displayed column.
       Collect ALL non-merged ths (colspan<=1) from the entire thead — they represent individual columns
       regardless of which header row they're in (some have rowspan=2 in the first row). */
    let colgroup = table.querySelector('colgroup');
    if (colgroup) colgroup.remove();
    colgroup = document.createElement('colgroup');
    const thead = table.querySelector('thead');
    const individualThs = thead
      ? [...thead.querySelectorAll('th')].filter(t => (t.colSpan || 1) <= 1)
      : [...table.querySelectorAll('th')].filter(t => (t.colSpan || 1) <= 1);
    individualThs.forEach((th) => {
      const col = document.createElement('col');
      const key = th.dataset.resizeCol || '';
      col.dataset.col = key;
      // Give value columns (Vavg/Iavg/Pavg) a guaranteed minimum width
      if (th.classList.contains('value-col-full')) {
        col.style.width = '100px';
      }
      colgroup.appendChild(col);
    });
    table.prepend(colgroup);

    /* Only add resizers to non-merged ths (colspan <= 1) */
    individualThs.forEach((th) => {
      th.querySelectorAll(':scope > .col-resizer').forEach((old) => old.remove());
      const resizer = document.createElement('div');
      resizer.className = 'col-resizer';
      resizer.title = '拖动调整列宽，双击按内容自动适配';
      resizer.addEventListener('click', (event) => event.stopPropagation());
      resizer.addEventListener('dblclick', (event) => {
        event.preventDefault();
        event.stopPropagation();
        const bestWidth = measureColumnBestWidth(table, th);
        setColumnWidth(th, bestWidth);
      });
      const startResize = (clientX) => {
        /* Freeze all col widths on first drag so other columns stay put */
        if (!table.dataset.frozen) {
          table.dataset.frozen = '1';
          let totalW = 0;
          individualThs.forEach((t) => {
            const w = t.getBoundingClientRect().width;
            totalW += w;
            const k = t.dataset.resizeCol || '';
            const c = table.querySelector(`colgroup col[data-col="${k}"]`);
            if (c) c.style.width = `${w}px`;
            t.style.width = `${w}px`;
          });
          /* Switch table from 100% to exact pixel width so browser won't redistribute */
          table.style.width = `${totalW}px`;
        }
        const startX = clientX;
        const startWidth = th.getBoundingClientRect().width;
        table.classList.add('is-col-resizing');
        th.classList.add('is-resizing');
        document.body.classList.add('col-resizing');
        updateResizeGuide(table, th, startWidth);
        const onMove = (moveEvent) => {
          const moveX = moveEvent.touches ? moveEvent.touches[0].clientX : moveEvent.clientX;
          const width = clampColumnWidth(startWidth + moveX - startX);
          setColumnWidth(th, width);
          /* Update table total width = sum of all frozen col widths */
          let total = 0;
          table.querySelectorAll('colgroup col').forEach((c) => {
            total += parseFloat(c.style.width) || 0;
          });
          if (total > 0) table.style.width = `${total}px`;
          updateResizeGuide(table, th, width);
        };
        const onUp = () => {
          table.classList.remove('is-col-resizing');
          th.classList.remove('is-resizing');
          table.style.removeProperty('--resize-guide-left');
          document.body.classList.remove('col-resizing');
          document.removeEventListener('mousemove', onMove);
          document.removeEventListener('mouseup', onUp);
          document.removeEventListener('touchmove', onMove);
          document.removeEventListener('touchend', onUp);
          document.removeEventListener('touchcancel', onUp);
        };
        document.addEventListener('mousemove', onMove);
        document.addEventListener('mouseup', onUp);
        document.addEventListener('touchmove', onMove, { passive: false });
        document.addEventListener('touchend', onUp);
        document.addEventListener('touchcancel', onUp);
      };
      resizer.addEventListener('mousedown', (event) => {
        event.preventDefault();
        event.stopPropagation();
        startResize(event.clientX);
      });
      resizer.addEventListener('touchstart', (event) => {
        event.preventDefault();
        event.stopPropagation();
        if (event.touches && event.touches.length) startResize(event.touches[0].clientX);
      }, { passive: false });
      th.appendChild(resizer);
    });
    applyStickyHeaderOffsets(table);
  });
}

function renderCsvPreview() {
  const headEl = els.tableHead;
  const bodyEl = els.tableBody;
  headEl.innerHTML = '';
  bodyEl.innerHTML = '';
  if (!state.csvRows.length) return;

  const maxCols = getColCount(state.csvRows);
  const matchedSet = new Set(state.csvMatchedCols || []);

  let colIndices;
  if (state.csvSearchKeyword && matchedSet.size > 0) {
    colIndices = [0, ...state.csvMatchedCols.filter((c) => c !== 0)];
  } else {
    colIndices = [];
    for (let i = 0; i < Math.min(maxCols, 40); i++) colIndices.push(i);
  }

  const preLines = Math.max(5, state.csvDataStartRow + 1);
  const preview = state.csvRows.slice(0, Math.min(state.csvRows.length, preLines + 80));

  const header = state.csvRows[state.csvHeaderRowIndex] || [];
  const trh = document.createElement('tr');
  for (const ci of colIndices) {
    const th = document.createElement('th');
    const isFirst = ci === 0;
    if (isFirst) th.classList.add('col-first');
    if (matchedSet.has(ci)) th.classList.add('col-matched');
    th.dataset.resizeCol = String(ci);
    const hdrText = header[ci] || '';
    th.innerHTML = buildTruncateHeaderHtml(indexToColLetter(ci), ci, hdrText);
    trh.appendChild(th);
  }
  headEl.appendChild(trh);

  preview.forEach((row, rIdx) => {
    const tr = document.createElement('tr');
    if (rIdx === state.csvHeaderRowIndex) {
      tr.style.background = '#fff8e1';
      tr.style.fontWeight = '700';
    }
    for (const ci of colIndices) {
      const td = document.createElement('td');
      const isFirst = ci === 0;
      if (isFirst) td.classList.add('col-first');
      if (matchedSet.has(ci)) td.classList.add('col-matched');
      td.dataset.resizeCol = String(ci);
      const val = ci < row.length ? row[ci] : '';
      let displayText = '';
      if (rIdx === state.csvHeaderRowIndex) {
        displayText = val === '' ? '' : normalizeCell(val);
      } else {
        displayText = formatNumeric2(val);
      }
      td.textContent = displayText;
      // Hover to see full content (even when truncated by column resize)
      td.title = displayText;
      tr.appendChild(td);
    }
    bodyEl.appendChild(tr);
  });
  initResizableTables();
}

// ---------- Chart rendering ----------
function clearChart() {
  if (state.chartInstance) {
    state.chartInstance.destroy();
    state.chartInstance = null;
  }
  els.chartContainer.classList.add('hidden');
}

function showChartPlaceholder(title, meta) {
  if (state.chartInstance) {
    try { state.chartInstance.destroy(); } catch { /* noop */ }
    state.chartInstance = null;
  }
  if (els.chartContainer) els.chartContainer.classList.remove('hidden');
  if (els.chartTitle) els.chartTitle.textContent = title || '数据图表';
  if (els.chartMeta) els.chartMeta.textContent = meta || '';
  updateChartMetricSelector();
  const ctx = els.csvChart && els.csvChart.getContext ? els.csvChart.getContext('2d') : null;
  if (ctx) ctx.clearRect(0, 0, els.csvChart.width, els.csvChart.height);
}

function getChartZoomOptions() {
  return {
    pan: {
      enabled: true,
      mode: 'xy',
    },
    zoom: {
      wheel: { enabled: true },
      pinch: { enabled: true },
      mode: 'xy',
    },
    limits: {
      x: { min: 'original', max: 'original', minRange: 1 },
      y: { min: 'original', max: 'original' },
    },
  };
}

function buildChartOptions(extraPlugins = {}, extraOptions = {}) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: false,
    interaction: { mode: 'index', intersect: false },
    plugins: {
      legend: { position: 'top', labels: { boxWidth: 12, font: { size: 11 } } },
      zoom: getChartZoomOptions(),
      ...extraPlugins,
    },
    ...extraOptions,
  };
}

function bindChartInteractions(chart) {
  if (!chart || !chart.canvas) return;
  chart.canvas.ondblclick = () => {
    if (typeof chart.resetZoom === 'function') chart.resetZoom();
  };
}

function resetCurrentChartZoom() {
  if (state.chartInstance && typeof state.chartInstance.resetZoom === 'function') {
    state.chartInstance.resetZoom();
  }
}

function getChartXY() {
  if (state.csvHeaderRowIndex < 0) {
    const d = detectHeaderRow(state.csvRows);
    state.csvHeaderRowIndex = d.headerRowIndex;
    state.csvDataStartRow = d.dataStartRow;
  }
  const dataRows = state.csvRows.slice(state.csvDataStartRow >= 0 ? state.csvDataStartRow : 1);
  const total = dataRows.length;
  if (!total) return null;

  const step = Math.max(1, Math.ceil(total / CHART_MAX_POINTS));

  const xLabels = [];
  const ySeries = state.csvChartSelectedCols.map(() => []);

  for (let i = 0; i < total; i += step) {
    const row = dataRows[i];
    const xRaw = (row[0] || '').toString().trim();
    xLabels.push(xRaw);
    state.csvChartSelectedCols.forEach((c, seriesIdx) => {
      const v = (row[c] || '').toString().trim();
      let num = null;
      if (v !== '' && NUMERIC_RE.test(v)) {
        const n = Number(v);
        if (Number.isFinite(n)) num = n;
      }
      ySeries[seriesIdx].push(num);
    });
  }
  return { xLabels, ySeries, step, total };
}

function generateChart() {
  if (!state.csvRows.length) { alert('请先上传 CSV。'); return; }
  const selCols = state.csvChartSelectedCols || [];
  if (!selCols.length) { alert('请至少勾选一列作为纵坐标 Y。'); return; }
  const data = getChartXY();
  if (!data) { alert('未获取到数据区，请检查 CSV 内容。'); return; }
  const type = els.chartType.value || 'line';
  const header = state.csvRows[state.csvHeaderRowIndex] || [];
  const xHeaderName = header[0] || 'X 列';

  const palette = ['#4f46e5', '#ec4899', '#10b981', '#f59e0b', '#3b82f6', '#ef4444', '#8b5cf6', '#0ea5e9', '#14b8a6', '#f97316'];
  const datasets = selCols.map((c, i) => {
    const color = palette[i % palette.length];
    return {
      label: `${indexToColLetter(c)} · ${header[c] || '(空)'}`,
      data: data.ySeries[i],
      borderColor: color,
      backgroundColor: type === 'bar' ? color + 'B0' : color + '33',
      pointRadius: type === 'line' ? 0 : undefined,
      tension: 0.15,
      borderWidth: 1.5,
      spanGaps: true,
    };
  });

  clearChart();
  els.chartContainer.classList.remove('hidden');
  // Determine Y-axis label from active metric
  const activeMetric = state.channelMetricType || METRIC_CURRENT;
  const yLabel = METRIC_LABELS[activeMetric] ? `${METRIC_LABELS[activeMetric]} (${METRIC_UNITS[activeMetric]})` : '数值';
  els.chartTitle.textContent = `${type === 'bar' ? '柱状图' : '折线图'}：X = ${xHeaderName}`;
  els.chartMeta.textContent = `数据来源：${state.fileName || 'CSV'} · 共 ${data.total} 个数据点，绘图 ${data.xLabels.length} 点（步长 ${data.step}）`;

  // Show metric selector and update button states
  updateChartMetricSelector();

  const ctx = els.csvChart.getContext('2d');
  state.chartInstance = new Chart(ctx, {
    type,
    data: { labels: data.xLabels, datasets },
    options: buildChartOptions(
      {
        tooltip: { callbacks: { title: (items) => `${xHeaderName}: ${items[0]?.label ?? ''}` } },
      },
      {
        scales: {
          x: {
            title: { display: true, text: xHeaderName, font: { weight: 'bold' } },
            ticks: { autoSkip: true, maxTicksLimit: 12, font: { size: 10 } },
          },
          y: { title: { display: true, text: yLabel, font: { weight: 'bold' } }, ticks: { font: { size: 10 } } },
        },
      }
    ),
  });
  bindChartInteractions(state.chartInstance);
}

function updateChartMetricSelector() {
  if (!els.chartMetricSelector) return;
  const hasAnyMetricCols = Object.values(state.channelMetricCols).some((arr) => arr.length > 0);
  if (!hasAnyMetricCols) {
    els.chartMetricSelector.classList.add('hidden');
    return;
  }
  els.chartMetricSelector.classList.remove('hidden');
  const btns = els.chartMetricSelector.querySelectorAll('.chart-metric-btn');
  btns.forEach((btn) => {
    const metric = btn.dataset.metric;
    const hasCols = (state.channelMetricCols[metric] || []).length > 0;
    btn.classList.toggle('active', metric === state.channelMetricType);
    btn.classList.toggle('disabled-look', !hasCols);
  });
}

// ---------- Locate a sampling channel in CSV raw data and chart it ----------
function locateChannelInCsvAndChart(datasetId, channel) {
  const ds = state.datasets.find((d) => d.id === datasetId);
  if (ds && ds.rawCsvMissing) {
    alert('原始 CSV 未缓存，仅可预览/下载 XLSX，不能重新绘制波形。');
    return;
  }
  if (!ds || !ds.csvRows || !ds.csvRows.length) {
    alert('对应的原始 CSV 数据已不可用（可能已重新上传），无法定位绘图。');
    return;
  }
  const ch = String(channel === null || channel === undefined ? '' : channel).trim();
  if (!ch || ch === '-') { alert('该行没有有效的采样通道编号，无法定位。'); return; }

  setActiveDataset(datasetId);
  if (state.csvHeaderRowIndex < 0) {
    const d = detectHeaderRow(state.csvRows);
    ds.csvHeaderRowIndex = d.headerRowIndex;
    ds.csvDataStartRow = d.dataStartRow;
    syncActiveDatasetToState();
  }

  els.csvColSearch.value = ch;
  performColumnSearch();
  state.csvSearchMode = 'channel';
  const matched = state.csvMatchedCols || [];
  if (!matched.length) {
    updateSummary();
    renderPreview();
    alert(`未能在 CSV「${state.fileName}」的表头中找到采样通道「${ch}」对应的列。`);
    els.csvUploadArea && els.csvUploadArea.scrollIntoView({ behavior: 'smooth', block: 'center' });
    return;
  }
  // Prefer the Power column among matched (header ends with _P / contains power)
  const header = state.csvRows[state.csvHeaderRowIndex] || [];
  // Classify matched columns into current / voltage / power categories
  const metricCols = { [METRIC_CURRENT]: [], [METRIC_VOLTAGE]: [], [METRIC_POWER]: [] };
  matched.forEach((c) => {
    const h = String(header[c] || '');
    if (METRIC_REGEX[METRIC_CURRENT].test(h)) metricCols[METRIC_CURRENT].push(c);
    if (METRIC_REGEX[METRIC_VOLTAGE].test(h)) metricCols[METRIC_VOLTAGE].push(c);
    if (METRIC_REGEX[METRIC_POWER].test(h))   metricCols[METRIC_POWER].push(c);
  });
  state.channelMetricCols = metricCols;
  state.channelAllMatchedCols = matched;
  // 每次点击通道时，默认回到“电流”视图
  state.channelMetricType = METRIC_CURRENT;
  // Pick columns based on active metric type (default: current)
  const preferredMetric = state.channelMetricType;
  const preferredCols = metricCols[preferredMetric] || [];
  state.csvChartSelectedCols = preferredCols.slice(0, 6);

  updateSummary();
  updateMatchInfoAndPicker();
  renderPreview();
  if (preferredCols.length) {
    generateChart();
  } else {
    showChartPlaceholder(
      `未找到${METRIC_LABELS[preferredMetric]}数据`,
      `当前通道已匹配到其他原始列，但没有匹配到${METRIC_LABELS[preferredMetric]}列。可点击上方“电压 / 功率”切换查看。`
    );
  }
  setTimeout(() => {
    if (els.chartContainer && !els.chartContainer.classList.contains('hidden')) {
      els.chartContainer.scrollIntoView({ behavior: 'smooth', block: 'center' });
    } else if (els.csvUploadArea) {
      els.csvUploadArea.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
  }, 120);
}

// ---------- Processed preview (merge + 2-decimal + uniform numeric-col width) ----------

/**
 * Auto-detect default display columns by scanning header rows for known keywords.
 * Returns an array of 0-based column indices, or null if detection fails.
 */
function autoDetectDefaultColumns(rows) {
  if (!rows || rows.length < 2) return null;
  const maxCols = getColCount(rows);
  // Scan up to first 4 rows for headers
  const headerRows = rows.slice(0, Math.min(4, rows.length));
  // Build a text map: for each column, concatenate all header cells into one search string
  const colTexts = [];
  for (let c = 0; c < maxCols; c++) {
    let texts = '';
    for (const row of headerRows) {
      const val = row[c];
      if (val !== null && val !== undefined) texts += ' ' + String(val).trim();
    }
    colTexts.push(texts);
  }
  const matched = new Set();
  for (const rule of AUTO_DETECT_HEADERS) {
    for (let c = 0; c < maxCols; c++) {
      if (matched.has(c)) continue;
      const text = colTexts[c].toLowerCase();
      if (!text) continue;
      let hit = false;
      for (const kw of rule.keywords) {
        if (text.includes(kw.toLowerCase())) { hit = true; break; }
      }
      if (!hit) continue;
      // Check exclude patterns
      if (rule.exclude) {
        let excluded = false;
        for (const ex of rule.exclude) {
          if (text.includes(ex.toLowerCase())) { excluded = true; break; }
        }
        if (excluded) continue;
      }
      matched.add(c);
      break; // one rule matches at most one column
    }
  }
  if (matched.size < 4) return null; // too few matches, fall back
  const result = [...matched].sort((a, b) => a - b);
  return result;
}

function resolveProcessedColIndices(rowsArg) {
  const rows = rowsArg || state.processedRows;
  const maxCols = getColCount(rows.map((r) => r.map((v) => (v === undefined || v === null) ? '' : v)));
  let colIndices = state.processedColIndices;
  if (!colIndices) {
    // Try auto-detect from header keywords
    const autoDetected = autoDetectDefaultColumns(rows);
    colIndices = autoDetected || [0, 1, 2, 3, 12, 13, 14, 15];
  }
  colIndices = colIndices.filter((i) => i < maxCols);
  if (!colIndices.length) {
    colIndices = Array.from({ length: maxCols }, (_, i) => i);
  }
  return colIndices;
}

function buildDisplayMergeMeta(merges, colIndices, rowCount) {
  const displayPos = new Map();
  colIndices.forEach((ci, p) => displayPos.set(ci, p));

  const covered = new Set();
  const mergeAt = new Map();
  for (const m of (merges || [])) {
    if (m.startRow >= rowCount) continue;
    let firstPos = -1;
    let lastPos = -1;
    for (let c = m.startCol; c <= m.endCol; c += 1) {
      if (displayPos.has(c)) {
        const p = displayPos.get(c);
        if (firstPos < 0) firstPos = p;
        lastPos = p;
      }
    }
    if (firstPos < 0) continue;
    let contiguous = true;
    for (let p = firstPos; p <= lastPos; p += 1) {
      const orig = colIndices[p];
      if (orig < m.startCol || orig > m.endCol) {
        contiguous = false;
        break;
      }
    }
    if (!contiguous) continue;

    const colspan = lastPos - firstPos + 1;
    const rowspan = Math.min(m.endRow, rowCount - 1) - m.startRow + 1;
    if (rowspan <= 0) continue;

    const anchorKey = `${m.startRow}_${firstPos}`;
    mergeAt.set(anchorKey, { rowspan, colspan });
    for (let r = m.startRow; r < m.startRow + rowspan; r += 1) {
      for (let p = firstPos; p <= lastPos; p += 1) {
        if (r === m.startRow && p === firstPos) continue;
        covered.add(`${r}_${p}`);
      }
    }
  }

  return { covered, mergeAt };
}

function appendProcessedSectionRows(targetEl, rows, startRow, endRow, colIndices, numericMap, mergeMeta, options = {}) {
  const { header = false, tableType = 'processed', snapId = '', diffSet = null, powerCol = -1, channelAtMax = '', maxRow = -1, valueColSet = null, channelCol = CHANNEL_COL } = options;
  for (let rIdx = startRow; rIdx < endRow; rIdx += 1) {
    const row = rows[rIdx] || [];
    const tr = document.createElement('tr');
    for (let pos = 0; pos < colIndices.length; pos += 1) {
      const key = `${rIdx}_${pos}`;
      if (mergeMeta.covered.has(key)) continue;
      const ci = colIndices[pos];
      const cell = document.createElement(header ? 'th' : 'td');
      const merge = mergeMeta.mergeAt.get(key);
      if (merge) {
        if (merge.rowspan > 1) cell.rowSpan = merge.rowspan;
        if (merge.colspan > 1) cell.colSpan = merge.colspan;
      }
      cell.dataset.resizeCol = String(ci);
      if (numericMap.get(ci)) cell.classList.add('numeric-col');
      if (isProcessedCenteredMetricCol(ci)) cell.classList.add('num-center');
      // Value columns (Vavg/Iavg/Pavg) should display full content without truncation
      if (valueColSet && valueColSet.has(ci)) cell.classList.add('value-col-full');
      /* sticky-m-col removed — was causing ghost display bug with merged cells */
      if (header) cell.classList.add('sticky-head');

      const val = ci < row.length ? normalizeCellValue(row[ci]) : '';
      const displayText = formatNumeric2(val);

      if (header) {
        cell.innerHTML = buildTruncateHeaderHtml(indexToColLetter(ci), ci, displayText);
      } else {
        cell.textContent = displayText;
        cell.title = displayText;
      }

      const channelText = String(val === null || val === undefined ? '' : val).trim();
      if (!header && tableType === 'processed' && ci === channelCol && !merge && channelText && channelText !== '-') {
        cell.classList.add('qepm-clickable');
        cell.title = `${channelText}（点击定位 CSV 原始数据对应通道并绘制波形）`;
        cell.addEventListener('click', () => locateChannelInCsvAndChart(state.activeDatasetId, channelText));
      }

      if (!header && tableType === 'snapshot') {
        if (diffSet && diffSet.has(`${rIdx}_${ci}`)) cell.classList.add('compare-diff');
        if (channelAtMax && channelAtMax !== '-' && rIdx === maxRow) {
          if (ci === powerCol) {
            cell.classList.add('compare-max', 'clickable-power');
            cell.dataset.action = 'chart-channel';
            cell.dataset.snapid = snapId;
            cell.dataset.channel = channelAtMax;
            cell.title = '该组 Power 最大，点击在 CSV 中定位通道并绘图';
          } else if (ci === channelCol) {
            cell.classList.add('channel-highlight', 'clickable-power');
            cell.dataset.action = 'chart-channel';
            cell.dataset.snapid = snapId;
            cell.dataset.channel = channelAtMax;
            cell.title = '采样通道，点击在 CSV 中定位并绘图';
          }
        }
      }

      tr.appendChild(cell);
    }
    targetEl.appendChild(tr);
  }
}

function renderProcessedPreview() {
  const headEl = els.processedHead;
  const bodyEl = els.processedBody;
  headEl.innerHTML = '';
  bodyEl.innerHTML = '';

  // Render the dataset selection panel (self-select preview / compare)
  renderDatasetSelectPanel();

  // All processed datasets, then apply the user's preview selection (hidden set)
  const allProcessedDs = state.datasets.filter((d) => d.processedRows && d.processedRows.length);
  const processedDs = allProcessedDs.filter((d) => !state.previewHiddenIds.has(d.id));

  // Detect multi-project scenario among the SELECTED datasets: if two or more
  // selected datasets belong to different projects, render the split panel.
  const projGroups = groupDatasetsByProject(processedDs);
  const isMultiProject = projGroups.length >= 2;

  if (isMultiProject) {
    // Hide single-table wrapper, show split container
    if (els.processedSingleWrap) els.processedSingleWrap.classList.add('hidden');
    if (els.processedSplitContainer) els.processedSplitContainer.classList.remove('hidden');
    renderSplitProcessedPreview(projGroups);
    return;
  }

  // Single-project (or single dataset) mode
  if (els.processedSingleWrap) els.processedSingleWrap.classList.remove('hidden');
  if (els.processedSplitContainer) {
    els.processedSplitContainer.classList.add('hidden');
    els.processedSplitContainer.innerHTML = '';
  }

  // Nothing selected to preview
  if (!processedDs.length) return;

  // Pick which dataset to show: prefer the globally active one when it is
  // among the selected datasets, otherwise the first selected dataset.
  const showDs = processedDs.find((d) => d.id === state.activeDatasetId) || processedDs[0];
  const rows = showDs.processedRows;
  if (!rows || !rows.length) return;

  const colIndices = resolveProcessedColIndicesForRows(rows);

  // Sync the column range input to reflect auto-detected or current columns
  if (!state.processedColIndices && els.processedColRange) {
    const expr = colIndices.map((i) => indexToColLetter(i)).join(',');
    els.processedColRange.value = expr;
  }

  // Detect value columns (Vavg/Iavg/Pavg) that should NOT be truncated
  const valueColSet = detectValueColumns(rows, colIndices);

  const valueCols = colIndices.filter((ci) => ci >= COMPARE_COMMON_END);
  const stickyProbeCol = valueCols.length ? detectPowerCol(rows, valueCols) : (colIndices[colIndices.length - 1] || 0);
  const headerRowCount = Math.max(1, getProcessedDataStart(rows, stickyProbeCol));
  const dataRows = rows.slice(headerRowCount);
  const numericMap = computeNumericColumnMap(dataRows, colIndices);
  const mergeMeta = buildDisplayMergeMeta(showDs.processedMerges || [], colIndices, rows.length);

  // Detect the channel column dynamically
  const channelCol = detectChannelCol(rows);

  appendProcessedSectionRows(headEl, rows, 0, headerRowCount, colIndices, numericMap, mergeMeta, { header: true, tableType: 'processed', valueColSet, channelCol });
  appendProcessedSectionRows(bodyEl, rows, headerRowCount, rows.length, colIndices, numericMap, mergeMeta, { tableType: 'processed', valueColSet, channelCol });
  initResizableTables();
  applyStickyColumns(document);
}

/**
 * Renders the dataset selection panel that lets the user self-select which
 * datasets appear in the preview and which are included in the comparison.
 * Only shown when there are 2+ processed datasets.
 */
function renderDatasetSelectPanel() {
  const panel = els.datasetSelectPanel;
  if (!panel) return;
  const processedDs = state.datasets.filter((d) => d.processedRows && d.processedRows.length);
  if (processedDs.length < 2) {
    panel.classList.add('hidden');
    panel.innerHTML = '';
    return;
  }
  panel.classList.remove('hidden');
  panel.className = 'mt-3 ds-sel-panel';

  // Membership map: datasetId -> Set(groupId) for quick lookup
  const membership = new Map();
  state.compareGroups.forEach((g) => {
    g.snapshots.forEach((s) => {
      if (!membership.has(s.datasetId)) membership.set(s.datasetId, new Set());
      membership.get(s.datasetId).add(g.id);
    });
  });

  const rowsHtml = processedDs.map((d) => {
    const idx = state.datasets.indexOf(d);
    const color = DS_COLORS[idx % DS_COLORS.length];
    const inPreview = !state.previewHiddenIds.has(d.id);
    const memberSet = membership.get(d.id) || new Set();
    const chipsHtml = state.compareGroups.map((g, gi) => {
      const active = memberSet.has(g.id);
      return `<button type="button" class="ds-sel-gchip${active ? ' active' : ''}" data-action="toggle-group-member" data-ds-id="${escapeHtml(d.id)}" data-group-id="${escapeHtml(g.id)}" title="${active ? '点击移出' : '点击加入'}「${escapeHtml(g.name)}」">${escapeHtml(g.name)}</button>`;
    }).join('');
    const chipsWrap = state.compareGroups.length
      ? `<div class="ds-sel-gchips">${chipsHtml}</div>`
      : '<span class="ds-sel-gempty">暂无对比组</span>';
    return `<div class="ds-sel-row">
      <span class="ds-sel-badge" style="background:${color}">#${idx + 1}</span>
      <span class="ds-sel-name" title="${escapeHtml(d.fileName)}">${escapeHtml(d.fileName)}</span>
      <span class="ds-sel-proj" title="${escapeHtml(d.projectName || '未指定项目')}">${escapeHtml(d.projectName || '未指定项目')}</span>
      <label class="ds-sel-check"><input type="checkbox" data-action="toggle-preview" data-ds-id="${escapeHtml(d.id)}" ${inPreview ? 'checked' : ''}/> 预览</label>
      <span class="ds-sel-cmp-label">加入对比组：</span>
      ${chipsWrap}
    </div>`;
  }).join('');

  panel.innerHTML = `
    <div class="ds-sel-head">
      <span class="ds-sel-title">数据组选择</span>
      <span class="ds-sel-hint">勾选"预览"控制上方展示；点击组标签把该数据集加入/移出对比组（每组最多 ${COMPARE_MAX} 份，一个数据集可加入多个组）</span>
      <div class="ds-sel-quick">
        <button type="button" class="ds-sel-qbtn" data-action="preview-all">全选预览</button>
        <button type="button" class="ds-sel-qbtn" data-action="preview-none">清空预览</button>
        <button type="button" class="ds-sel-qbtn" data-action="add-group">➕ 新建对比组</button>
        <button type="button" class="ds-sel-qbtn" data-action="compare-all">全部加入对比组 1</button>
        <button type="button" class="ds-sel-qbtn" data-action="compare-none">清空对比组 1</button>
      </div>
    </div>
    <div class="ds-sel-list">${rowsHtml}</div>
  `;
}

/**
 * Toggle whether a dataset is shown in the result preview.
 */
function togglePreviewSelect(dsId, shouldShow) {
  if (shouldShow) state.previewHiddenIds.delete(dsId);
  else state.previewHiddenIds.add(dsId);
  renderProcessedPreview();
}

/**
 * Toggle a dataset's membership in a specific compare group.
 */
function toggleCompareGroupMembership(dsId, groupId) {
  const ds = state.datasets.find((d) => d.id === dsId);
  const group = getCompareGroup(groupId);
  if (!ds || !group) return;
  const existing = group.snapshots.find((s) => s.datasetId === dsId);
  if (existing) {
    group.snapshots = group.snapshots.filter((s) => s.datasetId !== dsId);
  } else {
    if (group.snapshots.length >= COMPARE_MAX) {
      alert(`「${group.name}」最多同时对比 ${COMPARE_MAX} 份，请先移除一份。`);
      renderDatasetSelectPanel(); // revert
      return;
    }
    group.snapshots.push(makeSnapshotFromDataset(ds));
  }
  renderCompareArea();
  renderDatasetSelectPanel();
  syncPreviewSectionVisibility();
}

/**
 * Group datasets by projectId. Datasets missing a projectId are grouped under
 * a synthetic '__unassigned__' key (should be rare — datasets always inherit
 * the global default at upload time).
 * Returns: [{ projectId, projectName, colorIdx, datasets: [] }]
 */
function groupDatasetsByProject(datasets) {
  const orderMap = new Map(); // projectId -> group
  const projectOrder = [];
  datasets.forEach((d) => {
    const pid = d.projectId || '__unassigned__';
    if (!orderMap.has(pid)) {
      orderMap.set(pid, {
        projectId: pid,
        projectName: d.projectName || (pid === '__unassigned__' ? '未指定项目' : pid),
        colorIdx: projectOrder.length,
        datasets: [],
      });
      projectOrder.push(pid);
    }
    orderMap.get(pid).datasets.push(d);
  });
  return projectOrder.map((pid) => orderMap.get(pid));
}

/**
 * Ensure state.projectActiveDsIds has a valid entry for each project group.
 * If missing or the pointed dataset was removed, default to the group's first
 * dataset. Prefer the globally active dataset when it belongs to the group.
 */
function ensureProjectActiveDsIds(projGroups) {
  const nextMap = {};
  projGroups.forEach((g) => {
    const prevId = state.projectActiveDsIds[g.projectId];
    const found = prevId && g.datasets.find((d) => d.id === prevId);
    if (found) {
      nextMap[g.projectId] = prevId;
    } else if (g.datasets.find((d) => d.id === state.activeDatasetId)) {
      nextMap[g.projectId] = state.activeDatasetId;
    } else {
      nextMap[g.projectId] = g.datasets[0].id;
    }
  });
  state.projectActiveDsIds = nextMap;
}

/**
 * Render the multi-project split preview into #processedSplitContainer.
 * Each project group gets a column. When a project group has multiple
 * datasets, a per-column tab strip allows switching between them.
 */
function renderSplitProcessedPreview(projGroups) {
  const container = els.processedSplitContainer;
  if (!container) return;
  ensureProjectActiveDsIds(projGroups);

  const nCols = Math.min(projGroups.length, 4);
  container.className = `mt-3 proj-split-grid cols-${nCols}`;
  container.innerHTML = '';

  projGroups.forEach((group) => {
    const activeDsId = state.projectActiveDsIds[group.projectId];
    const activeDs = group.datasets.find((d) => d.id === activeDsId) || group.datasets[0];
    const color = DS_COLORS[group.colorIdx % DS_COLORS.length];

    const col = document.createElement('div');
    col.className = 'proj-split-col';

    // Header
    const head = document.createElement('div');
    head.className = 'proj-split-head';
    head.innerHTML = `
      <span class="proj-split-title" title="${escapeHtml(group.projectName)}">
        <span class="proj-dot" style="background:${color}"></span>
        <span>${escapeHtml(group.projectName)}</span>
      </span>
      <span class="proj-split-meta">${group.datasets.length} 组数据</span>
    `;
    col.appendChild(head);

    // Per-project dataset tabs (only when >1 dataset in this project)
    if (group.datasets.length > 1) {
      const tabs = document.createElement('div');
      tabs.className = 'proj-split-tabs';
      tabs.innerHTML = group.datasets.map((d, i) => {
        const active = d.id === activeDs.id;
        return `<div class="ds-tab ${active ? 'ds-tab-active' : ''}" data-action="pick-split-ds"
          data-project-id="${escapeHtml(group.projectId)}" data-ds-id="${escapeHtml(d.id)}"
          title="${escapeHtml(d.fileName)}">
          <span class="ds-dot" style="background:${color}"></span>
          <span class="ds-name">${escapeHtml(d.fileName)}</span>
        </div>`;
      }).join('');
      col.appendChild(tabs);
    }

    // Body — one table per column
    const body = document.createElement('div');
    body.className = 'proj-split-body tight-scroll';
    const table = document.createElement('table');
    table.className = 'text-xs processed-table preview-table compact-table numeric-uniform resizable-table';
    const thead = document.createElement('thead');
    thead.className = 'sticky top-0 bg-indigo-50 text-slate-700';
    const tbody = document.createElement('tbody');
    tbody.className = 'divide-y divide-slate-100';
    table.appendChild(thead);
    table.appendChild(tbody);
    body.appendChild(table);
    col.appendChild(body);

    // Render the active dataset's processed data into this column
    renderProcessedIntoTable(activeDs, thead, tbody);

    container.appendChild(col);
  });

  // Initialize resizers for all inner tables + sticky heads
  initResizableTables(container);
  applyStickyColumns(container);
}

/**
 * Renders a single dataset's processed rows into the given thead/tbody
 * elements. Shared logic extracted so both single-view and split-view can use
 * the same rendering path.
 */
function renderProcessedIntoTable(ds, headEl, bodyEl) {
  headEl.innerHTML = '';
  bodyEl.innerHTML = '';
  const rows = ds.processedRows || [];
  if (!rows.length) return;

  const colIndices = resolveProcessedColIndicesForRows(rows);
  const valueColSet = detectValueColumns(rows, colIndices);
  const valueCols = colIndices.filter((ci) => ci >= COMPARE_COMMON_END);
  const stickyProbeCol = valueCols.length ? detectPowerCol(rows, valueCols) : (colIndices[colIndices.length - 1] || 0);
  const headerRowCount = Math.max(1, getProcessedDataStart(rows, stickyProbeCol));
  const dataRows = rows.slice(headerRowCount);
  const numericMap = computeNumericColumnMap(dataRows, colIndices);
  const mergeMeta = buildDisplayMergeMeta(ds.processedMerges || [], colIndices, rows.length);
  const channelCol = detectChannelCol(rows);

  appendProcessedSectionRows(headEl, rows, 0, headerRowCount, colIndices, numericMap, mergeMeta, { header: true, tableType: 'processed', valueColSet, channelCol });
  appendProcessedSectionRows(bodyEl, rows, headerRowCount, rows.length, colIndices, numericMap, mergeMeta, { tableType: 'processed', valueColSet, channelCol });
}

/**
 * Resolve column indices for a given rows array (independent of state).
 * Falls back to auto-detected default columns if no explicit user filter.
 */
function resolveProcessedColIndicesForRows(rows) {
  if (state.processedColIndices && state.processedColIndices.length) {
    return state.processedColIndices.slice();
  }
  return autoDetectDefaultColumns(rows);
}

/**
 * Detect which columns in colIndices are "value" columns (Vavg/Iavg/Pavg) based on header keywords.
 */
function detectValueColumns(rows, colIndices) {
  const result = new Set();
  const headerRows = rows.slice(0, Math.min(4, rows.length));
  for (const ci of colIndices) {
    let text = '';
    for (const row of headerRows) {
      const v = row[ci];
      if (v !== null && v !== undefined) text += ' ' + String(v).trim();
    }
    const lower = text.toLowerCase();
    for (const kw of VALUE_COL_KEYWORDS) {
      if (lower.includes(kw)) { result.add(ci); break; }
    }
  }
  return result;
}

/**
 * Detect the "采样通道编号(QEPM)" column index from processed data headers.
 * Returns the detected index or falls back to CHANNEL_COL constant.
 */
function detectChannelCol(rows) {
  if (!rows || rows.length < 2) return CHANNEL_COL;
  const maxCols = getColCount(rows);
  const headerRows = rows.slice(0, Math.min(4, rows.length));
  const channelKeywords = ['采样通道编号', 'qepm', '通道编号'];
  for (let c = 0; c < maxCols; c++) {
    let text = '';
    for (const row of headerRows) {
      const v = row[c];
      if (v !== null && v !== undefined) text += ' ' + String(v).trim();
    }
    const lower = text.toLowerCase();
    for (const kw of channelKeywords) {
      if (lower.includes(kw)) return c;
    }
  }
  return CHANNEL_COL;
}

/**
 * Detect the Vavg column index from processed data headers.
 * Keywords (case-insensitive, spaces ignored): vavg, v_avg, v avg, vavg(v).
 * Returns -1 when not found.
 */
function detectVavgCol(rows) {
  if (!rows || !rows.length) return -1;
  const maxCols = getColCount(rows);
  const headerRows = rows.slice(0, Math.min(4, rows.length));
  const keywords = ['vavg', 'v_avg', 'v avg', 'vavg(v)'];
  const normalizedKeywords = keywords.map((kw) => kw.toLowerCase().replace(/\s+/g, ''));
  for (let c = 0; c < maxCols; c++) {
    let text = '';
    for (const row of headerRows) {
      const v = row[c];
      if (v !== null && v !== undefined) text += ' ' + String(v).trim();
    }
    const lower = text.toLowerCase().replace(/\s+/g, '');
    for (const nk of normalizedKeywords) {
      if (lower.includes(nk)) return c;
    }
  }
  return -1;
}

/**
 * Get the channel column index for current processed data.
 */
function getChannelCol() {
  return detectChannelCol(state.processedRows);
}

function renderPreview() {
  renderCsvPreview();
  renderProcessedPreview();
  renderCompareArea();
}

// When the active project changes, all datasets' processed results become stale.
// Note: per-file project associations (projectId/projectName) are intentionally
// preserved here — they are user-controlled assignments, not derived output.
function resetAllProcessed() {
  state.datasets.forEach((d) => {
    d.processedRows = [];
    d.processedMerges = [];
    d.workbookBlob = null;
  });
  syncActiveDatasetToState();
}

// Backward-compatible "set all" shortcut: applies the given project as the
// default to every dataset that does not yet have its own project assigned.
// Datasets that already carry an explicit per-file project are left untouched.
function applyDefaultProjectToDatasets(record) {
  if (!record) return;
  state.datasets.forEach((d) => {
    if (!d.projectId) {
      d.projectId = record.id;
      d.projectName = record.name;
    }
  });
}

// Handles a change on a per-file project dropdown in the uploaded file list.
function handleDatasetProjectChange(dsId, projectId) {
  const ds = state.datasets.find((d) => d.id === dsId);
  if (!ds) return;
  if (projectId) {
    const item = state.templateItems.find((x) => x.id === projectId);
    ds.projectId = projectId;
    ds.projectName = item ? item.name : (ds.projectName || '未命名项目');
  } else if (state.activeProject) {
    // Empty selection falls back to the global default project.
    ds.projectId = state.activeProject.id;
    ds.projectName = state.activeProject.name;
  } else {
    ds.projectId = null;
    ds.projectName = null;
  }
  // The project changed, so any previously generated result for this file is stale.
  ds.processedRows = [];
  ds.processedMerges = [];
  ds.workbookBlob = null;
  if (ds.id === state.activeDatasetId) syncActiveDatasetToState();
  updateSummary();
  renderPreview();
}

// ---------- Template list: pagination + views ----------
function switchTemplateView(view) {
  state.templateView = view;
  const isList = view === 'list';
  const isDetail = view === 'detail';
  const isUpload = view === 'upload';
  els.templateListView.classList.toggle('hidden', !isList);
  els.templateDetailView.classList.toggle('hidden', !isDetail);
  els.templateUploadView.classList.toggle('hidden', !isUpload);
  if (isDetail) renderTemplateDetail();
}

async function refreshTemplateList() {
  let serverItems = [];
  let serverErr = null;
  try {
    serverItems = await templateApi.list();
    if (!Array.isArray(serverItems)) serverItems = [];
    state.templatesLoadError = null;
  } catch (error) {
    serverErr = error;
    state.templatesLoadError = error && error.message ? error.message : String(error);
    console.warn('failed to load templates from backend', error);
  }

  // Normalize the backend records into the shape the rest of the UI expects.
  const items = serverItems.map((rec) => ({
    id: rec.id,
    name: rec.name || rec.id,
    fileName: rec.fileName || `${rec.name || rec.id}.xlsx`,
    size: Number(rec.sizeBytes || 0),
    sizeBytes: Number(rec.sizeBytes || 0),
    version: Number(rec.version || 1),
    createdAt: rec.createdAt ? Date.parse(rec.createdAt) : Date.now(),
    updatedAt: rec.updatedAt ? Date.parse(rec.updatedAt) : Date.now(),
    createdBy: rec.createdBy || '',
    updatedBy: rec.updatedBy || '',
    source: 'server',
  }));

  items.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  state.templateItems = items;

  // If the server-side list changed the active project's version, drop the
  // cached buffer so the next use will re-download the latest bytes.
  if (state.activeProject) {
    const match = items.find((it) => it.id === state.activeProject.id);
    if (!match) {
      state.activeProject = null;
    } else if (state.activeProject.version !== match.version) {
      state.activeProject = Object.assign({}, match);
    } else {
      state.activeProject = Object.assign({}, match, {
        buffer: state.activeProject.buffer,
      });
    }
  }

  refreshProjectSelect(items);

  const totalPages = Math.max(1, Math.ceil(items.length / TEMPLATE_PAGE_SIZE));
  if (state.templatePage > totalPages) state.templatePage = totalPages;
  if (state.templatePage < 1) state.templatePage = 1;

  renderTemplateListPage();
  if (state.templateView === 'detail') renderTemplateDetail();

  if (serverErr && els.mainStatus) {
    // Non-fatal notification. Other CSV/preview flows keep working.
    console.warn('[templates] backend unreachable — see /api/templates for details');
  }
}

function renderTemplateListPage() {
  const items = state.templateItems;
  if (!items.length) {
    const errMsg = state.templatesLoadError
      ? `<div class="text-red-500 mt-1">加载失败：${escapeHtml(state.templatesLoadError)}</div>`
      : '';
    els.templateList.innerHTML = `
      <div class="empty-state rounded-xl p-3 text-slate-500 bg-white border border-dashed border-slate-200 text-center text-xs">
        暂无保存的项目模板${errMsg}
      </div>`;
    els.templatePaginationBar.classList.add('hidden');
    return;
  }

  const totalPages = Math.max(1, Math.ceil(items.length / TEMPLATE_PAGE_SIZE));
  const page = state.templatePage;
  const start = (page - 1) * TEMPLATE_PAGE_SIZE;
  const pageItems = items.slice(start, start + TEMPLATE_PAGE_SIZE);

  els.templateList.innerHTML = pageItems.map((item) => {
    const isActive = state.activeProject && state.activeProject.id === item.id;
    const by = item.updatedBy ? ` · 由 ${escapeHtml(item.updatedBy)}` : '';
    return `
      <div class="template-item ${isActive ? 'template-item-active' : ''}" data-id="${item.id}" data-role="row">
        <div class="ti-main">
          <span class="material-symbols-outlined text-emerald-600" style="font-size:18px">description</span>
          <div class="min-w-0 flex-1">
            <div class="ti-name">
              ${escapeHtml(item.name)}
              ${isActive ? '<span class="active-badge">当前使用</span>' : ''}
            </div>
            <div class="ti-meta">${escapeHtml(item.fileName)} · ${formatSize(item.size)} · ${formatTime(item.updatedAt)}${by}</div>
          </div>
        </div>
        <div class="ti-actions">
          <button class="template-btn template-btn-primary template-btn-sm" data-action="use" data-id="${item.id}">使用</button>
          <button class="template-btn template-btn-secondary template-btn-sm" data-action="detail" data-id="${item.id}" title="进入详情（下载 / 上传 / 删除）">
            <span class="material-symbols-outlined text-sm">chevron_right</span>
          </button>
        </div>
      </div>`;
  }).join('');

  els.templatePaginationBar.classList.toggle('hidden', totalPages <= 1);
  els.templatePageInfo.textContent = `共 ${items.length} 项 · 第 ${page} / ${totalPages} 页`;
  els.templatePrevBtn.disabled = page <= 1;
  els.templatePrevBtn.classList.toggle('opacity-40', page <= 1);
  els.templateNextBtn.disabled = page >= totalPages;
  els.templateNextBtn.classList.toggle('opacity-40', page >= totalPages);
}

function renderTemplateDetail() {
  const id = state.templateDetailId;
  const item = state.templateItems.find((x) => x.id === id);
  if (!item) {
    els.templateDetailBody.innerHTML = `<div class="empty-state rounded-xl p-3 text-slate-500 bg-white border border-dashed border-slate-200 text-center text-xs">未找到模板，可能已被删除。</div>`;
    return;
  }
  const isActive = state.activeProject && state.activeProject.id === item.id;
  els.templateDetailBody.innerHTML = `
    <div class="template-detail-card">
      <div class="flex items-start justify-between gap-2 flex-wrap">
        <div class="min-w-0">
          <div class="template-detail-title">
            ${escapeHtml(item.name)}
            ${isActive ? '<span class="active-badge">当前使用</span>' : ''}
          </div>
          <div class="template-detail-meta">
            文件：${escapeHtml(item.fileName)}<br/>
            大小：${formatSize(item.size)} · 版本：v${item.version || 1}<br/>
            更新：${formatTime(item.updatedAt)}${item.updatedBy ? ` · 由 ${escapeHtml(item.updatedBy)}` : ''}
          </div>
        </div>
      </div>
      <div class="mt-3 flex flex-wrap gap-2">
        <button class="template-btn template-btn-primary" data-action="use" data-id="${item.id}">
          <span class="material-symbols-outlined text-sm">check_circle</span>
          使用此模板
        </button>
        <button class="template-btn template-btn-secondary" data-action="download" data-id="${item.id}">
          <span class="material-symbols-outlined text-sm">download</span>
          下载模板
        </button>
        <button class="template-btn template-btn-secondary admin-only-btn" data-action="rename" data-id="${item.id}" title="重命名此模板（管理员）">
          <span class="material-symbols-outlined text-sm">edit</span>
          重命名
        </button>
        <button class="template-btn template-btn-primary admin-only-btn" data-action="replace" data-id="${item.id}" title="上传新文件替换此模板（管理员）">
          <span class="material-symbols-outlined text-sm">upload</span>
          上传替换
        </button>
        <button class="template-btn template-btn-danger admin-only-btn" data-action="delete" data-id="${item.id}">
          <span class="material-symbols-outlined text-sm">delete</span>
          删除模板
        </button>
      </div>
    </div>
  `;
}

function refreshProjectSelect(items) {
  if (!els.projectSelect) return;
  const activeId = state.activeProject ? state.activeProject.id : '';
  const opts = ['<option value="">-- 请选择一个项目 --</option>'];
  for (const it of items) {
    const sel = it.id === activeId ? ' selected' : '';
    opts.push(`<option value="${escapeHtml(it.id)}"${sel}>${escapeHtml(it.name)}</option>`);
  }
  els.projectSelect.innerHTML = opts.join('');
  if (activeId) els.projectSelect.value = activeId;
  if (!items.length) {
    els.projectSelectHint.textContent = '暂无可用项目，请在右侧"项目模板管理"新增模板。';
  } else {
    els.projectSelectHint.textContent = '选择后将自动把该项目的模板设为当前使用模板。';
  }
}

// ---------- Template blob cache (IndexedDB) ----------
// The metadata list always comes from the backend (`templateApi.list()`),
// but the xlsx binary is cached locally per (id, version). This dramatically
// speeds up "generate result" and template downloads once a template has
// been fetched once. Cache key = tpl_${id}_v${version}.

function templateCacheKey(id, version) {
  return `tpl_${id}_v${version || 1}`;
}

async function loadTemplateBufferCached(templateItem) {
  if (!templateItem || !templateItem.id) throw new Error('无效的模板记录');
  const key = templateCacheKey(templateItem.id, templateItem.version);

  // Fast path: in-memory buffer on the active project.
  if (state.activeProject && state.activeProject.id === templateItem.id
      && state.activeProject.version === templateItem.version
      && state.activeProject.buffer) {
    return state.activeProject.buffer;
  }

  // Try IndexedDB.
  try {
    const cached = await dbGet(key);
    if (cached && cached.buffer && cached.version === templateItem.version) {
      return cached.buffer;
    }
  } catch (e) { console.warn('template cache read failed', e); }

  // Fallback: download from backend.
  const dl = await templateApi.download(templateItem.id);
  const buffer = dl.buffer;

  try {
    await dbPut({
      id: key,
      templateId: templateItem.id,
      name: templateItem.name,
      fileName: templateItem.fileName,
      version: templateItem.version,
      size: buffer.byteLength,
      updatedAt: templateItem.updatedAt,
      buffer,
    });
    await pruneStaleBlobCacheForTemplate(templateItem.id, templateItem.version);
  } catch (e) { console.warn('template cache write failed', e); }

  return buffer;
}

async function pruneStaleBlobCacheForTemplate(templateId, keepVersion) {
  try {
    const all = await dbGetAll();
    for (const rec of all) {
      if (rec && rec.templateId === templateId && rec.version !== keepVersion) {
        try { await dbDelete(rec.id); } catch (e) { /* ignore */ }
      }
    }
  } catch (e) { /* ignore */ }
}

async function pruneObsoleteBlobCache() {
  // Drop any cached blobs whose template no longer exists on the server.
  try {
    const [all, serverItems] = await Promise.all([dbGetAll(), templateApi.list().catch(() => [])]);
    const validKeys = new Set();
    (serverItems || []).forEach((it) => validKeys.add(templateCacheKey(it.id, it.version)));
    for (const rec of all) {
      // Legacy records used the templateId as the key. Delete anything that
      // is neither a valid current cache entry nor an unknown legacy shape.
      if (rec && rec.id && rec.id.startsWith('tpl_') && rec.id.indexOf('_v') !== -1) {
        if (!validKeys.has(rec.id)) {
          try { await dbDelete(rec.id); } catch (e) { /* ignore */ }
        }
      }
    }
  } catch (e) { /* ignore */ }
}

// ---------- Template actions ----------
async function useTemplate(id) {
  const record = state.templateItems.find((x) => x.id === id);
  if (!record) { alert('未找到该模板，可能已被删除或与服务端不同步。'); return; }
  state.activeProject = Object.assign({}, record);
  applyDefaultProjectToDatasets(record);
  resetAllProcessed();
  updateSummary();
  renderPreview();
  refreshTemplateList();
  // Warm the local cache in the background so downstream export is snappy.
  loadTemplateBufferCached(record).catch((error) => {
    console.warn('warm-cache template failed', error);
  });
}

async function downloadTemplate(id) {
  const record = state.templateItems.find((x) => x.id === id);
  if (!record) { alert('未找到该模板，可能已被删除。'); return; }
  let buffer;
  try {
    buffer = await loadTemplateBufferCached(record);
  } catch (error) {
    alert('下载模板失败：' + (error.message || '未知错误'));
    return;
  }
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = record.fileName || `${record.name}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function deleteTemplate(id) {
  const record = state.templateItems.find((x) => x.id === id);
  if (!record) return;
  if (!confirm(`确定要删除项目模板"${record.name}"？该操作不可恢复，所有成员都会看不到这个模板。`)) return;
  try {
    await templateApi.remove(id);
  } catch (error) {
    alert('删除失败：' + (error.message || '未知错误'));
    return;
  }
  if (state.activeProject && state.activeProject.id === id) {
    state.activeProject = null;
    resetAllProcessed();
  }
  state.datasets.forEach((d) => {
    if (d.projectId === id) {
      d.projectId = state.activeProject ? state.activeProject.id : null;
      d.projectName = state.activeProject ? state.activeProject.name : null;
    }
  });
  if (state.templateView === 'detail' && state.templateDetailId === id) {
    state.templateDetailId = null;
    switchTemplateView('list');
  }
  updateSummary();
  renderPreview();
  await refreshTemplateList();
  pruneObsoleteBlobCache();
}

function syncRenamedTemplateLocally(id, trimmed) {
  if (state.activeProject && state.activeProject.id === id) {
    state.activeProject.name = trimmed;
  }
  state.datasets.forEach((d) => {
    if (d.projectId === id) d.projectName = trimmed;
  });
  updateSummary();
}

function deriveTemplateNameFromFile(file) {
  const raw = String((file && file.name) || '').replace(/\.xlsx$/i, '').trim();
  return raw || `模板_${Date.now().toString(36)}`;
}

function closeTemplateRenameModal() {
  state.templateRenameContext = null;
  if (els.templateRenameModal) els.templateRenameModal.classList.add('hidden');
  if (els.templateRenameError) els.templateRenameError.classList.add('hidden');
}

function openTemplateRenameModal({ id, currentName, afterUpload = false, fileName = '' }) {
  state.templateRenameContext = { id, afterUpload, fileName };
  if (els.templateRenameDesc) {
    els.templateRenameDesc.textContent = afterUpload
      ? `模板文件已上传成功${fileName ? `（${fileName}）` : ''}，请现在补充项目名称，方便后续选择和管理。`
      : '请输入新的模板名称。';
  }
  if (els.templateRenameInput) {
    els.templateRenameInput.value = currentName || '';
  }
  if (els.templateRenameError) els.templateRenameError.classList.add('hidden');
  if (els.templateRenameModal) els.templateRenameModal.classList.remove('hidden');
  requestAnimationFrame(() => {
    if (els.templateRenameInput) {
      els.templateRenameInput.focus();
      els.templateRenameInput.select();
    }
  });
}

async function submitTemplateRenameModal() {
  const ctx = state.templateRenameContext;
  if (!ctx) return;
  const id = ctx.id;
  const value = (els.templateRenameInput && els.templateRenameInput.value || '').trim();
  if (!value) {
    if (els.templateRenameError) els.templateRenameError.classList.remove('hidden');
    return;
  }
  const record = state.templateItems.find((x) => x.id === id);
  if (record && value === record.name) {
    closeTemplateRenameModal();
    return;
  }
  try {
    await templateApi.update(id, { name: value });
  } catch (error) {
    alert('重命名失败：' + (error.message || '未知错误'));
    return;
  }
  syncRenamedTemplateLocally(id, value);
  await refreshTemplateList();
  state.templateDetailId = id;
  switchTemplateView('detail');
  closeTemplateRenameModal();
}

async function renameTemplate(id) {
  const record = state.templateItems.find((x) => x.id === id);
  if (!record) { alert('未找到该模板，可能已被删除。'); return; }
  openTemplateRenameModal({ id, currentName: record.name, afterUpload: false, fileName: record.fileName });
}

// Kept for backward compatibility. The signature (file, projectName) is
// preserved but the implementation now goes through the backend.
async function saveTemplate(file, projectName, opts = {}) {
  const record = opts.replaceId
    ? await templateApi.update(opts.replaceId, { file, name: projectName, fileName: file.name })
    : await templateApi.upload({ file, name: projectName, fileName: file.name });
  return {
    id: record.id,
    name: record.name,
    fileName: record.fileName,
    size: Number(record.sizeBytes || 0),
    version: Number(record.version || 1),
    updatedAt: record.updatedAt ? Date.parse(record.updatedAt) : Date.now(),
  };
}

// ---------- Upload handlers ----------
function handleTemplateFile(file, opts = {}) {
  if (!file) return;
  if (!state.isAdmin && !state.serverIsAdmin) { alert('只有管理员可以上传模板文件，请先登录管理员。'); return; }
  if (!file.name.toLowerCase().endsWith('.xlsx')) { alert('只支持上传 .xlsx 模板文件。'); return; }
  let projectName;
  let shouldPromptRenameAfterUpload = false;
  if (opts.replaceId) {
    const existing = state.templateItems.find((x) => x.id === opts.replaceId);
    projectName = existing ? existing.name : '';
  } else {
    projectName = (els.projectNameInput.value || '').trim();
    if (!projectName) {
      projectName = deriveTemplateNameFromFile(file);
      shouldPromptRenameAfterUpload = true;
    }
  }
  saveTemplate(file, projectName, opts).then(async (record) => {
    // Prime the local cache with the just-uploaded buffer to avoid a redundant download.
    try {
      const buffer = await file.arrayBuffer();
      await dbPut({
        id: templateCacheKey(record.id, record.version),
        templateId: record.id,
        name: record.name,
        fileName: record.fileName,
        version: record.version,
        size: buffer.byteLength,
        updatedAt: record.updatedAt,
        buffer,
      });
      await pruneStaleBlobCacheForTemplate(record.id, record.version);
    } catch (e) { /* cache warm-up failure is non-fatal */ }

    state.activeProject = Object.assign({}, record);
    applyDefaultProjectToDatasets(record);
    resetAllProcessed();
    if (els.projectNameInput) els.projectNameInput.value = '';
    updateSummary();
    renderPreview();
    await refreshTemplateList();
    state.templateDetailId = record.id;
    switchTemplateView('detail');
    if (shouldPromptRenameAfterUpload) {
      openTemplateRenameModal({ id: record.id, currentName: record.name, afterUpload: true, fileName: file.name });
    }
  }).catch((error) => {
    console.error(error);
    if (error && error.status === 403) {
      alert('后端拒绝了此操作：您不在管理员白名单中。');
    } else if (error && error.status === 401) {
      alert('未登录或身份识别失败，请刷新页面重试。');
    } else {
      alert('保存模板失败：' + (error.message || '未知错误'));
    }
  });
}

// ---------- Batch CSV upload ----------
function handleCsvFiles(fileList) {
  const files = Array.from(fileList || []).filter((f) => f && f.name.toLowerCase().endsWith('.csv'));
  if (!files.length) { alert('请选择 .csv 格式文件。'); return; }
  if (files.length > BATCH_MAX) { alert(`一次最多批量上传 ${BATCH_MAX} 个 CSV 文件。`); return; }

  const results = new Array(files.length);
  let pending = files.length;
  let hadError = false;

  files.forEach((file, i) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const rows = parseCsvText(reader.result);
        const d = detectHeaderRow(rows);
        results[i] = {
          id: genDatasetId(),
          fileName: file.name,
          csvRows: rows,
          csvHeaderRowIndex: d.headerRowIndex,
          csvDataStartRow: d.dataStartRow,
          processedRows: [],
          processedMerges: [],
          workbookBlob: null,
          // Each dataset carries its own project association. Initialize it from
          // the currently selected global project (acts as the default), but it
          // can be overridden per-file later via the per-file project dropdown.
          projectId: state.activeProject ? state.activeProject.id : null,
          projectName: state.activeProject ? state.activeProject.name : null,
        };
      } catch (error) {
        console.error(error);
        hadError = true;
        alert(`CSV「${file.name}」读取失败，已跳过。`);
      }
      pending -= 1;
      if (pending === 0) finalizeBatch(results.filter(Boolean), hadError);
    };
    reader.onerror = () => {
      hadError = true;
      pending -= 1;
      if (pending === 0) finalizeBatch(results.filter(Boolean), hadError);
    };
    reader.readAsText(file, 'utf-8');
  });
}

function finalizeBatch(newDatasets) {
  if (!newDatasets.length) return;
  // Replace the current working set with the newly uploaded batch.
  // Compare snapshots are preserved so cross-project comparisons keep working.
  state.datasets = newDatasets;
  state.activeDatasetId = newDatasets[0].id;
  syncActiveDatasetToState();

  state.csvSearchKeyword = '';
  state.csvSearchMode = 'manual';
  state.csvMatchedCols = [];
  state.csvChartSelectedCols = [];
  clearChart();
  if (els.csvColSearch) els.csvColSearch.value = '';

  updateMatchInfoAndPicker();
  updateSummary();
  renderPreview();
}

// ---------- Excel processing ----------
// Loads an .xlsx template workbook. When `projectId` is provided the template
// is resolved from `state.templateItems` (server-authoritative list) and its
// binary loaded from IndexedDB cache or the backend. When omitted it falls back
// to the currently active project.
async function loadTemplateWorkbook(projectId) {
  let record;
  if (projectId) {
    record = state.templateItems.find((x) => x.id === projectId) || null;
    if (!record && state.activeProject && state.activeProject.id === projectId) {
      record = state.activeProject;
    }
    if (!record) throw new Error('该文件所选的项目模板已不存在，请重新选择模板。');
  } else {
    if (!state.activeProject) throw new Error('请先选择或上传项目模板。');
    const fresh = state.templateItems.find((x) => x.id === state.activeProject.id);
    record = fresh || state.activeProject;
  }
  const buffer = await loadTemplateBufferCached(record);
  if (state.activeProject && state.activeProject.id === record.id) {
    state.activeProject.buffer = buffer;
    state.activeProject.version = record.version;
  }
  return XlsxPopulate.fromDataAsync(buffer.slice(0));
}
function createBlankMatrix(rowCount, colCount) {
  return Array.from({ length: rowCount }, () => Array(colCount).fill(''));
}
function padRowsInPlace(rows, maxCols) {
  rows.forEach((row) => { while (row.length < maxCols) row.push(''); });
  return rows;
}
function toCellValue(raw) {
  if (raw === null || raw === undefined) return '';
  if (typeof raw === 'number') return raw;
  const s = String(raw);
  const trimmed = s.trim();
  if (trimmed === '') return s;
  if (NUMERIC_RE.test(trimmed)) {
    const n = Number(trimmed);
    if (Number.isFinite(n)) return n;
  }
  return s;
}
function waitForNextFrame() {
  return new Promise((resolve) => {
    if (typeof requestAnimationFrame === 'function') requestAnimationFrame(() => resolve());
    else setTimeout(resolve, 0);
  });
}
async function clearSheetRange(sheet, startRow, startCol, endRow, endCol) {
  if (endRow < startRow || endCol < startCol) return;
  const colCount = endCol - startCol + 1;
  let batchCount = 0;
  for (let row = startRow; row <= endRow; row += WRITE_BATCH_SIZE) {
    const batchEndRow = Math.min(endRow, row + WRITE_BATCH_SIZE - 1);
    sheet.range(row, startCol, batchEndRow, endCol).value(createBlankMatrix(batchEndRow - row + 1, colCount));
    batchCount += 1;
    if (batchCount % UI_YIELD_EVERY_BATCHES === 0) await waitForNextFrame();
  }
}
async function clearTrailingSheetValues(sheet, rowCount, colCount) {
  const usedRange = sheet.usedRange();
  if (!usedRange) return;
  const usedEndRow = usedRange.endCell().rowNumber();
  const usedEndCol = usedRange.endCell().columnNumber();
  if (!rowCount || !colCount) { await clearSheetRange(sheet, 1, 1, usedEndRow, usedEndCol); return; }
  if (usedEndRow > rowCount) await clearSheetRange(sheet, rowCount + 1, 1, usedEndRow, usedEndCol);
  if (usedEndCol > colCount) await clearSheetRange(sheet, 1, colCount + 1, Math.min(rowCount, usedEndRow), usedEndCol);
}
async function writeCsvRowsToSheet(sheet, rows, onProgress) {
  const rowCount = rows.length;
  const maxCols = getColCount(rows);
  await clearTrailingSheetValues(sheet, rowCount, maxCols);
  if (!rowCount || !maxCols) return;
  padRowsInPlace(rows, maxCols);
  // Pre-convert all values upfront for better performance
  const converted = new Array(rowCount);
  for (let i = 0; i < rowCount; i++) {
    const row = rows[i];
    const out = new Array(maxCols);
    for (let j = 0; j < maxCols; j++) out[j] = toCellValue(row[j]);
    converted[i] = out;
  }
  const totalBatches = Math.ceil(rowCount / WRITE_BATCH_SIZE);
  let batchCount = 0;
  for (let start = 0; start < rowCount; start += WRITE_BATCH_SIZE) {
    const end = Math.min(start + WRITE_BATCH_SIZE, rowCount);
    const chunk = converted.slice(start, end);
    sheet.range(start + 1, 1, end, maxCols).value(chunk);
    batchCount += 1;
    if (onProgress) onProgress(batchCount / totalBatches);
    if (batchCount % UI_YIELD_EVERY_BATCHES === 0 || batchCount === totalBatches) await waitForNextFrame();
  }
}
function readSheetRows(sheet, maxRows = 200, maxCols = 60, valueOverrides = null) {
  const usedRange = sheet.usedRange();
  const rows = [];
  if (!usedRange) return { rows, merges: [] };
  const endRow = Math.min(usedRange.endCell().rowNumber(), maxRows);
  const endCol = Math.min(usedRange.endCell().columnNumber(), maxCols);
  for (let row = 1; row <= endRow; row += 1) {
    const cur = [];
    let hasValue = false;
    for (let col = 1; col <= endCol; col += 1) {
      const overrideKey = `${row},${col}`;
      const rawValue = valueOverrides && Object.prototype.hasOwnProperty.call(valueOverrides, overrideKey)
        ? valueOverrides[overrideKey]
        : sheet.cell(row, col).value();
      const value = normalizeCellValue(rawValue);
      if (value !== null && value !== undefined && String(value) !== '') hasValue = true;
      cur.push(value === undefined ? null : value);
    }
    if (hasValue || rows.length > 0) rows.push(cur);
  }
  const merges = [];
  try {
    const mc = sheet._mergeCells || sheet.mergeCells || {};
    const keys = Object.keys(mc || {});
    for (const key of keys) {
      const m = key.match(/^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$/);
      if (!m) continue;
      const sc = colLetterToIndex(m[1]);
      const sr = parseInt(m[2], 10) - 1;
      const ec = colLetterToIndex(m[3]);
      const er = parseInt(m[4], 10) - 1;
      if (sr > endRow - 1) continue;
      merges.push({ startRow: sr, startCol: sc, endRow: Math.min(er, endRow - 1), endCol: ec });
    }
  } catch { console.warn('无法读取合并单元格'); }
  return { rows, merges };
}

function toFiniteNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : NaN;
  if (typeof value === 'string' && NUMERIC_RE.test(value.trim())) {
    const parsed = Number(value.trim());
    return Number.isFinite(parsed) ? parsed : NaN;
  }
  return NaN;
}

function normalizeLookupKey(value) {
  const normalized = normalizeCellValue(value);
  return normalized === undefined || normalized === null ? '' : String(normalized).trim().toUpperCase();
}

// 从写入的 CSV 数据计算「6-QEPM原始数据」每一列的算术平均值（忽略空串与非数值），
// 用于复刻 Sheet5 里的 AVERAGE → Sheet3 里的 HLOOKUP 链路。
function computeQepmColumnStats(csvRows) {
  const header = (csvRows && csvRows[QEPM_HEADER_ROW_INDEX]) || [];
  const colCount = header.length;
  const sums = new Array(colCount).fill(0);
  const counts = new Array(colCount).fill(0);
  const total = Array.isArray(csvRows) ? csvRows.length : 0;
  for (let i = QEPM_DATA_START_INDEX; i < total; i += 1) {
    const row = csvRows[i];
    if (!row) continue;
    const n = Math.min(row.length, colCount);
    for (let c = 0; c < n; c += 1) {
      const value = toFiniteNumber(row[c]);
      if (!Number.isFinite(value)) continue;
      sums[c] += value;
      counts[c] += 1;
    }
  }
  return { header, sums, counts };
}

function averageForChannelSuffix(stats, channel, suffix) {
  const target = `${normalizeLookupKey(channel)}${String(suffix || '').toUpperCase()}`;
  if (!target) return null;
  const { header, sums, counts } = stats;
  for (let c = 0; c < header.length; c += 1) {
    const h = normalizeLookupKey(header[c]);
    if (h && h.endsWith(target)) return counts[c] > 0 ? sums[c] / counts[c] : null;
  }
  return null;
}

function readSheetCellText(sheet, row, col) {
  if (!sheet || !row || !col) return '';
  return String(normalizeCellValue(sheet.cell(row, col).value()) || '').trim();
}

function colLettersToNumber(letters) {
  const s = String(letters || '').toUpperCase().replace(/[^A-Z]/g, '');
  if (!s) return -1;
  let result = 0;
  for (let i = 0; i < s.length; i += 1) result = result * 26 + (s.charCodeAt(i) - 64);
  return result;
}

function previewFallbackCols() {
  return {
    channelCol: COL_L_CHANNEL,
    vavgCol: COL_M_VAVG,
    iavgCol: COL_N_IAVG,
    pavgCol: COL_O_PAVG,
    dcdcKeyCol: COL_P_DCDC_NAME,
    dcdcEffCol: COL_Q_DCDC_EFF,
    ldoKeyCol: COL_R_LDO_NAME,
    ldoEffCol: COL_S_LDO_EFF,
    totalEffCol: COL_T_TOTAL_EFF,
    powerAtVbattCol: COL_U_PWR_VBATT,
    includeInSumCol: COL_V_INCLUDE_SUM,
    powerInSumCol: COL_W_POWER_IN_SUM,
  };
}

function detectPreviewFormulaColumns(previewSheet) {
  const usedRange = previewSheet.usedRange();
  if (!usedRange) return previewFallbackCols();
  const endRow = Math.min(usedRange.endCell().rowNumber(), 40);
  const endCol = Math.min(usedRange.endCell().columnNumber(), 60);
  const info = previewFallbackCols();
  let foundAny = false;

  for (let r = 3; r <= endRow; r += 1) {
    for (let c = 1; c <= endCol; c += 1) {
      let formula = null;
      try { formula = previewSheet.cell(r, c).formula(); } catch { formula = null; }
      if (!formula || typeof formula !== 'string' || formula.charAt(0) !== '=') continue;

      let match = formula.match(/HLOOKUP\("\*"&([A-Z]+)\d+&"_V"/i);
      if (match) {
        info.vavgCol = c;
        info.channelCol = colLettersToNumber(match[1]);
        foundAny = true;
        continue;
      }
      match = formula.match(/HLOOKUP\("\*"&([A-Z]+)\d+&"_I"/i);
      if (match) {
        info.iavgCol = c;
        info.channelCol = colLettersToNumber(match[1]);
        foundAny = true;
        continue;
      }
      match = formula.match(/HLOOKUP\("\*"&([A-Z]+)\d+&"_P"/i);
      if (match) {
        info.pavgCol = c;
        info.channelCol = colLettersToNumber(match[1]);
        foundAny = true;
        continue;
      }
      match = formula.match(/^=IF\(([A-Z]+)\d+="-",1,VLOOKUP\(\1\d+,'2_电源信息表'!\$B\$\d+:\$M\$\d+,12,FALSE\)\)$/i);
      if (match) {
        info.dcdcEffCol = c;
        info.dcdcKeyCol = colLettersToNumber(match[1]);
        foundAny = true;
        continue;
      }
      match = formula.match(/^=IF\(([A-Z]+)\d+="-",1,VLOOKUP\(\1\d+,'2_电源信息表'!\$B\$\d+:\$[IM]\$\d+,(8|12),FALSE\)\)$/i);
      if (match) {
        info.ldoEffCol = c;
        info.ldoKeyCol = colLettersToNumber(match[1]);
        foundAny = true;
        continue;
      }
      match = formula.match(/^=([A-Z]+)\d+\*([A-Z]+)\d+$/i);
      if (match) {
        info.totalEffCol = c;
        foundAny = true;
        continue;
      }
      match = formula.match(/^=([A-Z]+)\d+\/([A-Z]+)\d+$/i);
      if (match) {
        const numeratorCol = colLettersToNumber(match[1]);
        const denominatorCol = colLettersToNumber(match[2]);
        if (numeratorCol > 0 && denominatorCol > 0) {
          info.powerAtVbattCol = c;
          info.totalEffCol = denominatorCol;
          foundAny = true;
        }
        continue;
      }
      match = formula.match(/^=IF\(([A-Z]+)\d+="Y",([A-Z]+)\d+,0\)$/i);
      if (match) {
        info.powerInSumCol = c;
        info.includeInSumCol = colLettersToNumber(match[1]);
        info.powerAtVbattCol = colLettersToNumber(match[2]);
        foundAny = true;
      }
    }
  }
  return foundAny ? info : previewFallbackCols();
}

function buildExactLookupMap(sheet, startRow, endRow, keyCol, valueCol) {
  const map = new Map();
  if (!sheet) return map;
  for (let row = startRow; row <= endRow; row += 1) {
    const key = normalizeLookupKey(sheet.cell(row, keyCol).value());
    if (!key || map.has(key)) continue;
    map.set(key, sheet.cell(row, valueCol).value());
  }
  return map;
}

function interpolateSeries(x, xs, ys) {
  const points = [];
  for (let i = 0; i < xs.length; i += 1) {
    const px = toFiniteNumber(xs[i]);
    const py = toFiniteNumber(ys[i]);
    if (Number.isFinite(px) && Number.isFinite(py)) points.push({ x: px, y: py });
  }
  if (!points.length) return NaN;
  if (points.length === 1) return points[0].y;
  if (!Number.isFinite(x)) return NaN;
  if (x <= points[0].x) return points[0].y;
  for (let i = 1; i < points.length; i += 1) {
    if (x <= points[i].x) {
      const left = points[i - 1];
      const right = points[i];
      if (right.x === left.x) return right.y;
      return left.y + ((x - left.x) / (right.x - left.x)) * (right.y - left.y);
    }
  }
  const left = points[points.length - 2];
  const right = points[points.length - 1];
  if (right.x === left.x) return right.y;
  return left.y + ((x - left.x) / (right.x - left.x)) * (right.y - left.y);
}

function buildDcdcEfficiencyMap(sheet, currentByRail) {
  const map = new Map();
  if (!sheet) return map;
  const loads = [];
  for (let col = 5; col <= 11; col += 1) loads.push(sheet.cell(8, col).value());
  for (let row = 9; row <= 31; row += 1) {
    const key = normalizeLookupKey(sheet.cell(row, 2).value());
    if (!key || map.has(key)) continue;
    const efficiencies = [];
    for (let col = 5; col <= 11; col += 1) efficiencies.push(sheet.cell(row, col).value());
    const load = currentByRail[key] || 0;
    let eff = interpolateSeries(load, loads, efficiencies);
    if (!Number.isFinite(eff)) eff = toFiniteNumber(sheet.cell(row, 13).value());
    if (Number.isFinite(eff)) map.set(key, eff);
  }
  return map;
}

function readPreviewFormula(sheet, row, col) {
  if (!sheet || !row || !col) return '';
  try {
    const formula = sheet.cell(row, col).formula();
    return typeof formula === 'string' ? formula : '';
  } catch {
    return '';
  }
}

function buildPreviewValueOverrides(workbook, previewSheet, csvRows) {
  const usedRange = previewSheet.usedRange();
  if (!usedRange) return null;
  const lastRow = usedRange.endCell().rowNumber();
  const stats = computeQepmColumnStats(csvRows);
  const cols = detectPreviewFormulaColumns(previewSheet);
  const overrides = {};
  const rowStates = [];

  // 模板前两行是表头，真实数据从第 3 行开始；这里必须从第 3 行起算，
  // 否则首条数据会落回模板缓存值，表现成“幽灵数据”。
  for (let row = 3; row <= lastRow; row += 1) {
    const vFormula = readPreviewFormula(previewSheet, row, cols.vavgCol);
    if (!vFormula || !/HLOOKUP/i.test(vFormula)) continue;

    const channel = readSheetCellText(previewSheet, row, cols.channelCol);
    const vAvg = averageForChannelSuffix(stats, channel, '_V');
    const iAvg = averageForChannelSuffix(stats, channel, '_I');
    const pAvg = averageForChannelSuffix(stats, channel, '_P');
    const vVal = vAvg === null ? 0 : vAvg / 1000;
    const iVal = iAvg === null ? 0 : iAvg;
    const pVal = pAvg === null ? 0 : pAvg;

    overrides[`${row},${cols.vavgCol}`] = vVal;
    overrides[`${row},${cols.iavgCol}`] = iVal;
    overrides[`${row},${cols.pavgCol}`] = pVal;

    rowStates.push({
      row,
      iVal,
      pVal,
      dcdcKey: normalizeLookupKey(readSheetCellText(previewSheet, row, cols.dcdcKeyCol)),
      ldoKey: normalizeLookupKey(readSheetCellText(previewSheet, row, cols.ldoKeyCol)),
      includeInSum: normalizeLookupKey(readSheetCellText(previewSheet, row, cols.includeInSumCol)) === 'Y',
      sFormula: readPreviewFormula(previewSheet, row, cols.ldoEffCol),
    });
  }

  if (!rowStates.length) return null;

  const currentByRail = {};
  rowStates.forEach(({ dcdcKey, iVal }) => {
    if (!dcdcKey || dcdcKey === '-') return;
    currentByRail[dcdcKey] = (currentByRail[dcdcKey] || 0) + iVal;
  });

  const powerSheet = workbook.sheet('2_电源信息表');
  const dcdcEffMap = buildDcdcEfficiencyMap(powerSheet, currentByRail);
  const cachedDcdcEffMap = buildExactLookupMap(powerSheet, 9, 31, 2, 13);
  const ldoEffMap = buildExactLookupMap(powerSheet, 1, 167, 2, 9);

  rowStates.forEach(({ row, pVal, dcdcKey, ldoKey, includeInSum, sFormula }) => {
    let qVal;
    if (!dcdcKey) qVal = toFiniteNumber(previewSheet.cell(row, cols.dcdcEffCol).value());
    else if (dcdcKey === '-') qVal = 1;
    else {
      qVal = toFiniteNumber(dcdcEffMap.get(dcdcKey));
      if (!Number.isFinite(qVal)) qVal = toFiniteNumber(cachedDcdcEffMap.get(dcdcKey));
    }
    if (!Number.isFinite(qVal)) qVal = 0;

    let sVal;
    if (!ldoKey) sVal = toFiniteNumber(previewSheet.cell(row, cols.ldoEffCol).value());
    else if (ldoKey === '-') sVal = 1;
    else if (/\$B\$9:\$M\$31,12,FALSE/i.test(sFormula)) {
      sVal = toFiniteNumber(dcdcEffMap.get(ldoKey));
      if (!Number.isFinite(sVal)) sVal = toFiniteNumber(cachedDcdcEffMap.get(ldoKey));
    } else {
      sVal = toFiniteNumber(ldoEffMap.get(ldoKey));
    }
    if (!Number.isFinite(sVal)) sVal = 0;

    const totalEff = qVal * sVal;
    const powerAtVbatt = Number.isFinite(totalEff) && totalEff !== 0 ? pVal / totalEff : 0;
    const powerInSum = includeInSum ? powerAtVbatt : 0;

    overrides[`${row},${cols.dcdcEffCol}`] = qVal;
    overrides[`${row},${cols.ldoEffCol}`] = sVal;
    overrides[`${row},${cols.totalEffCol}`] = totalEff;
    overrides[`${row},${cols.powerAtVbattCol}`] = powerAtVbatt;
    overrides[`${row},${cols.powerInSumCol}`] = powerInSum;
  });

  return overrides;
}

// 兜底：在 workbook.xml 的 <calcPr> 上设置 fullCalcOnLoad="1"，
// 使 Excel/WPS 打开文件时强制对仍保留的公式做一次全量重算。
function forceFullCalcOnLoad(workbook) {
  try {
    const node = workbook._node;
    if (!node || !Array.isArray(node.children)) return;
    let calcPr = node.children.find((k) => k && k.name === 'calcPr');
    if (!calcPr) {
      calcPr = { name: 'calcPr', attributes: {}, children: [] };
      node.children.push(calcPr);
    }
    if (!calcPr.attributes) calcPr.attributes = {};
    calcPr.attributes.fullCalcOnLoad = 1;
  } catch (e) {
    console.warn('设置 fullCalcOnLoad 失败（不影响主流程）：', e);
  }
}

async function buildDatasetResult(ds, onProgress) {
  // Progress phases: loading=0-10%, writing=10-65%, computing=65-75%, output=75-100%
  if (onProgress) onProgress(2, '加载模板...');
  // Resolve the project template for THIS dataset. Prefer the dataset's own
  // per-file association; fall back to the global active project as the default.
  const projectId = ds.projectId || (state.activeProject ? state.activeProject.id : null);
  if (!projectId) throw new Error(`文件「${ds.fileName}」尚未选择项目模板，请先为其指定项目。`);
  const workbook = await loadTemplateWorkbook(projectId);
  const targetSheet = workbook.sheet(TARGET_SHEET_NAME);
  if (!targetSheet) throw new Error(`模板中未找到工作表：${TARGET_SHEET_NAME}`);
  if (onProgress) onProgress(10, '写入数据...');
  await writeCsvRowsToSheet(targetSheet, ds.csvRows, (batchPct) => {
    if (onProgress) onProgress(10 + batchPct * 55, '写入数据...');
  });
  const previewSheet = workbook.sheet(PREVIEW_SHEET_NAME);
  if (!previewSheet) throw new Error(`模板中未找到预览工作表：${PREVIEW_SHEET_NAME}`);
  if (onProgress) onProgress(67, '计算预览结果...');
  await waitForNextFrame();
  const previewValueOverrides = buildPreviewValueOverrides(workbook, previewSheet, ds.csvRows);
  forceFullCalcOnLoad(workbook);
  if (onProgress) onProgress(75, '读取预览...');
  await waitForNextFrame();
  const { rows, merges } = readSheetRows(previewSheet, 200, 60, previewValueOverrides);
  if (onProgress) onProgress(80, '生成文件...');
  await waitForNextFrame();
  ds.workbookBlob = await workbook.outputAsync();
  ds.processedRows = rows;
  ds.processedMerges = merges;
  // Persist the project that was actually used to build this dataset's result.
  ds.projectId = projectId;
  const usedItem = state.templateItems.find((x) => x.id === projectId);
  ds.projectName = usedItem
    ? usedItem.name
    : (state.activeProject && state.activeProject.id === projectId
        ? state.activeProject.name
        : (ds.projectName || '未命名项目'));
  if (onProgress) onProgress(100, '完成');
}
function buildHistoryDatasetPayload(ds) {
  const csvRows = Array.isArray(ds.csvRows) ? ds.csvRows : [];
  const estimatedCsvSize = estimateJsonSize(csvRows);
  const shouldCacheCsvRows = estimatedCsvSize <= HISTORY_CSV_CACHE_LIMIT_BYTES;
  const limitedCsvRows = shouldCacheCsvRows ? csvRows : [];
  return {
    id: ds.id,
    fileName: ds.fileName,
    projectId: ds.projectId,
    projectName: ds.projectName,
    processedRows: (ds.processedRows || []).map((row) => row.slice()),
    processedMerges: (ds.processedMerges || []).map((merge) => ({ ...merge })),
    workbookBlob: ds.workbookBlob,
    csvHeaderRowIndex: ds.csvHeaderRowIndex,
    csvDataStartRow: ds.csvDataStartRow,
    csvRows: limitedCsvRows,
    rawCsvCached: shouldCacheCsvRows,
    rawCsvOmittedReason: shouldCacheCsvRows ? '' : `原始 CSV 估算约 ${formatSize(estimatedCsvSize)}，超过 ${formatSize(HISTORY_CSV_CACHE_LIMIT_BYTES)}，已跳过缓存`,
    csvRowCount: csvRows.length,
    cachedCsvRowCount: shouldCacheCsvRows ? csvRows.length : Math.min(csvRows.length, HISTORY_CSV_ROW_SOFT_LIMIT),
  };
}
function getHistoryCacheSummary(datasets) {
  const skipped = datasets.filter((ds) => !ds.rawCsvCached);
  if (!skipped.length) return '已缓存原始 CSV、结果预览和 XLSX，可直接回溯、预览和下载。';
  const names = skipped.map((ds) => ds.fileName || '未命名文件').join('、');
  return `以下数据因体积过大未缓存原始 CSV：${names}；仍可回溯预览/下载 XLSX，但不能重新绘制波形图。`;
}
async function persistCurrentResultsToHistory(targets) {
  if (!targets.length) return null;
  const historyStoreId = `hs_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  const datasets = targets.slice(0, BATCH_MAX).map(buildHistoryDatasetPayload);
  const entry = {
    id: historyStoreId,
    createdAt: Date.now(),
    projectId: state.activeProject ? state.activeProject.id : '',
    projectName: state.activeProject ? state.activeProject.name : '未命名项目',
    datasets,
  };
  await saveHistoryEntry(entry);
  return {
    historyStoreId,
    historyCacheSummary: getHistoryCacheSummary(datasets),
  };
}
async function restoreHistoryWorkspace(id) {
  const item = getHistory().find((record) => record.id === id);
  if (!item || !item.historyStoreId) {
    alert('该历史条目没有缓存结果，只能查看留痕信息。');
    return;
  }
  const entry = await getHistoryEntry(item.historyStoreId);
  if (!entry || !Array.isArray(entry.datasets) || !entry.datasets.length) {
    alert('未找到该历史条目的缓存结果，可能已被清理。');
    return;
  }
  const restoredDatasets = entry.datasets.slice(0, BATCH_MAX).map((ds, index) => ({
    id: ds.id || genDatasetId(),
    fileName: ds.fileName,
    csvRows: Array.isArray(ds.csvRows) ? ds.csvRows.map((row) => row.slice()) : [],
    csvHeaderRowIndex: Number.isInteger(ds.csvHeaderRowIndex) ? ds.csvHeaderRowIndex : -1,
    csvDataStartRow: Number.isInteger(ds.csvDataStartRow) ? ds.csvDataStartRow : -1,
    processedRows: (ds.processedRows || []).map((row) => row.slice()),
    processedMerges: (ds.processedMerges || []).map((merge) => ({ ...merge })),
    workbookBlob: ds.workbookBlob || null,
    projectId: ds.projectId || entry.projectId || null,
    projectName: ds.projectName || entry.projectName || '未命名项目',
    rawCsvCached: ds.rawCsvCached !== false,
    rawCsvOmittedReason: ds.rawCsvOmittedReason || '',
    rawCsvMissing: ds.rawCsvCached === false,
    historyRestored: true,
    historyStoreId: entry.id,
    historyDatasetIndex: index,
  }));
  const activeProjectRecord = entry.projectId
    ? (state.templateItems.find((x) => x.id === entry.projectId) || null)
    : null;
  state.activeProject = activeProjectRecord ? Object.assign({}, activeProjectRecord) : (entry.projectId || entry.projectName ? {
    id: entry.projectId || `history_project_${entry.id}`,
    name: entry.projectName || '历史项目',
  } : null);
  state.datasets = restoredDatasets;
  state.activeDatasetId = restoredDatasets[0].id;
  resetCompareGroups();
  state.csvSearchKeyword = '';
  state.csvSearchMode = 'manual';
  state.csvMatchedCols = [];
  state.csvChartSelectedCols = [];
  if (els.csvColSearch) els.csvColSearch.value = '';
  clearChart();
  syncActiveDatasetToState();
  updateMatchInfoAndPicker();
  renderDatasetTabs();
  renderProcessedPreview();
  renderPreview();
  updateCsvFileDisplay();
  updateSummary();
  const missingCsv = restoredDatasets.filter((ds) => ds.rawCsvMissing);
  const prefix = `已回溯到 ${formatDateTime(item.time)} 的历史结果，当前工作区已切换。`;
  if (missingCsv.length) {
    alert(`${prefix}\n\n以下数据未缓存原始 CSV，仅可预览/下载 XLSX，不能重新绘制波形：${missingCsv.map((ds) => ds.fileName).join('、')}`);
  } else {
    alert(prefix);
  }
  scrollToEl(els.processedSection);
}
async function downloadHistoryEntryResults(id) {
  const item = getHistory().find((record) => record.id === id);
  if (!item || !item.historyStoreId) {
    alert('该历史条目没有缓存结果，无法直接下载。');
    return;
  }
  const entry = await getHistoryEntry(item.historyStoreId);
  if (!entry || !Array.isArray(entry.datasets) || !entry.datasets.length) {
    alert('未找到该历史条目的缓存结果，可能已被清理。');
    return;
  }
  for (let i = 0; i < entry.datasets.length; i += 1) {
    const ds = entry.datasets[i];
    if (!ds.workbookBlob) continue;
    downloadBlob(ds.workbookBlob, ds.fileName, ds.projectName || entry.projectName);
    if (entry.datasets.length > 1) await new Promise((resolve) => setTimeout(resolve, 400));
  }
}

async function ensureActiveProcessed() {
  const ds = getActiveDataset();
  if (!ds || !ds.csvRows || !ds.csvRows.length) throw new Error('请先上传 CSV。');
  if (!ds.projectId && !state.activeProject) throw new Error('请先为该文件选择项目模板。');
  if (!ds.processedRows || !ds.processedRows.length || !ds.workbookBlob) {
    await buildDatasetResult(ds);
    syncActiveDatasetToState();
    renderPreview();
    updateSummary();
  }
}

function downloadBlob(blob, sourceFileName, projectName) {
  const pName = projectName || (state.activeProject ? state.activeProject.name : '');
  const projectPart = pName ? `_${pName}` : '';
  const baseName = (sourceFileName || 'qepm_data').replace(/\.csv$/i, '');
  const outputName = `${baseName}${projectPart}_已写入模板.xlsx`;
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = outputName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function runExport() {
  if (!hasReadyInput()) return;
  const targets = state.datasets.filter((d) => d.csvRows && d.csvRows.length);
  const total = targets.length;
  els.exportBtn.disabled = true;
  els.downloadTip.textContent = '正在基于所选模板处理数据，请稍候。';

  // Show progress bar
  const progressWrap = document.createElement('div');
  progressWrap.className = 'export-progress-wrap';
  progressWrap.innerHTML = `
    <div class="export-progress-bar"><div class="export-progress-fill" style="width:0%"></div></div>
    <div class="export-progress-text">准备中...</div>
  `;
  els.exportBtn.innerHTML = '';
  els.exportBtn.appendChild(progressWrap);

  const fillEl = progressWrap.querySelector('.export-progress-fill');
  const textEl = progressWrap.querySelector('.export-progress-text');

  function setProgress(pct, label) {
    fillEl.style.width = `${Math.min(100, Math.round(pct))}%`;
    textEl.textContent = label || `${Math.round(pct)}%`;
  }

  try {
    for (let i = 0; i < total; i++) {
      const ds = targets[i];
      const baseProgress = (i / total) * 90; // 0-90% for processing
      const sliceSize = 90 / total; // each dataset takes this % of the bar
      setProgress(baseProgress, `处理 ${i + 1}/${total}：${ds.fileName.slice(0, 20)}...`);
      await waitForNextFrame();
      await buildDatasetResult(ds, (innerPct, phase) => {
        const pct = baseProgress + (innerPct / 100) * sliceSize;
        setProgress(pct, `${i + 1}/${total} ${phase || ''} ${Math.round(innerPct)}%`);
      });
      setProgress(baseProgress + sliceSize, `已完成 ${i + 1}/${total}`);
      await waitForNextFrame();
    }
    setProgress(92, '保存历史记录...');
    await waitForNextFrame();
    syncActiveDatasetToState();
    const historyCache = await persistCurrentResultsToHistory(targets);
    await addHistoryEntry({
      action: total > 1 ? `批量结果导出（${total} 组）` : '结果导出',
      fileName: total > 1 ? `${targets[0].fileName} 等 ${total} 个文件` : targets[0].fileName,
      projectName: state.activeProject ? state.activeProject.name : '',
      rowCount: targets.reduce((sum, ds) => sum + getRowCount(ds.csvRows), 0),
      colCount: Math.max(...targets.map((ds) => getColCount(ds.csvRows)), 0),
      datasetCount: total,
      historyStoreId: historyCache ? historyCache.historyStoreId : '',
      historyCacheSummary: historyCache ? historyCache.historyCacheSummary : '',
    });
    setProgress(100, '✓ 导出完成');
    await waitForNextFrame();
    els.mainStatus.textContent = '已生成结果';
    els.downloadTip.textContent = total > 1
      ? `已完成 ${total} 组数据处理。已自动打开结果预览，并生成对比区快照。`
      : `已完成处理：CSV 已写入"${TARGET_SHEET_NAME}"，预览取自"${PREVIEW_SHEET_NAME}"。`;
    els.resultActions.classList.remove('hidden');
    els.exportBtn.innerHTML = '<span class="material-symbols-outlined">refresh</span> 重新生成结果';
    renderProcessedPreview();
    scrollToEl(els.processedSection);
    if (state.datasets.length >= 2) {
      replaceCompareSnapshotsWithAllProcessed();
      renderCompareAuto();
      setTimeout(() => scrollToEl(els.compareArea), 180);
    }
    renderPreview();
    updateSummary();
  } catch (error) {
    console.error(error);
    alert(error.message || '处理失败，请稍后重试。');
    els.downloadTip.textContent = '处理失败，请检查项目模板和 CSV。';
    els.exportBtn.innerHTML = '<span class="material-symbols-outlined">play_arrow</span> 结果导出';
  } finally {
    els.exportBtn.disabled = !hasReadyInput();
  }
}

function replaceCompareSnapshotsWithAllProcessed() {
  const done = state.datasets.filter((d) => d.processedRows && d.processedRows.length).slice(0, COMPARE_MAX);
  const group = getDefaultCompareGroup();
  group.snapshots = done.map((ds) => makeSnapshotFromDataset(ds));
  renderCompareArea();
  renderDatasetSelectPanel();
  return done.length;
}
function renderCompareAuto() {
  renderCompareArea();
}
function scrollToPreview() {
  scrollToEl(els.processedSection || document.querySelector('.processed-table'), 'start');
}
async function previewProcessed() {
  try { await ensureActiveProcessed(); scrollToPreview(); }
  catch (error) { console.error(error); alert(error.message || '预览失败，请稍后重试。'); }
}
async function downloadProcessed() {
  els.downloadBtn.disabled = true;
  const original = els.downloadBtn.innerHTML;
  els.downloadBtn.innerHTML = '<span class="material-symbols-outlined animate-spin">progress_activity</span> 正在下载...';
  try {
    await ensureActiveProcessed();
    const ds = getActiveDataset();
    downloadBlob(ds.workbookBlob, ds.fileName, ds.projectName);
    await addHistoryEntry({
      action: '导出下载', fileName: ds.fileName, projectName: ds.projectName,
      rowCount: getRowCount(ds.csvRows), colCount: getColCount(ds.csvRows),
    });
    els.mainStatus.textContent = '导出成功';
  } catch (error) {
    console.error(error);
    alert(error.message || '导出失败，请稍后重试。');
  } finally {
    els.downloadBtn.disabled = false;
    els.downloadBtn.innerHTML = original;
  }
}
async function downloadAllProcessed() {
  const targets = state.datasets.filter((d) => d.csvRows && d.csvRows.length);
  if (!targets.length) return;
  els.downloadAllBtn.disabled = true;
  const original = els.downloadAllBtn.innerHTML;
  try {
    for (let i = 0; i < targets.length; i++) {
      const ds = targets[i];
      els.downloadAllBtn.innerHTML = `<span class="material-symbols-outlined animate-spin">progress_activity</span> 下载中 ${i + 1}/${targets.length}`;
      if (!ds.workbookBlob) await buildDatasetResult(ds);
      downloadBlob(ds.workbookBlob, ds.fileName, ds.projectName);
      await new Promise((r) => setTimeout(r, 400));
    }
    syncActiveDatasetToState();
    await addHistoryEntry({
      action: `批量下载 ${targets.length} 组`, fileName: `${targets.length} 个文件`,
      projectName: state.activeProject ? state.activeProject.name : '', rowCount: 0, colCount: 0,
    });
    els.mainStatus.textContent = '导出成功';
  } catch (error) {
    console.error(error);
    alert(error.message || '批量下载失败，请稍后重试。');
  } finally {
    els.downloadAllBtn.disabled = false;
    els.downloadAllBtn.innerHTML = original;
    renderPreview();
    updateSummary();
  }
}
async function downloadDataset(id) {
  const ds = state.datasets.find((d) => d.id === id);
  if (!ds) return;
  try {
    if (!ds.workbookBlob) {
      if (!ds.projectId && !state.activeProject) { alert('请先为该文件选择项目模板。'); return; }
      await buildDatasetResult(ds);
      syncActiveDatasetToState();
      renderPreview();
      updateSummary();
    }
    downloadBlob(ds.workbookBlob, ds.fileName, ds.projectName);
  } catch (error) { console.error(error); alert(error.message || '下载失败。'); }
}

// ---------- Compare: helpers ----------
function getProcessedDataStart(rows, powerCol) {
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r && isNumericLike(r[powerCol])) return i;
  }
  return Math.min(1, Math.max(0, rows.length - 1));
}
function processedHeaderLabel(rows, dataStart, col) {
  for (let i = dataStart - 1; i >= 0; i--) {
    const v = rows[i] && rows[i][col];
    if (v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
  }
  return indexToColLetter(col);
}
function detectPowerCol(rows, valueCols) {
  if (!valueCols.length) return -1;
  const probe = valueCols[valueCols.length - 1];
  const dataStart = getProcessedDataStart(rows, probe);
  for (const c of valueCols) {
    const label = processedHeaderLabel(rows, dataStart, c);
    if (/pavg|power|功率|\(mw\)/i.test(label)) return c;
  }
  if (valueCols.includes(POWER_COL_DEFAULT)) return POWER_COL_DEFAULT;
  return valueCols[valueCols.length - 1];
}

// ---------- Compare groups ----------
function genCompareGroupId() {
  return `grp_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 6)}`;
}
function makeCompareGroup(name) {
  return { id: genCompareGroupId(), name: name || '对比组', snapshots: [], highlightDiff: false };
}
function getCompareGroup(groupId) {
  return state.compareGroups.find((g) => g.id === groupId) || null;
}
function nextCompareGroupName() {
  const names = new Set(state.compareGroups.map((g) => g.name));
  let n = 1;
  while (names.has(`对比组 ${n}`)) n += 1;
  return `对比组 ${n}`;
}
function getDefaultCompareGroup() {
  if (!state.compareGroups.length) state.compareGroups.push(makeCompareGroup('对比组 1'));
  return state.compareGroups[0];
}
function resetCompareGroups() {
  state.compareGroups = [makeCompareGroup('对比组 1')];
  state.crossProjectChannelSelection = {};
  Object.values(state.crossProjectChartInstances || {}).forEach((chart) => { try { chart.destroy(); } catch (e) { /* noop */ } });
  state.crossProjectChartInstances = {};
}
function addCompareGroup() {
  const g = makeCompareGroup(nextCompareGroupName());
  state.compareGroups.push(g);
  renderCompareArea();
  renderDatasetSelectPanel();
  syncPreviewSectionVisibility();
  scrollToEl(els.compareArea, 'nearest');
  return g;
}
function removeCompareGroup(groupId) {
  const idx = state.compareGroups.findIndex((g) => g.id === groupId);
  if (idx < 0) return;
  const g = state.compareGroups[idx];
  if (g.snapshots.length && !confirm(`确定要删除「${g.name}」吗？该组包含 ${g.snapshots.length} 份数据。`)) return;
  state.compareGroups.splice(idx, 1);
  clearCrossProjectChannelSelection(groupId);
  renderCompareArea();
  renderDatasetSelectPanel();
  syncPreviewSectionVisibility();
}
function clearCompareGroup(groupId) {
  const g = getCompareGroup(groupId);
  if (!g || !g.snapshots.length) return;
  if (!confirm(`确定要清空「${g.name}」中的所有数据吗？`)) return;
  g.snapshots = [];
  clearCrossProjectChannelSelection(groupId);
  renderCompareArea();
  renderDatasetSelectPanel();
}
function renameCompareGroup(groupId) {
  const g = getCompareGroup(groupId);
  if (!g) return;
  const name = prompt('重命名对比组：', g.name);
  if (name === null) return;
  const trimmed = name.trim();
  if (!trimmed) return;
  g.name = trimmed;
  renderCompareArea();
  renderDatasetSelectPanel();
}
function findSnapshotById(snapId) {
  for (const g of state.compareGroups) {
    const s = g.snapshots.find((x) => x.id === snapId);
    if (s) return s;
  }
  return null;
}

// ---------- Compare snapshots ----------
function makeSnapshotFromDataset(ds) {
  const colIndices = resolveProcessedColIndices(ds.processedRows);
  const dataRows = ds.processedRows.slice(1);
  const numericMap = computeNumericColumnMap(dataRows, colIndices);
  const numericObj = {};
  colIndices.forEach((ci) => { numericObj[ci] = !!numericMap.get(ci); });
  const projectName = ds.projectName || (state.activeProject ? state.activeProject.name : '未命名项目');
  return {
    id: `snap_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
    datasetId: ds.id,
    projectId: ds.projectId || (state.activeProject ? state.activeProject.id : null),
    projectName,
    csvName: ds.fileName || '(未命名 CSV)',
    timestamp: Date.now(),
    colIndices: colIndices.slice(),
    rows: ds.processedRows.map((r) => r.slice()),
    merges: (ds.processedMerges || []).map((m) => ({ ...m })),
    numeric: numericObj,
    csvRows: ds.csvRows,
    csvHeaderRowIndex: ds.csvHeaderRowIndex,
    csvDataStartRow: ds.csvDataStartRow,
    rawCsvMissing: !!ds.rawCsvMissing,
  };
}
function snapshotCurrentProcessed() {
  const ds = getActiveDataset();
  if (!ds || !ds.processedRows || !ds.processedRows.length) { alert('当前数据组还没有结果，请先点击"结果导出"。'); return; }
  const group = getDefaultCompareGroup();
  if (group.snapshots.some((s) => s.datasetId === ds.id)) { alert(`该数据组已在「${group.name}」中。`); scrollToEl(els.compareArea, 'nearest'); return; }
  if (group.snapshots.length >= COMPARE_MAX) { alert(`「${group.name}」最多同时对比 ${COMPARE_MAX} 份，请先移除一份。`); return; }
  group.snapshots.push(makeSnapshotFromDataset(ds));
  renderCompareArea();
  renderDatasetSelectPanel();
  syncPreviewSectionVisibility();
  scrollToEl(els.compareArea, 'nearest');
}
function snapshotAllProcessed() {
  const done = state.datasets.filter((d) => d.processedRows && d.processedRows.length);
  if (!done.length) { alert('没有已生成结果的数据组，请先"结果导出"。'); return; }
  const group = getDefaultCompareGroup();
  let added = 0;
  for (const ds of done) {
    if (group.snapshots.length >= COMPARE_MAX) break;
    if (group.snapshots.some((s) => s.datasetId === ds.id)) continue;
    group.snapshots.push(makeSnapshotFromDataset(ds));
    added += 1;
  }
  renderCompareArea();
  renderDatasetSelectPanel();
  syncPreviewSectionVisibility();
  scrollToEl(els.compareArea, 'nearest');
  if (group.snapshots.length >= COMPARE_MAX && added < done.length) {
    alert(`「${group.name}」最多 ${COMPARE_MAX} 份，已加入 ${group.snapshots.length} 组。`);
  }
}
function removeSnapshot(id) {
  for (const g of state.compareGroups) {
    const before = g.snapshots.length;
    g.snapshots = g.snapshots.filter((s) => s.id !== id);
    if (g.snapshots.length !== before) {
      const sel = state.crossProjectChannelSelection[g.id] || {};
      Object.keys(sel).forEach((key) => { if (sel[key] && sel[key].snapId === id) delete sel[key]; });
      closeCrossProjectChannelChart(g.id);
      break;
    }
  }
  renderCompareArea();
  renderDatasetSelectPanel();
}
function clearAllSnapshots() {
  const group = getDefaultCompareGroup();
  if (!group.snapshots.length) return;
  if (!confirm(`确定要清空「${group.name}」中的所有快照吗？`)) return;
  group.snapshots = [];
  renderCompareArea();
  renderDatasetSelectPanel();
}

function computeCompareDiffMap(snaps, highlightDiff) {
  const result = new Map();
  snaps = snaps || [];
  if (snaps.length < 2 || !highlightDiff) {
    for (const s of snaps) result.set(s.id, new Set());
    return result;
  }
  const baseCols = snaps[0].colIndices;
  const maxRow = Math.max(...snaps.map((s) => s.rows.length));
  for (const s of snaps) result.set(s.id, new Set());
  for (let r = 0; r < maxRow; r++) {
    for (const ci of baseCols) {
      const vals = [];
      let valid = true;
      for (const s of snaps) {
        if (!s.colIndices.includes(ci)) { valid = false; break; }
        const row = s.rows[r];
        if (!row) { valid = false; break; }
        const n = toNumberOrNull(row[ci]);
        if (n === null) { valid = false; break; }
        vals.push(Number(n.toFixed(2)));
      }
      if (!valid) continue;
      const first = vals[0];
      if (vals.some((v) => v !== first)) {
        for (const s of snaps) result.get(s.id).add(`${r}_${ci}`);
      }
    }
  }
  return result;
}

// Click on a highlighted compare cell -> locate the channel in that snapshot's CSV and chart it.
function ensureDatasetForSnapshot(snap) {
  let ds = state.datasets.find((d) => d.id === snap.datasetId);
  if (ds) return ds;
  ds = {
    id: snap.datasetId,
    fileName: snap.csvName,
    csvRows: snap.csvRows,
    csvHeaderRowIndex: snap.csvHeaderRowIndex,
    csvDataStartRow: snap.csvDataStartRow,
    processedRows: snap.rows,
    processedMerges: snap.merges,
    workbookBlob: null,
    projectId: snap.projectId,
    projectName: snap.projectName,
    rawCsvMissing: !!snap.rawCsvMissing,
  };
  state.datasets.push(ds);
  return ds;
}
function handleCompareChartClick(snapId, channel) {
  const snap = findSnapshotById(snapId);
  if (!snap) { alert('对比快照已被移除。'); return; }
  const ds = ensureDatasetForSnapshot(snap);
  locateChannelInCsvAndChart(ds.id, channel);
}

// Find a channel column by metric type (current / voltage / power)
function findChannelColByMetric(rows, hdrIdx, channel, metric) {
  const header = rows[hdrIdx] || [];
  const chLower = String(channel || '').toLowerCase().trim();
  const cand = [];
  for (let c = 0; c < header.length; c++) {
    const h = String(header[c] || '').toLowerCase();
    if (h && h.includes(chLower)) cand.push(c);
  }
  const regex = METRIC_REGEX[metric];
  if (regex) {
    const matched = cand.find((c) => regex.test(String(header[c] || '')));
    if (matched !== undefined) return matched;
  }
  return -1;
}
function compareChannelWaveformsAcrossSnapshots(channel, groupId) {
  const group = (groupId && getCompareGroup(groupId)) || getDefaultCompareGroup();
  const snaps = group ? group.snapshots : [];
  const ch = String(channel || '').trim();
  if (!snaps.length) return;
  if (!ch || ch === '-') { alert('该行没有有效的采样通道编号，无法对比波形。'); return; }

  const activeMetric = state.channelMetricType || METRIC_CURRENT;
  const seriesList = [];
  let refLabels = null;
  let maxLen = 0;
  snaps.forEach((snap, i) => {
    const rows = snap.csvRows;
    if (!rows || !rows.length) return;
    let hdrIdx = snap.csvHeaderRowIndex;
    let dataStart = snap.csvDataStartRow;
    if (hdrIdx < 0) { const d = detectHeaderRow(rows); hdrIdx = d.headerRowIndex; dataStart = d.dataStartRow; }
    const targetC = findChannelColByMetric(rows, hdrIdx, ch, activeMetric);
    if (targetC < 0) return;
    const data = rows.slice(dataStart >= 0 ? dataStart : 1);
    const total = data.length;
    const step = Math.max(1, Math.ceil(total / CHART_MAX_POINTS));
    const ys = [];
    const xs = [];
    for (let r = 0; r < total; r += step) {
      const row = data[r];
      const v = (row[targetC] || '').toString().trim();
      let num = null;
      if (v !== '' && NUMERIC_RE.test(v)) { const n = Number(v); if (Number.isFinite(n)) num = n; }
      ys.push(num);
      xs.push((row[0] || '').toString().trim());
    }
    if (ys.length > maxLen) { maxLen = ys.length; refLabels = xs; }
    const color = DS_COLORS[i % DS_COLORS.length];
    seriesList.push({
      label: `#${i + 1} ${snap.csvName}`,
      data: ys,
      borderColor: color,
      backgroundColor: color + '33',
      pointRadius: 0,
      borderWidth: 1.6,
      tension: 0.15,
      spanGaps: true,
    });
  });
  if (!seriesList.length) { alert(`未能在任一文件中找到采样通道「${ch}」对应的${METRIC_LABELS[activeMetric]}列。`); return; }

  // Reflect context in the CSV preview area.
  const firstSnap = snaps[0];
  const ds = ensureDatasetForSnapshot(firstSnap);
  setActiveDataset(ds.id);
  els.csvColSearch.value = ch;
  performColumnSearch();
  updateSummary();
  renderPreview();

  const yLabel = `${METRIC_LABELS[activeMetric]} (${METRIC_UNITS[activeMetric]})`;
  const type = els.chartType.value || 'line';
  clearChart();
  els.chartContainer.classList.remove('hidden');
  els.chartTitle.textContent = `波形对比：${ch} · ${yLabel}`;
  els.chartMeta.textContent = `叠加 ${seriesList.length} 个文件同一采样通道的${METRIC_LABELS[activeMetric]}波形（X = 时间/第一列，最多绘制 ${CHART_MAX_POINTS} 点/文件）`;
  updateChartMetricSelector();
  const ctx = els.csvChart.getContext('2d');
  state.chartInstance = new Chart(ctx, {
    type,
    data: { labels: refLabels || [], datasets: seriesList },
    options: buildChartOptions(
      {
        title: { display: true, text: `采样通道 ${ch} ${METRIC_LABELS[activeMetric]}波形对比`, font: { size: 12 } },
      },
      {
        scales: {
          x: { title: { display: true, text: '时间 / 第一列', font: { weight: 'bold' } }, ticks: { autoSkip: true, maxTicksLimit: 12, font: { size: 10 } } },
          y: { title: { display: true, text: yLabel, font: { weight: 'bold' } }, ticks: { font: { size: 10 } } },
        },
      }
    ),
  });
  bindChartInteractions(state.chartInstance);
  setTimeout(() => { if (els.chartContainer && !els.chartContainer.classList.contains('hidden')) els.chartContainer.scrollIntoView({ behavior: 'smooth', block: 'center' }); }, 120);
}

function getProjectSelectionKey(snapOrProjectId) {
  if (snapOrProjectId && typeof snapOrProjectId === 'object') return snapOrProjectId.projectId || '__default__';
  return snapOrProjectId || '__default__';
}
function getCrossProjectSelection(groupId) {
  if (!state.crossProjectChannelSelection[groupId]) state.crossProjectChannelSelection[groupId] = {};
  return state.crossProjectChannelSelection[groupId];
}
function clearCrossProjectChannelSelection(groupId) {
  if (groupId && state.crossProjectChannelSelection) delete state.crossProjectChannelSelection[groupId];
  if (groupId && state.crossProjectChartInstances && state.crossProjectChartInstances[groupId]) {
    try { state.crossProjectChartInstances[groupId].destroy(); } catch (e) { /* noop */ }
    delete state.crossProjectChartInstances[groupId];
  }
}
function toggleCrossProjectChannelSelection(groupId, snapId, projectId, channel, colIndex) {
  const group = getCompareGroup(groupId);
  const snap = findSnapshotById(snapId);
  if (!group || !snap) { alert('对比组或快照已被移除。'); return; }
  const key = getProjectSelectionKey(projectId || snap.projectId);
  const sel = getCrossProjectSelection(groupId);
  const ch = String(channel || '').trim();
  if (!ch || ch === '-') return;
  if (sel[key] && sel[key].snapId === snapId && sel[key].channel === ch && Number(sel[key].colIndex) === Number(colIndex)) {
    delete sel[key];
  } else {
    sel[key] = { snapId, channel: ch, colIndex: Number(colIndex) };
  }
  if (state.crossProjectChartInstances[groupId]) {
    try { state.crossProjectChartInstances[groupId].destroy(); } catch (e) { /* noop */ }
    delete state.crossProjectChartInstances[groupId];
  }
  renderCompareArea();
}
function getCrossProjectSeriesFromSelection(groupId) {
  const sel = state.crossProjectChannelSelection[groupId] || {};
  const entries = Object.keys(sel).map((key) => ({ key, item: sel[key], snap: findSnapshotById(sel[key].snapId) })).filter((x) => x.snap);
  const seriesList = [];
  const issues = [];
  let refLabels = [];
  let maxLen = 0;
  const metric = state.crossProjectMetricType[groupId] || METRIC_CURRENT;
  entries.forEach(({ item, snap }, i) => {
    const projectName = snap.projectName || '未指定项目';
    const record = {
      projectName,
      channel: item.channel,
      snapId: item.snapId,
      data: [],
      hasData: false,
      reason: '未找到 CSV 原始数据',
    };
    const rows = snap.csvRows;
    if (!rows || !rows.length) {
      record.reason = 'CSV 原始数据未找到';
    } else {
      let hdrIdx = snap.csvHeaderRowIndex;
      let dataStart = snap.csvDataStartRow;
      if (hdrIdx < 0) {
        const d = detectHeaderRow(rows);
        hdrIdx = d.headerRowIndex;
        dataStart = d.dataStartRow;
      }
      const targetC = findChannelColByMetric(rows, hdrIdx, item.channel, metric);
      if (targetC < 0) {
        record.reason = `${METRIC_LABELS[metric]}列缺失`;
      } else {
        const dataRows = rows.slice(dataStart >= 0 ? dataStart : 1);
        if (!dataRows.length) {
          record.reason = '数据为空';
        } else {
          const step = Math.max(1, Math.ceil(dataRows.length / CHART_MAX_POINTS));
          const xs = [];
          const ys = [];
          let hasNumericPoint = false;
          let hasNonZeroPoint = false;
          for (let r = 0; r < dataRows.length; r += step) {
            const row = dataRows[r] || [];
            const raw = (row[targetC] || '').toString().trim();
            let num = null;
            if (raw !== '' && NUMERIC_RE.test(raw)) {
              const n = Number(raw);
              if (Number.isFinite(n)) {
                num = n;
                hasNumericPoint = true;
                if (n !== 0) hasNonZeroPoint = true;
              }
            }
            xs.push((row[0] || '').toString().trim());
            ys.push(num);
          }
          if (!hasNumericPoint) {
            record.reason = '数据为空';
          } else {
            record.data = ys;
            record.hasData = true;
            if (!hasNonZeroPoint) record.reason = '数据点为 0';
            if (ys.length > maxLen) {
              maxLen = ys.length;
              refLabels = xs;
            }
          }
        }
      }
    }
    const color = record.hasData ? DS_COLORS[i % DS_COLORS.length] : '#B3B3B3';
    if (!record.hasData) issues.push(record);
    seriesList.push({
      label: `${projectName} · ${item.channel}${record.hasData ? '' : '（无原始数据）'}`,
      data: record.hasData ? record.data : [],
      borderColor: color,
      backgroundColor: record.hasData ? color + '33' : 'rgba(179, 179, 179, 0.18)',
      pointRadius: 0,
      borderWidth: 1.7,
      tension: 0.15,
      spanGaps: true,
      borderDash: record.hasData ? [] : [4, 4],
    });
  });
  return { labels: refLabels, datasets: seriesList, issues };
}
function renderCrossProjectChannelBar(group) {
  const snaps = group.snapshots || [];
  const projectMap = new Map();
  snaps.forEach((snap) => {
    const key = getProjectSelectionKey(snap);
    if (!projectMap.has(key)) projectMap.set(key, snap.projectName || (key === '__default__' ? '未指定项目' : key));
  });
  const sel = state.crossProjectChannelSelection[group.id] || {};
  const selectedCount = Array.from(projectMap.keys()).filter((key) => sel[key]).length;
  const projectCount = projectMap.size;
  const disabled = projectCount < 2 || selectedCount < projectCount;
  const chips = Array.from(projectMap.entries()).map(([key, name]) => {
    const item = sel[key];
    const text = item ? `${name}：${item.channel}（点击表格切换）` : `${name}：未选择`;
    return `<span class="cp-sel-chip ${item ? 'active' : ''}" title="${escapeHtml(text)}">${escapeHtml(text)}</span>`;
  }).join('');
  const tip = projectCount < 2 ? '跨项目通道对比需要至少 2 个项目' : '每个项目各选 1 个通道后可生成对比曲线';
  return `<div class="cp-channel-bar" data-group-id="${escapeHtml(group.id)}"><div class="cp-channel-title"><span class="material-symbols-outlined" style="font-size:15px;vertical-align:-3px">show_chart</span> 通道对比</div><div class="cp-sel-list">${chips}</div><button type="button" class="cmp-commoncol-btn cmp-commoncol-btn-apply" data-action="cross-generate" data-group-id="${escapeHtml(group.id)}" ${disabled ? 'disabled' : ''}>生成对比曲线</button><button type="button" class="cmp-commoncol-btn" data-action="cross-clear" data-group-id="${escapeHtml(group.id)}">清空选择</button><button type="button" class="cmp-commoncol-btn" data-action="cross-close" data-group-id="${escapeHtml(group.id)}">关闭图表</button><span class="cmp-commoncol-state">${escapeHtml(tip)}</span></div>`;
}
function renderCrossProjectChartShell(groupId) {
  const activeMetric = state.crossProjectMetricType[groupId] || METRIC_CURRENT;
  const metricBtns = [METRIC_CURRENT, METRIC_VOLTAGE, METRIC_POWER].map((m) => {
    const label = METRIC_LABELS[m];
    const cls = m === activeMetric ? 'chart-metric-btn active' : 'chart-metric-btn';
    return `<button type="button" class="${cls}" data-cross-metric="${m}" data-group-id="${escapeHtml(groupId)}">${escapeHtml(label)}</button>`;
  }).join('');
  const metricSelector = `<div class="chart-metric-selector" style="margin-left:12px"><span class="chart-metric-label">指标：</span>${metricBtns}</div>`;
  const unit = METRIC_UNITS[activeMetric] || 'mW';
  const metricLabel = METRIC_LABELS[activeMetric] || '功率 (P)';
  return `<div class="cp-chart-panel hidden" data-cross-chart-panel="${escapeHtml(groupId)}"><div class="cp-chart-head"><div><div class="cp-chart-title">跨项目通道波形对比${metricSelector}</div><div class="cp-chart-meta">X = 时间/第一列，最多绘制 ${CHART_MAX_POINTS} 点/文件；Y = ${metricLabel} (${unit})</div></div></div><div class="cp-chart-warning hidden" data-cross-chart-warning="${escapeHtml(groupId)}"></div><div class="cp-chart-canvas-wrap"><canvas data-cross-chart-canvas="${escapeHtml(groupId)}"></canvas></div></div>`;
}
function updateCrossProjectChartWarning(panel, groupId, issues) {
  if (!panel) return;
  const warningEl = panel.querySelector(`[data-cross-chart-warning="${CSS.escape(groupId)}"]`);
  if (!warningEl) return;
  if (!issues || !issues.length) {
    warningEl.classList.add('hidden');
    warningEl.innerHTML = '';
    return;
  }
  const items = issues.map((issue) => `<li>${escapeHtml(`${issue.projectName} · ${issue.channel}：${issue.reason || '未找到 CSV 原始数据'}`)}</li>`).join('');
  warningEl.innerHTML = `<div class="cp-chart-warning-title">⚠️ 以下项目通道未找到 CSV 原始数据或暂无法绘制曲线：</div><ul>${items}</ul>`;
  warningEl.classList.remove('hidden');
}
function generateCrossProjectChannelChart(groupId) {
  const group = getCompareGroup(groupId);
  if (!group) return;
  const projectKeys = new Set((group.snapshots || []).map((s) => getProjectSelectionKey(s)));
  const sel = state.crossProjectChannelSelection[groupId] || {};
  const selectedCount = Array.from(projectKeys).filter((key) => sel[key]).length;
  if (projectKeys.size < 2 || selectedCount < projectKeys.size) { alert('请先为每个项目各选择一个通道。'); return; }
  const data = getCrossProjectSeriesFromSelection(groupId);
  if (!data.datasets.length) { alert('请先为每个项目各选择一个通道。'); return; }
  const panel = els.compareGroupsContainer.querySelector(`[data-cross-chart-panel="${CSS.escape(groupId)}"]`);
  const canvas = els.compareGroupsContainer.querySelector(`[data-cross-chart-canvas="${CSS.escape(groupId)}"]`);
  if (!panel || !canvas) return;
  if (state.crossProjectChartInstances[groupId]) { try { state.crossProjectChartInstances[groupId].destroy(); } catch (e) { /* noop */ } }
  panel.classList.remove('hidden');
  updateCrossProjectChartWarning(panel, groupId, data.issues || []);
  const chartData = { labels: data.labels, datasets: data.datasets };
  const hasDrawableData = data.datasets.some((ds) => Array.isArray(ds.data) && ds.data.some((value) => Number.isFinite(value)));
  const activeMetric = state.crossProjectMetricType[groupId] || METRIC_CURRENT;
  const yAxisLabel = `${METRIC_LABELS[activeMetric]} (${METRIC_UNITS[activeMetric]})`;
  const chartExtraOptions = hasDrawableData ? { scales: { x: { title: { display: true, text: '时间 / 第一列', font: { weight: 'bold' } }, ticks: { autoSkip: true, maxTicksLimit: 12, font: { size: 10 } } }, y: { title: { display: true, text: yAxisLabel, font: { weight: 'bold' } }, ticks: { font: { size: 10 } } } } } : { scales: { x: { title: { display: true, text: '时间 / 第一列', font: { weight: 'bold' } }, ticks: { autoSkip: true, maxTicksLimit: 12, font: { size: 10 } } }, y: { title: { display: true, text: yAxisLabel, font: { weight: 'bold' } }, ticks: { font: { size: 10 } }, suggestedMin: 0, suggestedMax: 1 } } };
  state.crossProjectChartInstances[groupId] = new Chart(canvas.getContext('2d'), {
    type: 'line',
    data: chartData,
    options: buildChartOptions({ title: { display: true, text: `跨项目通道波形对比 · ${METRIC_LABELS[activeMetric]}`, font: { size: 12 } } }, chartExtraOptions),
  });
  bindChartInteractions(state.crossProjectChartInstances[groupId]);
  // Update metric buttons active state in the panel
  panel.querySelectorAll('.chart-metric-btn[data-cross-metric]').forEach((btn) => {
    btn.classList.toggle('active', btn.dataset.crossMetric === activeMetric);
  });
  panel.scrollIntoView({ behavior: 'smooth', block: 'center' });
}
function closeCrossProjectChannelChart(groupId) {
  if (state.crossProjectChartInstances[groupId]) { try { state.crossProjectChartInstances[groupId].destroy(); } catch (e) { /* noop */ } delete state.crossProjectChartInstances[groupId]; }
  const panel = els.compareGroupsContainer.querySelector(`[data-cross-chart-panel="${CSS.escape(groupId)}"]`);
  if (panel) panel.classList.add('hidden');
}

function renderSnapshotTable(snap, diffSet) {
  const rows = snap.rows;
  if (!rows.length) return '<div class="p-3 text-xs text-slate-500">无数据</div>';
  const colIndices = snap.colIndices;
  const numeric = snap.numeric || {};

  const valueCols = colIndices.filter((c) => c >= COMPARE_COMMON_END);
  const powerCol = valueCols.length ? detectPowerCol(rows, valueCols) : -1;
  let maxRow = -1;
  let maxVal = -Infinity;
  if (powerCol >= 0) {
    const dstart = getProcessedDataStart(rows, powerCol);
    for (let r = dstart; r < rows.length; r += 1) {
      const v = toNumberOrNull(rows[r] && rows[r][powerCol]);
      if (v !== null && v > maxVal) {
        maxVal = v;
        maxRow = r;
      }
    }
  }
  const snapChannelCol = detectChannelCol(rows);
  const channelAtMax = maxRow >= 0 && rows[maxRow]
    ? (rows[maxRow][snapChannelCol] === null || rows[maxRow][snapChannelCol] === undefined ? '' : String(rows[maxRow][snapChannelCol]))
    : '';

  const numericMap = new Map();
  colIndices.forEach((ci) => numericMap.set(ci, !!numeric[ci]));
  const mergeMeta = buildDisplayMergeMeta(snap.merges || [], colIndices, rows.length);
  const headerRowCount = Math.max(1, powerCol >= 0 ? getProcessedDataStart(rows, powerCol) : 1);

  const table = document.createElement('table');
  table.className = 'resizable-table compact-table numeric-uniform';
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  table.appendChild(thead);
  table.appendChild(tbody);

  appendProcessedSectionRows(thead, rows, 0, headerRowCount, colIndices, numericMap, mergeMeta, { header: true, tableType: 'snapshot', valueColSet: detectValueColumns(rows, colIndices), channelCol: detectChannelCol(rows) });
  appendProcessedSectionRows(tbody, rows, headerRowCount, rows.length, colIndices, numericMap, mergeMeta, {
    tableType: 'snapshot',
    snapId: snap.id,
    diffSet,
    powerCol,
    channelAtMax,
    maxRow,
    valueColSet: detectValueColumns(rows, colIndices),
    channelCol: detectChannelCol(rows),
  });

  return table.outerHTML;
}

// Resolve which common columns (index < COMPARE_COMMON_END) to show in compare
// tables. Default (state.compareCommonColIndices === null) shows only the
// sampling-channel (QEPM) column.
function resolveCompareCommonCols(cmpChannelCol, maxColsRef) {
  const sel = state.compareCommonColIndices;
  let cols;
  if (!sel) {
    cols = (cmpChannelCol >= 0 && cmpChannelCol < maxColsRef) ? [cmpChannelCol] : [];
  } else {
    cols = sel.filter((c) => c >= 0 && c < maxColsRef && c < COMPARE_COMMON_END);
  }
  if (!cols.length) {
    if (cmpChannelCol >= 0 && cmpChannelCol < maxColsRef) cols = [cmpChannelCol];
    else if (maxColsRef > 0) cols = [0];
  }
  return [...new Set(cols)].sort((a, b) => a - b);
}

function resolveCompareCommonColsForSnap(snap) {
  const availableCommonCols = (snap.colIndices || []).filter((c) => c >= 0 && c < COMPARE_COMMON_END);
  if (!availableCommonCols.length) return [];
  const availableSet = new Set(availableCommonCols);
  const rows = snap.rows || [];
  const channelCol = detectChannelCol(rows);
  const vavgCol = detectVavgCol(rows);

  // Normalise project key (empty -> "__default__") for per-project overrides.
  const projectId = snap.projectId || '__default__';
  const perProjectMap = state.compareCommonColIndicesByProject || {};
  const hasProjectKey = Object.prototype.hasOwnProperty.call(perProjectMap, projectId);
  const projectSetting = hasProjectKey ? perProjectMap[projectId] : undefined;

  let cols;
  if (Array.isArray(projectSetting) && projectSetting.length) {
    // 1) Per-project explicit common-column selection.
    cols = projectSetting.filter((c) => availableSet.has(c));
  } else if (Array.isArray(state.compareCommonColIndices) && state.compareCommonColIndices.length) {
    // 2) Global common-column selection used as fallback when no per-project override.
    cols = state.compareCommonColIndices.filter((c) => availableSet.has(c));
  } else {
    // 3) Default columns: QEPM channel + Vavg (if present).
    cols = [];
    if (channelCol >= 0 && availableSet.has(channelCol)) cols.push(channelCol);
    if (vavgCol >= 0 && availableSet.has(vavgCol) && vavgCol !== channelCol) cols.push(vavgCol);
  }

  if (!cols.length) {
    // Still nothing after filtering (e.g. channel/Vavg missing in this snapshot) —
    // fall back to at least one common header so the table remains usable.
    if (channelCol >= 0 && availableSet.has(channelCol)) cols.push(channelCol);
    else cols.push(availableCommonCols[0]);
  }

  return [...new Set(cols)].sort((a, b) => a - b);
}

// Lightweight control bar (rendered atop the compare groups) that lets users
// customise which common columns appear in every same-project / cross-project
// compare view.
function renderCompareCommonColBar() {
  const sel = state.compareCommonColIndices;
  const val = sel ? sel.map((c) => indexToColLetter(c)).join(',') : '';
  const stateText = sel
    ? `当前：全局自定义公共列 ${escapeHtml(val)}（未针对某个项目单独设置时生效）`
    : '当前：默认显示「采样通道编号（QEPM）」+「Vavg」列（未针对某个项目单独设置时生效，同项目 / 跨项目对比均生效）';
  const placeholder = '留空 = 使用默认「采样通道编号（QEPM）」+「Vavg」列，可填如 A-D,M 作为全局默认；仅对未按项目单独设置的对比生效';
  return `
    <div class="cmp-commoncol-bar">
      <span class="cmp-commoncol-label"><span class="material-symbols-outlined" style="font-size:15px;vertical-align:-3px">view_column</span> 对比公共列（全局默认）：</span>
      <input id="cmpCommonColRange" type="text" class="cmp-commoncol-input" placeholder="${escapeHtml(placeholder)}" value="${escapeHtml(val)}" />
      <button type="button" class="cmp-commoncol-btn cmp-commoncol-btn-apply" data-action="cmp-common-apply">应用</button>
      <button type="button" class="cmp-commoncol-btn" data-action="cmp-common-default">仅默认列</button>
      <button type="button" class="cmp-commoncol-btn" data-action="cmp-common-all">全部公共列</button>
      <span class="cmp-commoncol-state">${stateText}</span>
    </div>`;
}

// Apply the expression typed in #cmpCommonColRange to the global common-column
// selection. Empty / invalid input falls back to default (QEPM + Vavg).
function applyCompareCommonColExpr() {
  const input = document.getElementById('cmpCommonColRange');
  if (!input) return;
  const expr = (input.value || '').trim();
  if (!expr) { state.compareCommonColIndices = null; renderCompareArea(); return; }
  const parsed = parseColRange(expr);
  const filtered = (parsed || []).filter((c) => c >= 0 && c < COMPARE_COMMON_END);
  state.compareCommonColIndices = filtered.length ? filtered : null;
  renderCompareArea();
}

function getProjectKey(projectId) {
  return projectId || '__default__';
}

function getProjectAvailableCommonCols(projectId) {
  const key = getProjectKey(projectId);
  const colsSet = new Set();
  for (const g of state.compareGroups) {
    for (const snap of g.snapshots || []) {
      if (!snap) continue;
      const snapKey = getProjectKey(snap.projectId);
      if (snapKey !== key) continue;
      (snap.colIndices || []).forEach((ci) => {
        if (ci >= 0 && ci < COMPARE_COMMON_END) colsSet.add(ci);
      });
    }
  }
  const cols = Array.from(colsSet).sort((a, b) => a - b);
  // If we couldn't infer any project-specific common cols, fall back to the
  // generic "all common" range [0, COMPARE_COMMON_END).
  return cols.length ? cols : Array.from({ length: COMPARE_COMMON_END }, (_, i) => i);
}

function getProjectCurrentCommonCols(projectId) {
  const key = getProjectKey(projectId);
  const perProjectMap = state.compareCommonColIndicesByProject || {};
  const hasProjectKey = Object.prototype.hasOwnProperty.call(perProjectMap, key);
  const projectSetting = hasProjectKey ? perProjectMap[key] : undefined;
  if (Array.isArray(projectSetting) && projectSetting.length) {
    return projectSetting.slice();
  }
  if (Array.isArray(state.compareCommonColIndices) && state.compareCommonColIndices.length) {
    return state.compareCommonColIndices.slice();
  }
  return null;
}

function renderProjectCommonColBars(snaps) {
  const projectMap = new Map(); // key -> { name }
  (snaps || []).forEach((snap) => {
    if (!snap) return;
    const key = getProjectKey(snap.projectId);
    if (!projectMap.has(key)) {
      const name = snap.projectName || (key === '__default__' ? '未指定项目' : key);
      projectMap.set(key, { name });
    }
  });
  if (!projectMap.size) return '';
  const placeholder = '留空 = 使用默认「采样通道编号（QEPM）」+「Vavg」列，或继承全局设置；仅影响当前项目的对比视图';
  const parts = [];
  projectMap.forEach(({ name }, key) => {
    const current = getProjectCurrentCommonCols(key);
    const val = Array.isArray(current) ? current.map((c) => indexToColLetter(c)).join(',') : '';
    parts.push(`
      <div class="cmp-commoncol-bar" data-project-common-bar="1" data-project-id="${escapeHtml(key)}">
        <span class="cmp-commoncol-label"><span class="material-symbols-outlined" style="font-size:15px;vertical-align:-3px">view_column</span> 项目：${escapeHtml(name)}</span>
        <input type="text" class="cmp-commoncol-input" data-role="cmp-project-input" data-project-id="${escapeHtml(key)}" placeholder="${escapeHtml(placeholder)}" value="${escapeHtml(val)}" />
        <button type="button" class="cmp-commoncol-btn cmp-commoncol-btn-apply" data-action="cmp-project-apply" data-project-id="${escapeHtml(key)}">应用</button>
        <button type="button" class="cmp-commoncol-btn" data-action="cmp-project-default" data-project-id="${escapeHtml(key)}">仅默认列</button>
        <button type="button" class="cmp-commoncol-btn" data-action="cmp-project-all" data-project-id="${escapeHtml(key)}">全部公共列</button>
        <button type="button" class="cmp-commoncol-btn" data-action="cmp-project-reset-global" data-project-id="${escapeHtml(key)}">重置为全局默认</button>
      </div>`);
  });
  return parts.join('');
}

function applyProjectCommonColExpr(projectId) {
  const key = getProjectKey(projectId);
  if (!key || !els.compareGroupsContainer) return;
  const inputs = els.compareGroupsContainer.querySelectorAll('input[data-role="cmp-project-input"]');
  let input = null;
  inputs.forEach((el) => {
    if (el.dataset.projectId === key) input = el;
  });
  const expr = input ? (input.value || '').trim() : '';
  const parsed = expr ? parseColRange(expr) : null;
  const filtered = (parsed || []).filter((c) => c >= 0 && c < COMPARE_COMMON_END);
  // Empty expression or invalid input falls back to per-project default (null).
  state.compareCommonColIndicesByProject[key] = filtered.length ? filtered : null;
  renderCompareArea();
}

function renderSameProjectCompare(snaps) {
  const ref = snaps[0];
  const cmpChannelCol = detectChannelCol(ref.rows);
  const maxColsRef = getColCount(ref.rows.map((r) => r.map((v) => (v === null || v === undefined ? '' : v))));
  const commonCols = resolveCompareCommonColsForSnap(ref);
  let valueCols = COMPARE_VALUE_COLS.filter((c) => c < maxColsRef);
  if (!valueCols.length) valueCols = ref.colIndices.filter((c) => c >= COMPARE_COMMON_END);
  const powerCol = detectPowerCol(ref.rows, valueCols);
  const dataStart = getProcessedDataStart(ref.rows, powerCol >= 0 ? powerCol : valueCols[valueCols.length - 1]);
  const maxRows = Math.max(...snaps.map((s) => s.rows.length));

  let gSpread = -Infinity;
  let gSpreadRow = -1;
  for (let r = dataStart; r < maxRows; r += 1) {
    let mn = Infinity;
    let mx = -Infinity;
    let cnt = 0;
    for (const s of snaps) {
      const v = toNumberOrNull(s.rows[r] && s.rows[r][powerCol]);
      if (v === null) continue;
      cnt += 1;
      if (v < mn) mn = v;
      if (v > mx) mx = v;
    }
    if (cnt >= 2) {
      const spread = mx - mn;
      if (spread > gSpread) {
        gSpread = spread;
        gSpreadRow = r;
      }
    }
  }

  const commonLabels = commonCols.map((c) => ({ letter: indexToColLetter(c), idx: c, name: processedHeaderLabel(ref.rows, dataStart, c) }));
  const valueLabels = valueCols.map((c) => ({ letter: indexToColLetter(c), idx: c, name: processedHeaderLabel(ref.rows, dataStart, c), isPower: c === powerCol }));
  const mergeMeta = buildDisplayMergeMeta(ref.merges || [], commonCols, ref.rows.length);

  let html = '<div class="cmp-legend">';
  snaps.forEach((s, i) => {
    const color = DS_COLORS[i % DS_COLORS.length];
    html += `<div class="cmp-legend-item"><span class="cmp-legend-swatch" style="background:${color}"></span><span class="cmp-legend-name" title="${escapeHtml(s.csvName)}">#${i + 1} ${escapeHtml(s.csvName)}</span><button class="cmp-legend-x" data-action="remove-snap" data-id="${s.id}" title="移除此份"><span class="material-symbols-outlined" style="font-size:14px">close</span></button></div>`;
  });
  html += '</div>';

  const table = document.createElement('table');
  table.className = 'cmp-aligned compact-table numeric-uniform resizable-table';
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  table.appendChild(thead);
  table.appendChild(tbody);

  const headerTr1 = document.createElement('tr');
  commonLabels.forEach((cl) => {
    const th = document.createElement('th');
    th.rowSpan = 2;
    th.className = 'cmp-common-th';
    th.dataset.resizeCol = `common_${cl.idx}`;
    th.innerHTML = `<div class="th-letter" style="font-size:9px">${cl.letter}</div><span class="th-truncate" style="max-width:80px" data-full="${escapeHtml(cl.name)}" data-letter="${cl.letter}" data-idx="${cl.idx + 1}">${escapeHtml(cl.name || '—')}</span>`;
    headerTr1.appendChild(th);
  });
  snaps.forEach((s, i) => {
    const color = DS_COLORS[i % DS_COLORS.length];
    const th = document.createElement('th');
    th.colSpan = valueCols.length;
    th.className = 'cmp-ds-th';
    th.style.borderBottom = `3px solid ${color}`;
    th.innerHTML = `<span class="cmp-ds-badge" style="background:${color}">#${i + 1}</span> <span class="cmp-ds-name" title="${escapeHtml(s.csvName)}">${escapeHtml(s.csvName)}</span>`;
    headerTr1.appendChild(th);
  });
  thead.appendChild(headerTr1);

  const headerTr2 = document.createElement('tr');
  snaps.forEach(() => {
    valueLabels.forEach((vl) => {
      const th = document.createElement('th');
      th.dataset.resizeCol = `value_${vl.idx}`;
      if (vl.isPower) th.classList.add('cmp-power-th');
      th.innerHTML = `<div class="th-letter" style="font-size:9px">${vl.letter}</div><span class="th-truncate" style="max-width:66px" data-full="${escapeHtml(vl.name)}" data-letter="${vl.letter}" data-idx="${vl.idx + 1}">${escapeHtml(vl.name || '—')}</span>`;
      headerTr2.appendChild(th);
    });
  });
  thead.appendChild(headerTr2);

  for (let r = dataStart; r < maxRows; r += 1) {
    const pVals = snaps.map((s) => toNumberOrNull(s.rows[r] && s.rows[r][powerCol]));
    const present = pVals.filter((v) => v !== null);
    const distinct = new Set(present.map((v) => formatNumeric2(v)));
    const isDiff = present.length >= 2 && distinct.size > 1;
    let pMax = -Infinity;
    let pMin = Infinity;
    present.forEach((v) => {
      if (v > pMax) pMax = v;
      if (v < pMin) pMin = v;
    });
    const channelVal = ref.rows[r] ? (ref.rows[r][cmpChannelCol] === null || ref.rows[r][cmpChannelCol] === undefined ? '' : String(ref.rows[r][cmpChannelCol])) : '';
    const canClick = channelVal && channelVal !== '-';

    const tr = document.createElement('tr');
    commonCols.forEach((c) => {
      const td = document.createElement('td');
      td.className = 'cmp-common-td';
      td.dataset.resizeCol = `common_${c}`;
      const val = ref.rows[r] ? (ref.rows[r][c] === null || ref.rows[r][c] === undefined ? '' : ref.rows[r][c]) : '';
      const displayText = formatNumeric2(val);
      if (c === cmpChannelCol && canClick) {
        td.classList.add('clickable-power');
        if (isDiff) td.classList.add('channel-highlight');
        td.dataset.action = 'chart-overlay';
        td.dataset.channel = channelVal;
        td.title = `点击：对比各文件通道「${channelVal}」的功率波形差异`;
      } else {
        td.title = displayText;
      }
      td.textContent = displayText;
      tr.appendChild(td);
    });

    snaps.forEach((s, si) => {
      valueCols.forEach((c) => {
        const td = document.createElement('td');
        td.dataset.resizeCol = `value_${c}`;
        td.classList.add('numeric-col', 'value-col-full');
        const val = s.rows[r] ? (s.rows[r][c] === null || s.rows[r][c] === undefined ? '' : s.rows[r][c]) : '';
        if (c === powerCol) {
          if (canClick) {
            td.classList.add('clickable-power');
            td.dataset.action = 'chart-overlay';
            td.dataset.channel = channelVal;
            td.title = '点击：对比各文件该通道的功率波形差异';
          }
          if (isDiff) {
            td.classList.add('compare-diff-power');
            const vNum = pVals[si];
            if (vNum !== null && vNum === pMax) td.classList.add('compare-diff-hi');
            else if (vNum !== null && vNum === pMin) td.classList.add('compare-diff-lo');
            if (r === gSpreadRow) td.classList.add('compare-max-diff');
          }
        }
        const displayText = formatNumeric2(val);
        td.textContent = displayText;
        if (!td.title) td.title = displayText;
        tr.appendChild(td);
      });
    });
    tbody.appendChild(tr);
  }

  html += `<div class="cmp-aligned-scroll tight-scroll">${table.outerHTML}</div>`;
  const powerLetter = powerCol >= 0 ? indexToColLetter(powerCol) : 'P';
  html += `<p class="cmp-hint"><span class="material-symbols-outlined" style="font-size:14px;vertical-align:-2px">tips_and_updates</span> 同项目对比：左侧公共列<b>默认展示「采样通道编号（QEPM）」+「Vavg」两列</b>，可在上方「对比公共列」全局控件或按项目控件条中自定义要展示的公共列（如 <code>A-D,M</code>，逐列独立展示、支持拖动表头边缘调节列宽）；右侧按文件 #1/#2/#3… 并排展示 N/O/P。系统逐行对比各文件 <b>Power（${powerLetter} 列 Pavg(mW)）</b> 的差异：数值不一致的行会<span class="cmp-hint-hl">高亮</span>（最大值<span class="cmp-hint-hi">偏红</span>、最小值<span class="cmp-hint-lo">偏蓝</span>），差异最大的一行以<span class="cmp-hint-max">加粗描边</span>标注。<b>点击 M 采样通道或任意 P 单元格</b>，即可在下方按同一采样通道叠加各文件的功率波形，直观对比波形差异。</p>`;
  return html;
}

function renderCrossProjectSnapshotTable(snap, diffSet, groupId) {
  const rows = snap.rows || [];
  if (!rows.length) return '<div class="p-3 text-xs text-slate-500">无数据</div>';

  const colIndices = snap.colIndices || [];
  const numeric = snap.numeric || {};
  const commonCols = resolveCompareCommonColsForSnap(snap);
  let valueCols = COMPARE_VALUE_COLS.filter((c) => colIndices.includes(c));
  if (!valueCols.length) valueCols = colIndices.filter((c) => c >= COMPARE_COMMON_END);
  const displayCols = [...commonCols, ...valueCols];
  if (!displayCols.length) return '<div class="p-3 text-xs text-slate-500">无可展示列</div>';

  const powerCol = valueCols.length ? detectPowerCol(rows, valueCols) : -1;
  const channelCol = detectChannelCol(rows);
  const headerRefCol = powerCol >= 0 ? powerCol : (valueCols[0] ?? displayCols[0]);
  const dataStart = getProcessedDataStart(rows, headerRefCol);

  const numericMap = new Map();
  displayCols.forEach((ci) => numericMap.set(ci, !!numeric[ci]));
  const valueColSet = new Set(valueCols);

  const table = document.createElement('table');
  table.className = 'resizable-table compact-table numeric-uniform';
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  table.appendChild(thead);
  table.appendChild(tbody);

  const headerTr = document.createElement('tr');
  displayCols.forEach((ci) => {
    const th = document.createElement('th');
    const isCommon = ci < COMPARE_COMMON_END;
    th.dataset.resizeCol = isCommon ? `xcommon_${snap.id}_${ci}` : `xvalue_${snap.id}_${ci}`;
    if (numericMap.get(ci)) th.classList.add('numeric-col');
    if (valueColSet.has(ci)) th.classList.add('value-col-full');
    const label = processedHeaderLabel(rows, dataStart, ci);
    th.innerHTML = `<div class="th-letter" style="font-size:9px">${indexToColLetter(ci)}</div><span class="th-truncate" style="max-width:${isCommon ? '92px' : '72px'}" data-full="${escapeHtml(label)}" data-letter="${indexToColLetter(ci)}" data-idx="${ci + 1}">${escapeHtml(label || '—')}</span>`;
    headerTr.appendChild(th);
  });
  thead.appendChild(headerTr);

  for (let r = dataStart; r < rows.length; r += 1) {
    const row = rows[r] || [];
    const tr = document.createElement('tr');
    displayCols.forEach((ci) => {
      const td = document.createElement('td');
      const isCommon = ci < COMPARE_COMMON_END;
      td.dataset.resizeCol = isCommon ? `xcommon_${snap.id}_${ci}` : `xvalue_${snap.id}_${ci}`;
      if (numericMap.get(ci)) td.classList.add('numeric-col');
      if (valueColSet.has(ci)) td.classList.add('value-col-full');
      const val = ci < row.length ? normalizeCellValue(row[ci]) : '';
      const displayText = formatNumeric2(val);
      td.textContent = displayText;
      td.title = displayText;
      if (diffSet && diffSet.has(`${r}_${ci}`)) td.classList.add('compare-diff');
      if (ci === powerCol) td.classList.add('compare-diff-power');

      const channelText = String(channelCol >= 0 && row[channelCol] !== null && row[channelCol] !== undefined ? row[channelCol] : '').trim();
      const projectKey = getProjectSelectionKey(snap);
      const selected = state.crossProjectChannelSelection[groupId] && state.crossProjectChannelSelection[groupId][projectKey];
      const isSelectedRow = selected && selected.snapId === snap.id && selected.channel === channelText;
      if (isSelectedRow) tr.classList.add('cp-selected-row');
      if ((ci === channelCol || ci === powerCol) && channelText && channelText !== '-') {
        td.classList.add('clickable-power');
        if (isSelectedRow) td.classList.add('cp-selected');
        td.dataset.action = 'cross-select';
        td.dataset.channel = channelText;
        td.dataset.groupId = groupId;
        td.dataset.snapId = snap.id;
        td.dataset.projectId = projectKey;
        td.dataset.colIndex = channelCol;
        td.title = isSelectedRow ? '已选择；再次点击取消本项目通道选择' : '点击：选择本项目该通道用于跨项目功率曲线对比';
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }

  return table.outerHTML;
}

function renderCompareArea() {
  const container = els.compareGroupsContainer;
  if (!container) return;
  const processedCount = state.datasets.filter((d) => d.processedRows && d.processedRows.length).length;
  const anySnap = state.compareGroups.some((g) => g.snapshots.length > 0);
  if (!processedCount && !anySnap) {
    els.compareArea.classList.add('hidden');
    container.innerHTML = '';
    if (els.compareGroupsChip) els.compareGroupsChip.textContent = '';
    return;
  }
  els.compareArea.classList.remove('hidden');
  if (els.compareGroupsChip) {
    const totalSnaps = state.compareGroups.reduce((sum, g) => sum + g.snapshots.length, 0);
    els.compareGroupsChip.textContent = `${state.compareGroups.length} 个对比组 · 共 ${totalSnaps} 份`;
  }
  if (!state.compareGroups.length) {
    container.innerHTML = '<div class="cmp-empty-groups">还没有对比组，点击右上角「➕ 新建对比组」创建。</div>';
    return;
  }
  const bar = anySnap ? renderCompareCommonColBar() : '';
  container.innerHTML = bar + state.compareGroups.map((g, gi) => renderCompareGroupBlock(g, gi)).join('');
  initResizableTables(container);
  applyStickyColumns(container);
}

function renderCompareGroupBlock(group, gIndex) {
  const snaps = group.snapshots;
  const count = snaps.length;
  const sameProject = count >= 1 && snaps.every((s) => s.projectId && s.projectId === snaps[0].projectId);
  let modeText;
  let modeCls;
  if (count === 0) { modeText = '空组'; modeCls = 'cmp-mode-wait'; }
  else if (count === 1) { modeText = '单组数据预览'; modeCls = 'cmp-mode-same'; }
  else if (sameProject) { modeText = '同项目对比 · 对齐 N/O/P'; modeCls = 'cmp-mode-same'; }
  else { modeText = '跨项目对比 · 并排'; modeCls = 'cmp-mode-diff'; }
  const gid = escapeHtml(group.id);
  return `
    <div class="cmp-group" data-group-id="${gid}">
      <div class="cmp-group-head">
        <div class="cmp-group-head-left">
          <span class="cmp-group-badge">#${gIndex + 1}</span>
          <span class="cmp-group-title" data-action="rename-group" data-group-id="${gid}" title="点击重命名">${escapeHtml(group.name)}</span>
          <span class="cmp-group-count">${count} / ${COMPARE_MAX}</span>
          <span class="cmp-mode-chip ${modeCls}">${modeText}</span>
        </div>
        <div class="cmp-group-head-right">
          <label class="cmp-group-hl"><input type="checkbox" data-action="group-highlight" data-group-id="${gid}" ${group.highlightDiff ? 'checked' : ''}/> 高亮差异</label>
          <button type="button" class="template-btn template-btn-sm" data-action="group-clear" data-group-id="${gid}">清空本组</button>
          <button type="button" class="cmp-group-remove-btn" data-action="group-remove" data-group-id="${gid}" title="删除整个对比组">
            <span class="material-symbols-outlined" style="font-size:14px;vertical-align:-2px">delete</span> 删除本组
          </button>
        </div>
      </div>
      <div class="cmp-group-body">${renderCompareGroupBody(group)}</div>
    </div>`;
}

function renderCompareGroupBody(group) {
  const snaps = group.snapshots;
  const count = snaps.length;
  const gid = escapeHtml(group.id);
  if (!count) {
    return `<div class="cmp-group-empty">本组暂无数据。在上方「数据组选择」面板里点击「${escapeHtml(group.name)}」标签，把数据集加入本组。</div>`;
  }
  const sameProject = snaps.every((s) => s.projectId && s.projectId === snaps[0].projectId);
  if (sameProject) {
    return `<div class="cmp-aligned-wrap" data-group-id="${gid}">${renderSameProjectCompare(snaps)}</div>`;
  }
  const diffMap = computeCompareDiffMap(snaps, group.highlightDiff);
  const projectBars = renderProjectCommonColBars(snaps);
  const cards = snaps.map((snap, idx) => `
    <div class="compare-card" data-id="${snap.id}">
      <div class="compare-card-head">
        <div class="min-w-0">
          <div class="compare-card-title" title="${escapeHtml(snap.projectName)}">#${idx + 1} ${escapeHtml(snap.projectName)}</div>
          <div class="compare-card-meta" title="${escapeHtml(snap.csvName)}">CSV：${escapeHtml(snap.csvName)} · ${formatTime(snap.timestamp)}</div>
        </div>
        <button class="template-btn template-btn-danger template-btn-sm" data-action="remove-snap" data-id="${snap.id}" title="移除此份对比">
          <span class="material-symbols-outlined text-sm">close</span>
        </button>
      </div>
      <div class="table-wrap tight-scroll">${renderCrossProjectSnapshotTable(snap, diffMap.get(snap.id), group.id)}</div>
    </div>`).join('');
  return `${projectBars}${renderCrossProjectChannelBar(group)}${renderCrossProjectChartShell(group.id)}<div class="cmp-group-grid cols-${Math.min(count, COMPARE_MAX)}" data-group-id="${gid}">${cards}</div>`;
}

// ---------- Header tooltip ----------
function showHeaderTooltip(anchorEl) {
  const full = anchorEl.dataset.full || anchorEl.textContent || '';
  const letter = anchorEl.dataset.letter || '';
  const idx = anchorEl.dataset.idx || '';
  els.headerTooltip.innerHTML = `
    <div class="htt-title">${escapeHtml(letter)}${idx ? ` · 第 ${escapeHtml(idx)} 列` : ''}</div>
    <div>${escapeHtml(full || '—')}</div>
    <div class="htt-close" data-role="htt-close">关闭 ✕</div>
  `;
  els.headerTooltip.classList.remove('hidden');
  const rect = anchorEl.getBoundingClientRect();
  const ttRect = els.headerTooltip.getBoundingClientRect();
  let top = rect.bottom + 6;
  let left = rect.left;
  const vw = window.innerWidth;
  const vh = window.innerHeight;
  if (left + ttRect.width + 8 > vw) left = Math.max(8, vw - ttRect.width - 8);
  if (top + ttRect.height + 8 > vh) top = Math.max(8, rect.top - ttRect.height - 6);
  els.headerTooltip.style.top = `${top}px`;
  els.headerTooltip.style.left = `${left}px`;
}
function hideHeaderTooltip() { els.headerTooltip.classList.add('hidden'); }

// ---------- Admin ----------
function checkAdminSession() { return Boolean(window.powerAuth && window.powerAuth.isAdmin()); }
function setAdminMode(isAdmin) {
  state.isAdmin = Boolean(isAdmin);
  state.serverIsAdmin = state.isAdmin;
  document.body.classList.toggle('is-admin', state.isAdmin);
  if (els.adminBadge) els.adminBadge.classList.toggle('hidden', !state.isAdmin);
  if (!state.isAdmin && state.templateView === 'upload') switchTemplateView('list');
}

// ---------- Events ----------
let _replaceTargetId = null;
function openReplaceFilePicker(id) { _replaceTargetId = id; els.templateInput.click(); }

function bindEvents() {
  els.templateInput.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (_replaceTargetId) {
      const id = _replaceTargetId;
      _replaceTargetId = null;
      handleTemplateFile(file, { replaceId: id });
    } else {
      handleTemplateFile(file);
    }
    event.target.value = '';
  });
  els.fileInput.addEventListener('change', (event) => { handleCsvFiles(event.target.files); event.target.value = ''; });
  els.fileInputNew.addEventListener('change', (event) => { handleCsvFiles(event.target.files); event.target.value = ''; });
  els.reuploadCsvBtn.addEventListener('click', () => els.fileInputNew.click());

  els.exportBtn.addEventListener('click', runExport);
  els.previewBtn.addEventListener('click', previewProcessed);
  els.downloadBtn.addEventListener('click', downloadProcessed);
  if (els.downloadAllBtn) els.downloadAllBtn.addEventListener('click', downloadAllProcessed);

  document.querySelectorAll('.upload-zone').forEach((zone) => {
    ['dragenter', 'dragover'].forEach((eventName) => {
      zone.addEventListener(eventName, (event) => { event.preventDefault(); zone.classList.add('drag-active'); });
    });
    ['dragleave', 'drop'].forEach((eventName) => {
      zone.addEventListener(eventName, (event) => { event.preventDefault(); zone.classList.remove('drag-active'); });
    });
  });
  // Drop CSV files onto CSV zones
  [els.csvDropzone, els.csvFileInfo].forEach((zone) => {
    if (!zone) return;
    zone.addEventListener('drop', (event) => {
      event.preventDefault();
      if (event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files.length) {
        handleCsvFiles(event.dataTransfer.files);
      }
    });
  });

  // Per-file project selector (delegated change handler on the uploaded-file list)
  if (els.csvFileInfo) {
    els.csvFileInfo.addEventListener('change', (event) => {
      const sel = event.target.closest('select[data-action="ds-project"]');
      if (!sel) return;
      handleDatasetProjectChange(sel.dataset.dsId, sel.value);
    });
    // Prevent a dropdown click from bubbling to any parent click handlers.
    els.csvFileInfo.addEventListener('click', (event) => {
      if (event.target.closest('select[data-action="ds-project"]')) event.stopPropagation();
    });
  }

  els.templateList.addEventListener('click', (event) => {
    const btn = event.target.closest('button[data-action]');
    if (btn) {
      event.stopPropagation();
      const id = btn.dataset.id;
      const action = btn.dataset.action;
      if (action === 'use') useTemplate(id);
      else if (action === 'detail') { state.templateDetailId = id; switchTemplateView('detail'); }
      return;
    }
    const row = event.target.closest('.template-item[data-role="row"]');
    if (row) { state.templateDetailId = row.dataset.id; switchTemplateView('detail'); }
  });

  els.templateDetailBody.addEventListener('click', (event) => {
    const btn = event.target.closest('button[data-action]');
    if (!btn) return;
    const id = btn.dataset.id;
    const action = btn.dataset.action;
    if (action === 'use') useTemplate(id);
    else if (action === 'download') downloadTemplate(id);
    else if (action === 'rename') { if (!state.isAdmin) { alert('只有管理员可以重命名模板。'); return; } renameTemplate(id); }
    else if (action === 'delete') { if (!state.isAdmin) { alert('只有管理员可以删除模板。'); return; } deleteTemplate(id); }
    else if (action === 'replace') { if (!state.isAdmin) { alert('只有管理员可以上传模板。'); return; } openReplaceFilePicker(id); }
  });

  els.templateDetailBack.addEventListener('click', () => switchTemplateView('list'));
  els.templateUploadBack.addEventListener('click', () => switchTemplateView('list'));
  els.templateNewBtn.addEventListener('click', () => { if (!state.isAdmin) { alert('只有管理员可以上传模板。'); return; } switchTemplateView('upload'); });

  els.templatePrevBtn.addEventListener('click', () => { if (state.templatePage > 1) { state.templatePage -= 1; renderTemplateListPage(); } });
  els.templateNextBtn.addEventListener('click', () => {
    const totalPages = Math.max(1, Math.ceil(state.templateItems.length / TEMPLATE_PAGE_SIZE));
    if (state.templatePage < totalPages) { state.templatePage += 1; renderTemplateListPage(); }
  });

  // Admin access follows the authenticated account email allowlist.

  if (els.templateRenameConfirmBtn) els.templateRenameConfirmBtn.addEventListener('click', submitTemplateRenameModal);
  if (els.templateRenameCancelBtn) els.templateRenameCancelBtn.addEventListener('click', closeTemplateRenameModal);
  if (els.templateRenameInput) els.templateRenameInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') submitTemplateRenameModal();
    if (event.key === 'Escape') closeTemplateRenameModal();
  });
  if (els.templateRenameModal) els.templateRenameModal.addEventListener('click', (event) => {
    if (event.target === els.templateRenameModal) closeTemplateRenameModal();
  });

  if (els.clearHistoryBtn) els.clearHistoryBtn.addEventListener('click', () => { clearHistory(); });
  if (els.clearHistoryCacheBtn) els.clearHistoryCacheBtn.addEventListener('click', () => { clearHistory(); });
  if (els.historyList) {
    els.historyList.addEventListener('click', (event) => {
      const target = event.target.closest('button[data-action]');
      if (!target) return;
      const { action, id } = target.dataset;
      if (action === 'delete-history') deleteHistoryEntry(id);
      else if (action === 'restore-history') restoreHistoryWorkspace(id).catch((error) => {
        console.error(error);
        alert(error.message || '历史回溯失败，请稍后重试。');
      });
      else if (action === 'download-history') downloadHistoryEntryResults(id).catch((error) => {
        console.error(error);
        alert(error.message || '历史结果下载失败，请稍后重试。');
      });
    });
  }

  if (els.projectSelect) {
    els.projectSelect.addEventListener('change', (event) => { const id = event.target.value; if (id) useTemplate(id); });
  }

  // Download current template button (next to project select)
  if (els.downloadTemplateBtn) {
    els.downloadTemplateBtn.addEventListener('click', () => {
      if (!state.activeProject) { alert('请先选择一个项目。'); return; }
      downloadTemplate(state.activeProject.id);
    });
  }

  // Dataset tabs (multi-group)
  if (els.datasetTabs) {
    els.datasetTabs.addEventListener('click', (event) => {
      const dlBtn = event.target.closest('[data-action="dl-dataset"]');
      if (dlBtn) { event.stopPropagation(); downloadDataset(dlBtn.dataset.id); return; }
      const tab = event.target.closest('[data-action="pick-dataset"]');
      if (tab) {
        setActiveDataset(tab.dataset.id);
        state.csvSearchKeyword = '';
        state.csvSearchMode = 'manual';
        state.csvMatchedCols = [];
        state.csvChartSelectedCols = [];
        if (els.csvColSearch) els.csvColSearch.value = '';
        clearChart();
        updateMatchInfoAndPicker();
        updateSummary();
        renderPreview();
      }
    });
  }

  // Per-project split preview tab switching
  if (els.processedSplitContainer) {
    els.processedSplitContainer.addEventListener('click', (event) => {
      const tab = event.target.closest('[data-action="pick-split-ds"]');
      if (!tab) return;
      const projectId = tab.dataset.projectId;
      const dsId = tab.dataset.dsId;
      if (!projectId || !dsId) return;
      state.projectActiveDsIds[projectId] = dsId;
      renderProcessedPreview();
    });
  }

  // Dataset selection panel: self-select preview / compare-group membership
  if (els.datasetSelectPanel) {
    els.datasetSelectPanel.addEventListener('change', (event) => {
      const cb = event.target.closest('input[type="checkbox"][data-action]');
      if (!cb) return;
      const dsId = cb.dataset.dsId;
      if (cb.dataset.action === 'toggle-preview') {
        togglePreviewSelect(dsId, cb.checked);
      }
    });
    els.datasetSelectPanel.addEventListener('click', (event) => {
      const btn = event.target.closest('button[data-action]');
      if (!btn) return;
      const action = btn.dataset.action;
      const processedDs = state.datasets.filter((d) => d.processedRows && d.processedRows.length);
      if (action === 'preview-all') {
        state.previewHiddenIds = new Set();
        renderProcessedPreview();
      } else if (action === 'preview-none') {
        state.previewHiddenIds = new Set(processedDs.map((d) => d.id));
        renderProcessedPreview();
      } else if (action === 'toggle-group-member') {
        toggleCompareGroupMembership(btn.dataset.dsId, btn.dataset.groupId);
      } else if (action === 'add-group') {
        addCompareGroup();
      } else if (action === 'compare-all') {
        replaceCompareSnapshotsWithAllProcessed();
        syncPreviewSectionVisibility();
      } else if (action === 'compare-none') {
        const group = getDefaultCompareGroup();
        group.snapshots = [];
        renderCompareArea();
        renderDatasetSelectPanel();
        syncPreviewSectionVisibility();
      }
    });
  }

  // Dataset switcher dropdown in preview sections (delegated)
  document.addEventListener('change', (event) => {
    if (!event.target.classList.contains('ds-switch-select')) return;
    const newId = event.target.value;
    if (!newId || newId === state.activeDatasetId) return;
    setActiveDataset(newId);
    state.csvSearchKeyword = '';
    state.csvSearchMode = 'manual';
    state.csvMatchedCols = [];
    state.csvChartSelectedCols = [];
    if (els.csvColSearch) els.csvColSearch.value = '';
    clearChart();
    updateMatchInfoAndPicker();
    updateSummary();
    renderPreview();
  });

  els.csvColSearchBtn.addEventListener('click', performColumnSearch);
  els.csvColResetBtn.addEventListener('click', resetColumnSearch);
  els.csvColSearch.addEventListener('keydown', (event) => { if (event.key === 'Enter') performColumnSearch(); });

  els.csvColCheckList.addEventListener('click', (event) => {
    const chip = event.target.closest('.col-chip');
    if (!chip) return;
    const c = parseInt(chip.dataset.col, 10);
    if (Number.isNaN(c)) return;
    const idx = state.csvChartSelectedCols.indexOf(c);
    if (idx >= 0) state.csvChartSelectedCols.splice(idx, 1);
    else state.csvChartSelectedCols.push(c);
    if (state.csvChartSelectedCols.length > 10) state.csvChartSelectedCols = state.csvChartSelectedCols.slice(-10);
    updateMatchInfoAndPicker();
  });

  els.generateChartBtn.addEventListener('click', generateChart);
  /* Auto-generate chart when chart type changes */
  const chartTypeEl = document.getElementById('chartType');
  if (chartTypeEl) chartTypeEl.addEventListener('change', generateChart);
  els.clearChartBtn.addEventListener('click', clearChart);
  els.chartResetBtn.addEventListener('click', resetCurrentChartZoom);

  // Metric selector toggle (电流/电压/功率)
  if (els.chartMetricSelector) {
    els.chartMetricSelector.addEventListener('click', (e) => {
      const btn = e.target.closest('.chart-metric-btn');
      if (!btn || btn.classList.contains('disabled-look')) return;
      const metric = btn.dataset.metric;
      if (!metric || metric === state.channelMetricType) return;
      state.channelMetricType = metric;
      // Update button active states
      els.chartMetricSelector.querySelectorAll('.chart-metric-btn').forEach((b) => b.classList.toggle('active', b.dataset.metric === metric));
      // Re-select columns based on new metric and regenerate chart
      const preferredCols = state.channelMetricCols[metric] || [];
      state.csvChartSelectedCols = preferredCols.slice(0, 6);
      updateMatchInfoAndPicker();
      if (preferredCols.length) {
        generateChart();
      } else {
        showChartPlaceholder(
          `未找到${METRIC_LABELS[metric]}数据`,
          `当前通道没有匹配到${METRIC_LABELS[metric]}列。可切换到其他指标继续查看。`
        );
      }
    });
  }

  els.processedApplyFilter.addEventListener('click', () => {
    applyProcessedColumnExpression(els.processedColRange.value);
  });
  els.processedShowAllBtn.addEventListener('click', () => {
    els.processedColRange.value = '';
    const maxCols = getColCount(state.processedRows.map((r) => r.map((v) => (v === undefined || v === null) ? '' : v)));
    state.processedColIndices = Array.from({ length: Math.max(maxCols, 1) }, (_, i) => i);
    renderProcessedPreview();
  });
  if (els.processedRestoreDefaultBtn) els.processedRestoreDefaultBtn.addEventListener('click', restoreDefaultProcessedColumns);
  if (els.processedSetDefaultBtn) els.processedSetDefaultBtn.addEventListener('click', saveDefaultProcessedColumns);

  // Compare
  els.processedAddCompareBtn.addEventListener('click', snapshotCurrentProcessed);
  if (els.compareAllBtn) els.compareAllBtn.addEventListener('click', snapshotAllProcessed);
  if (els.addCompareGroupBtn) els.addCompareGroupBtn.addEventListener('click', addCompareGroup);
  els.compareGroupsContainer.addEventListener('click', (event) => {
    const commonApplyBtn = event.target.closest('[data-action="cmp-common-apply"]');
    if (commonApplyBtn) { applyCompareCommonColExpr(); return; }
    const commonDefaultBtn = event.target.closest('[data-action="cmp-common-default"]');
    if (commonDefaultBtn) { state.compareCommonColIndices = null; renderCompareArea(); return; }
    const commonAllBtn = event.target.closest('[data-action="cmp-common-all"]');
    if (commonAllBtn) {
      state.compareCommonColIndices = Array.from({ length: COMPARE_COMMON_END }, (_, i) => i);
      renderCompareArea();
      return;
    }

    const projApplyBtn = event.target.closest('[data-action="cmp-project-apply"]');
    if (projApplyBtn) {
      applyProjectCommonColExpr(projApplyBtn.dataset.projectId);
      return;
    }
    const projDefaultBtn = event.target.closest('[data-action="cmp-project-default"]');
    if (projDefaultBtn) {
      const key = getProjectKey(projDefaultBtn.dataset.projectId);
      state.compareCommonColIndicesByProject[key] = null;
      renderCompareArea();
      return;
    }
    const projAllBtn = event.target.closest('[data-action="cmp-project-all"]');
    if (projAllBtn) {
      const key = getProjectKey(projAllBtn.dataset.projectId);
      state.compareCommonColIndicesByProject[key] = getProjectAvailableCommonCols(key);
      renderCompareArea();
      return;
    }
    const projResetBtn = event.target.closest('[data-action="cmp-project-reset-global"]');
    if (projResetBtn) {
      const key = getProjectKey(projResetBtn.dataset.projectId);
      if (state.compareCommonColIndicesByProject && Object.prototype.hasOwnProperty.call(state.compareCommonColIndicesByProject, key)) {
        delete state.compareCommonColIndicesByProject[key];
      }
      renderCompareArea();
      return;
    }

    const groupEl = event.target.closest('[data-group-id]');
    const groupId = groupEl ? groupEl.dataset.groupId : null;
    const crossCell = event.target.closest('[data-action="cross-select"]');
    if (crossCell) { toggleCrossProjectChannelSelection(crossCell.dataset.groupId || groupId, crossCell.dataset.snapId, crossCell.dataset.projectId, crossCell.dataset.channel, crossCell.dataset.colIndex); return; }
    const crossGenBtn = event.target.closest('[data-action="cross-generate"]');
    if (crossGenBtn) { generateCrossProjectChannelChart(crossGenBtn.dataset.groupId || groupId); return; }
    const crossClearBtn = event.target.closest('[data-action="cross-clear"]');
    if (crossClearBtn) { clearCrossProjectChannelSelection(crossClearBtn.dataset.groupId || groupId); renderCompareArea(); return; }
    const crossCloseBtn = event.target.closest('[data-action="cross-close"]');
    if (crossCloseBtn) { closeCrossProjectChannelChart(crossCloseBtn.dataset.groupId || groupId); return; }
    // Cross-project metric toggle
    const crossMetricBtn = event.target.closest('[data-cross-metric]');
    if (crossMetricBtn) {
      const gId = crossMetricBtn.dataset.groupId || groupId;
      const newMetric = crossMetricBtn.dataset.crossMetric;
      if (newMetric && newMetric !== (state.crossProjectMetricType[gId] || METRIC_CURRENT)) {
        state.crossProjectMetricType[gId] = newMetric;
        // Re-render chart shell to update title/meta and regenerate chart
        const panel = els.compareGroupsContainer.querySelector(`[data-cross-chart-panel="${CSS.escape(gId)}"]`);
        if (panel && !panel.classList.contains('hidden')) {
          // Update metric button UI
          panel.querySelectorAll('.chart-metric-btn[data-cross-metric]').forEach((b) => b.classList.toggle('active', b.dataset.crossMetric === newMetric));
          // Update meta text
          const metaEl = panel.querySelector('.cp-chart-meta');
          if (metaEl) metaEl.textContent = `X = 时间/第一列，最多绘制 ${CHART_MAX_POINTS} 点/文件；Y = ${METRIC_LABELS[newMetric]} (${METRIC_UNITS[newMetric]})`;
          generateCrossProjectChannelChart(gId);
        }
      }
      return;
    }
    const overlayCell = event.target.closest('[data-action="chart-overlay"]');
    if (overlayCell) { compareChannelWaveformsAcrossSnapshots(overlayCell.dataset.channel, groupId); return; }
    const chartCell = event.target.closest('[data-action="chart-channel"]');
    if (chartCell) { handleCompareChartClick(chartCell.dataset.snapid, chartCell.dataset.channel); return; }
    const removeBtn = event.target.closest('[data-action="remove-snap"]');
    if (removeBtn) { removeSnapshot(removeBtn.dataset.id); return; }
    const clearBtn = event.target.closest('[data-action="group-clear"]');
    if (clearBtn) { clearCompareGroup(clearBtn.dataset.groupId); return; }
    const removeGroupBtn = event.target.closest('[data-action="group-remove"]');
    if (removeGroupBtn) { removeCompareGroup(removeGroupBtn.dataset.groupId); return; }
    const renameEl = event.target.closest('[data-action="rename-group"]');
    if (renameEl) { renameCompareGroup(renameEl.dataset.groupId); return; }
  });
  els.compareGroupsContainer.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' && event.target.id === 'cmpCommonColRange') {
      event.preventDefault();
      applyCompareCommonColExpr();
    } else if (event.key === 'Enter' && event.target.dataset && event.target.dataset.role === 'cmp-project-input') {
      event.preventDefault();
      applyProjectCommonColExpr(event.target.dataset.projectId);
    }
  });
  els.compareGroupsContainer.addEventListener('change', (event) => {
    const hl = event.target.closest('input[data-action="group-highlight"]');
    if (hl) {
      const group = getCompareGroup(hl.dataset.groupId);
      if (group) { group.highlightDiff = !!hl.checked; renderCompareArea(); }
    }
  });

  document.body.addEventListener('click', (event) => {
    const closeBtn = event.target.closest('[data-role="htt-close"]');
    if (closeBtn) { hideHeaderTooltip(); return; }
    const anchor = event.target.closest('.th-truncate');
    if (anchor) { event.stopPropagation(); showHeaderTooltip(anchor); return; }
    if (!event.target.closest('#headerTooltip')) hideHeaderTooltip();
  });
  window.addEventListener('scroll', hideHeaderTooltip, true);
  window.addEventListener('resize', () => {
    hideHeaderTooltip();
    applyStickyColumns(document);
  });
}

registerChartZoomPlugin();
window.addEventListener('pageshow', (event) => {
  if (event.persisted) window.location.reload();
});
bindEvents();
window.__QEP_APP_READY__ = true;
clearTransientCaches();
setAdminMode(false);
updateMatchInfoAndPicker();
updateSummary();
renderPreview();
renderHistory();

window.addEventListener('power-auth-change', async (event) => {
  const detail = event.detail || {};
  state.serverUser = detail.user ? { username: detail.user.email, email: detail.user.email, id: detail.user.id } : null;
  setAdminMode(Boolean(detail.isAdmin));
  if (!detail.user) {
    state.templateItems = [];
    state.historyItems = [];
    renderHistory();
    return;
  }
  try {
    await Promise.all([refreshTemplateList(), loadHistoryFromCloud()]);
  } catch (error) {
    console.warn('Supabase data bootstrap failed', error);
  }
});
