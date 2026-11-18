/* ================================================================
 * 续航 / 功耗现状模块（endurance.js）
 * 纯前端、localStorage 持久化，不依赖 Supabase，游客模式可用。
 *
 * 续航模型为「公式驱动网格」（schema v2）：
 *  - 每个模型含若干对比列（使用模式），每列含 时长 / 耗电量 两子列；
 *  - 单元格 raw 为数字文本或以 = 开头的公式，由 formula.js 安全求值；
 *  - 默认口径对齐云表：场景耗电 = 时长 × 场景电流；
 *    DOU = 可用容量（Typ × 可放电比例）÷ 总耗电 × 总时长，容量不硬编码；
 *  - v1 旧模型（纯时长矩阵）加载时自动迁移为 v2。
 * ================================================================ */
(function () {
  'use strict';

  const STORAGE_KEY = 'endurance_status_v1';
  const CHART_COLORS = ['#6366f1', '#10b981', '#f59e0b', '#ef4444', '#0ea5e9', '#8b5cf6', '#ec4899', '#14b8a6', '#f97316', '#64748b'];
  // 图表实例统一登记，整卡重渲染前统一销毁，避免 Canvas 泄漏
  const chartInstances = [];
  function destroyCharts() {
    chartInstances.splice(0).forEach((c) => { try { c.destroy(); } catch (e) { /* ignore */ } });
  }

  // ---------- 基础工具 ----------
  function uid(prefix) {
    return (prefix || 'id') + '_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  }

  function escapeHtml(value) {
    return String(value === null || value === undefined ? '' : value)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function num(value) {
    if (value === null || value === undefined || value === '') return null;
    const n = Number(String(value).replace(/[,\s]/g, ''));
    return Number.isFinite(n) ? n : null;
  }

  function fmt(value, digits) {
    if (value === null || value === undefined || !Number.isFinite(value)) return '—';
    const d = digits === undefined ? 2 : digits;
    return Number(value.toFixed(d)).toLocaleString('zh-CN', { maximumFractionDigits: d });
  }

  function todayStr() {
    const d = new Date();
    const pad = (x) => String(x).padStart(2, '0');
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  }

  function debounce(fn, wait) {
    let timer = null;
    return function () {
      const args = arguments;
      clearTimeout(timer);
      timer = setTimeout(() => fn.apply(null, args), wait || 120);
    };
  }

  // ---------- 存储层 ----------
  // 电池可放电容量查询表（v60 起按电池规格书格式）：温度分组 × 放电倍率 × 各截止电压的累计可放电容量
  // { voltages: [3,3.1,3.2,3.3,3.4], rows: [{ id, temp:'25度', cRate:'0.2C', mah: { '3': 246, ... } }] }
  const DEFAULT_CUTOFF_VOLTAGES = [3, 3.1, 3.2, 3.3, 3.4];
  // 电压列：数值升序 + 去重，作为矩阵的统一列定义
  function matrixVoltages(matrix) {
    const list = ((matrix && matrix.voltages) || []).map((v) => num(v)).filter((v) => v !== null);
    const unique = [];
    list.forEach((v) => { if (unique.indexOf(v) === -1) unique.push(v); });
    return unique.sort((a, b) => a - b);
  }
  // 最低截止电压列（同一行里累计可放电最多的一格），用于规格书口径参考
  function matrixBaseVoltage(matrix) {
    const vs = matrixVoltages(matrix);
    return vs.length ? vs[0] : null;
  }
  function matrixCells(row) {
    return (row && row.mah && typeof row.mah === 'object') ? row.mah : {};
  }
  // 单元格取值：优先按数值 key，兼容导入时保留的字符串 key
  function matrixCellValue(row, voltage) {
    const cells = matrixCells(row);
    const v = num(voltage);
    if (v === null) return null;
    if (Object.prototype.hasOwnProperty.call(cells, String(v))) return num(cells[String(v)]);
    const key = Object.keys(cells).find((k) => num(k) === v);
    return key === undefined ? null : num(cells[key]);
  }
  function matrixSetCell(row, voltage, value) {
    if (!row.mah || typeof row.mah !== 'object') row.mah = {};
    row.mah[String(voltage)] = value;
  }
  function sampleDischargeMatrix() {
    // 示例数据 = 246mAh 电池（与 3.4V/0.95 比例的示例口径一致），按倍率递增逐格衰减
    const base = { 3: 246, 3.1: 243.5, 3.2: 240.5, 3.3: 238, 3.4: 233.7 };
    const rates = [
      ['0.2C', 1], ['0.5C', 0.965], ['1.0C', 0.93], ['1.5C', 0.9], ['2C', 0.875], ['4C', 0.82],
    ];
    return {
      voltages: DEFAULT_CUTOFF_VOLTAGES.slice(),
      rows: rates.map(([rate, k]) => {
        const mah = {};
        DEFAULT_CUTOFF_VOLTAGES.forEach((v) => { mah[String(v)] = Math.round(base[v] * k * 10) / 10; });
        return { id: uid('dt'), temp: '25度', cRate: rate, mah: mah };
      }),
    };
  }
  // 旧「电压 - 容量点」表 → 新的单行（25度 / 0.2C）矩阵，保证老数据不丢
  function matrixFromLegacyTable(rows) {
    const pts = sortDischargeTable(rows);
    if (!pts.length) return null;
    const voltages = pts.map((p) => p.voltage).sort((a, b) => a - b);
    const mah = {};
    pts.forEach((p) => { mah[String(p.voltage)] = p.capacityMah; });
    return { voltages: voltages, rows: [{ id: uid('dt'), temp: '25度', cRate: '0.2C', mah: mah }] };
  }
  // 取用于插值的曲线：优先 0.2C 行（比例口径就是 0.2C），其次第一行
  // 温度维度只做展示分组：同一张表内若存在多个温度，按 0.2C 行取第一条（默认 25 度），不自动跨温度换算
  function matrixCurve(matrix) {
    if (!matrix || !Array.isArray(matrix.rows) || !matrix.rows.length) return [];
    const row = matrix.rows.find((r) => String(r.cRate || '').toLowerCase().indexOf('0.2') === 0) || matrix.rows[0];
    return matrixVoltages(matrix).map((v) => ({ voltage: v, capacityMah: matrixCellValue(row, v) }))
      .filter((p) => p.capacityMah !== null);
  }

  // 查询表有效曲线：只要项目已经有矩阵，矩阵就是唯一口径（清空即无数据，不再回退旧表）；
  // 仅当对象里完全没有矩阵（异常残缺数据）时才读旧「电压-容量点」表
  function dischargeCurve(battery) {
    const b = battery || {};
    if (b.dischargeMatrix && typeof b.dischargeMatrix === 'object') return matrixCurve(b.dischargeMatrix);
    return sortDischargeTable(b.dischargeTable);
  }

  // 项目配置：关机电压、低电量门限属于整机口径，从电池参数里独立出来（v60）
  function defaultProjectConfig() {
    return { shutdownVoltage: '', lowThreshold: '' };
  }
  function projectConfig(project) {
    if (!project) return defaultProjectConfig();
    if (!project.config || typeof project.config !== 'object') project.config = defaultProjectConfig();
    const c = project.config;
    if (c.shutdownVoltage === undefined || c.shutdownVoltage === null) c.shutdownVoltage = '';
    if (c.lowThreshold === undefined || c.lowThreshold === null) c.lowThreshold = '';
    return c;
  }

  function defaultBattery() {
    return {
      capTyp: '', capMin: '', wh: '', voltage: '',
      dischargeRatio: '',
      // 电池可放电容量查询表：温度 × 放电倍率 × 截止电压矩阵（新项目为空表，示例项目由 seed 填充）
      dischargeMatrix: { voltages: DEFAULT_CUTOFF_VOLTAGES.slice(), rows: [] },
      // 兼容旧版本：单条「电压 - 容量点」曲线，仅作迁移来源
      dischargeTable: [],
    };
  }

  // 模型单元格默认公式：时长为 0 时直接得 0，避免缺电流场景把 #REF! 传播给 Total / DOU
  const DEFAULT_MAH_FORMULA = '=IF(时长=0,0,电流*时长)';
  const LEGACY_MAH_FORMULA = '=电流*时长';
  function defaultFoot(kind) {
    return {
      td: '=SUM(本列时长)',
      tm: '=SUM(本列耗电)',
      // avg 默认留空 → 渲染自动平均电流（总耗电 ÷ 总时长）；填入数字 / 公式后即被自定义覆盖
      avg: '',
      // 汇总区备注列（E/H/K…）的四个物理格：各自独立，允许数字 / 文本 / 公式，不再渲染成不可编辑灰块
      noteTotal: '',
      noteAvg: '',
      noteTyp: '',
      noteMin: '',
      // 平均/DOU 行的 C 列物理格：默认空白，但用户可自行写入数字 / 文本 / 公式
      blankAvg: '',
      blankTyp: '',
      blankMin: '',
      douTyp: kind === 'full' ? '=Typ/总耗电*总时长' : '=可用容量/总耗电*总时长',
      douMin: kind === 'full' ? '=Min/总耗电*总时长' : '=Min*比例/总耗电*总时长',
    };
  }

  // 汇总标签默认文案（v57 定稿口径）：平均电流_mA / 模型续航时长 Typ_H / 模型续航时长 Min_H。
  // 用户手动改过的标签不会被覆盖，只把仍是历史默认文案的自动换成新默认。
  const LEGACY_FOOTER_LABELS = {
    total: ['Total（可编辑公式）'],
    avg: ['平均电流 mA（自动）', '平均电流'],
    douTyp: ['可使用时长（Typ）h', '可用时长 Typ'],
    douMin: ['可使用时长（Min）h', '可用时长 Min'],
  };

  function defaultModelFooter() {
    return {
      labels: {
        total: 'Total',
        avg: '平均电流_mA',
        douTyp: '模型续航时长 Typ_H',
        douMin: '模型续航时长 Min_H',
      },
      current: { total: '', avg: '', douTyp: '', douMin: '' },
    };
  }

  function createColumn(name, kind) {
    const col = {
      id: uid('md'), name: name, kind: kind === 'full' ? 'full' : 'normal',
      window: '', note: '', needMah: null,
    };
    return col;
  }

  // 依据项目场景构造一个 v2 空模型
  function createModelV2(project, name, columnDefs) {
    const columns = (columnDefs && columnDefs.length ? columnDefs : [{ name: '典型使用', kind: 'normal' }])
      .map((def) => createColumn(def.name, def.kind));
    const foot = {};
    columns.forEach((col) => { foot[col.id] = defaultFoot(col.kind); });
    const rows = project.scenarios.map((scenario) => {
      const cells = {};
      columns.forEach((col) => { cells[col.id] = { d: '0', m: DEFAULT_MAH_FORMULA }; });
      return { id: uid('mr'), scenarioId: scenario.id, label: '', currentRaw: '', note: '', cells: cells };
    });
    return {
      v: 3, footerVersion: 2, id: uid('m'), name: name, columns: columns, rows: rows,
      foot: foot, footer: defaultModelFooter(),
    };
  }

  // v1（纯时长矩阵）模型迁移为 v2 公式网格；保留原 mode id 以减少外部引用断裂
  function migrateModelV1(model) {
    if (model.v === 2) return model;
    const columns = (model.modes || []).map((mode) => ({
      id: mode.id,
      name: mode.name,
      kind: mode.fullOnly ? 'full' : 'normal',
      window: mode.window || '',
      note: mode.note || '',
      needMah: mode.needMah === undefined ? null : mode.needMah,
    }));
    const foot = {};
    columns.forEach((col) => { foot[col.id] = defaultFoot(col.kind); });
    const rows = (model.rows || []).map((oldRow) => {
      const cells = {};
      columns.forEach((col) => {
        const raw = oldRow.durations && oldRow.durations[col.id] !== undefined && oldRow.durations[col.id] !== null
          ? String(oldRow.durations[col.id]) : '0';
        cells[col.id] = { d: raw, m: DEFAULT_MAH_FORMULA };
      });
      return { id: uid('mr'), scenarioId: oldRow.scenarioId, label: '', currentRaw: '', note: oldRow.note || '', cells: cells };
    });
    return {
      v: 2, footerVersion: 2, id: model.id, name: model.name, columns: columns, rows: rows,
      foot: foot, footer: defaultModelFooter(),
    };
  }

  // 云表 HrV2LC「续航模型」默认 12 场景及其场景电流（K 列，mA）
  // [场景名, 场景电流 mA, UT原始模型(3H) 时长, 环境记忆常开 时长, 重度模型 时长, 备注]
  function defaultScenarioDefs() {
    return [
      ['音乐', 22.82, 1.5, 0, 0, ''],
      ['音乐+环境记忆', 27.44, 0, 1.5, 1.5, ''],
      ['蓝牙通话', 21.8, 2, 2, 2, ''],
      ['单次对话+环境', 141.5, 0.083333, 0.083333, 0.166667, '单次 5 分钟'],
      ['视频通话（wifi RTC）', 105.76, 0.25, 0.25, 0.5, '单次 15 分钟'],
      ['拍照', 91.75, 0.03, 0.03, 0.03, '拍照 10 次'],
      ['录像', 198.41, 0.05, 0.05, 0.05, '录像 3 分钟'],
      ['录音纪要', 12.55, 0.2, 0.2, 0.4, '12 分钟'],
      ['同声传译', 17.99, 0, 0, 0, ''],
      ['待机+环境-慢档', 17.52, 2, 3.924444, 3.568889, '3H 环境记忆，慢档：中档=2:1'],
      ['待机+环境-中档', 19.5, 1, 1.962222, 1.784444, ''],
      ['待机', 10.31, 2.886667, 0, 0, ''],
    ];
  }

  // 构建默认 12 场景，电流直接作为一条 UT 测试记录附上（新项目 / 空项目一键载入共用）
  function buildDefaultScenarios(version, note) {
    return defaultScenarioDefs().map((def) => ({
      id: uid('sc'),
      name: def[0],
      remark: def[5],
      standard: '',
      customCurrent: '',
      fields: {},
      singleLife: '',
      records: [{ id: uid('rec'), date: todayStr(), version: version || 'UT', current: def[1], note: note || '云表默认场景电流' }],
    }));
  }

  // 把云表 12 个参考场景合并进项目：同名场景原样保留、不覆盖，仅补齐缺失项；
  // 新增场景同步给项目内已有模型补行。返回新增 / 跳过明细，供确认弹窗与提示使用。
  function mergeDefaultScenarios(project, options) {
    const version = (options && options.version) || 'UT';
    const note = (options && options.note) || '云表默认场景电流';
    const existingNames = new Set((project.scenarios || []).map((s) => s.name));
    const added = [];
    const skipped = [];
    const created = [];
    defaultScenarioDefs().forEach((def) => {
      if (existingNames.has(def[0])) { skipped.push(def[0]); return; }
      const scenario = {
        id: uid('sc'),
        name: def[0],
        remark: def[5],
        standard: '',
        customCurrent: '',
        fields: {},
        singleLife: '',
        records: [{ id: uid('rec'), date: todayStr(), version: version, current: def[1], note: note }],
      };
      project.scenarios.push(scenario);
      created.push(scenario);
      added.push(def[0]);
    });
    (project.models || []).forEach((model) => {
      created.forEach((scenario) => {
        if (model.rows.some((r) => r.scenarioId === scenario.id)) return;
        const cells = {};
        model.columns.forEach((col) => { cells[col.id] = { d: '0', m: DEFAULT_MAH_FORMULA }; });
        model.rows.push({ id: uid('mr'), scenarioId: scenario.id, label: '', currentRaw: '', note: '', cells: cells });
      });
    });
    return { added: added, skipped: skipped, created: created };
  }

  // ---------- 场景电流 vs 功耗标准（方案 A 徽标） ----------
  // 达标(≤标准)绿色 ✓；超标 0~20% 橙色 ⚠；超标 >20% 红色 ✗。无标准 / 无电流返回 null。
  function currentCompare(current, standard) {
    const cur = num(current);
    const std = num(standard);
    if (cur === null || std === null || std <= 0) return null;
    const pct = (cur - std) / std * 100;
    const level = pct <= 0 ? 'ok' : (pct <= 20 ? 'warn' : 'bad');
    const rounded = Math.round(pct);
    let shown;
    if (rounded === 0 && pct !== 0) {
      shown = (pct > 0 ? '+' : '−') + Math.abs(pct).toFixed(1) + '%';
    } else {
      shown = (rounded > 0 ? '+' : rounded < 0 ? '−' : '') + Math.abs(rounded) + '%';
    }
    const icon = level === 'ok' ? '✓' : (level === 'warn' ? '⚠' : '✗');
    return {
      level: level,
      pct: pct,
      text: icon + ' ' + shown,
      title: '当前电流相对功耗标准 ' + (pct >= 0 ? '+' : '−') + Math.abs(pct).toFixed(1)
        + '%；达标（≤标准）绿色，超标 0~20% 橙色，超标 >20% 红色',
    };
  }

  // ---------- DOU 网格布局偏好（列宽 / 行高），独立 localStorage，不与业务数据混存 ----------
  const GRID_LAYOUT_KEY = 'endurance_grid_layout_v1';
  const GRID_COL_ROLES = ['rownum', 'scenario', 'current', 'duration', 'mah', 'note'];
  // 桌面默认列宽（px）：紧凑、按内容收缩，不强制撑满容器；note 为电流列之后的备注/说明列
  const GRID_DEFAULT_COLW = { rownum: 38, scenario: 150, current: 82, duration: 94, mah: 102, note: 160 };
  const GRID_MOBILE_COLW = { rownum: 30, scenario: 116, current: 64, duration: 78, mah: 86, note: 120 };
  const GRID_COL_MIN = { rownum: 28, scenario: 88, current: 56, duration: 60, mah: 64, note: 80 };
  const GRID_COL_MAX = { rownum: 96, scenario: 320, current: 170, duration: 220, mah: 240, note: 340 };
  // 三档行高（px），用户可在工具栏切换
  const GRID_ROW_H = { compact: 30, normal: 38, comfortable: 48 };
  const GRID_DEFAULT_ROWH = 'normal';

  function defaultGridColWidths() {
    const mobile = typeof window !== 'undefined' && window.innerWidth && window.innerWidth <= 640;
    return Object.assign({}, mobile ? GRID_MOBILE_COLW : GRID_DEFAULT_COLW);
  }
  function defaultGridLayout() {
    return { colWidths: defaultGridColWidths(), rowHeight: GRID_DEFAULT_ROWH };
  }
  function loadGridLayout() {
    const base = defaultGridLayout();
    try {
      const raw = localStorage.getItem(GRID_LAYOUT_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        const cw = parsed && parsed.colWidths && typeof parsed.colWidths === 'object' ? parsed.colWidths : {};
        GRID_COL_ROLES.forEach((role) => {
          const v = Number(cw[role]);
          if (Number.isFinite(v) && v > 0) base.colWidths[role] = Math.round(v);
        });
        if (parsed && Object.prototype.hasOwnProperty.call(GRID_ROW_H, parsed.rowHeight)) base.rowHeight = parsed.rowHeight;
      }
    } catch (error) { /* 损坏的偏好直接回退默认 */ }
    return base;
  }
  function saveGridLayout(patch) {
    const cur = loadGridLayout();
    const nextRowH = patch && Object.prototype.hasOwnProperty.call(GRID_ROW_H, patch.rowHeight)
      ? patch.rowHeight
      : cur.rowHeight;
    const next = {
      colWidths: Object.assign({}, cur.colWidths, (patch && patch.colWidths) || {}),
      rowHeight: nextRowH,
    };
    try { localStorage.setItem(GRID_LAYOUT_KEY, JSON.stringify(next)); } catch (error) { /* ignore */ }
    return next;
  }
  function resetGridLayout() {
    try { localStorage.removeItem(GRID_LAYOUT_KEY); } catch (error) { /* ignore */ }
    return defaultGridLayout();
  }
  function clampGridColWidth(role, width) {
    const min = GRID_COL_MIN[role] || 48;
    const max = GRID_COL_MAX[role] || 260;
    return Math.max(min, Math.min(max, Math.round(width)));
  }

  // 首次使用的示例项目：对齐云表「G2 标准版续航模型 V1.0（UT 阶段）」
  function seedData() {
    // 时长取自云表 HrV2LC「续航模型」E/G/I 列，保留 6 位小数以贴近原表汇总
    const scenarioDefs = defaultScenarioDefs();
    const scenarios = buildDefaultScenarios('UT', '示例数据');
    const byName = {};
    scenarios.forEach((s) => { byName[s.name] = s.id; });

    const project = {
      id: uid('p'),
      name: '示例项目（可删除）',
      createdAt: Date.now(),
      scenarioFields: [],
      // 项目配置：关机电压 / 低电量门限（与电池参数解耦）
      config: { shutdownVoltage: '3.4', lowThreshold: '0.02' },
      battery: Object.assign(defaultBattery(), {
        capTyp: 246, capMin: 240, wh: 0.95, voltage: 3.87,
        dischargeRatio: 0.95,
        dischargeMatrix: sampleDischargeMatrix(),
      }),
      scenarios: scenarios,
      models: [],
      activeModelId: null,
    };

    const columnDefs = [
      { name: 'UT原始模型（环境记忆3H）', kind: 'normal' },
      { name: '基于UT模型（环境记忆常开）', kind: 'normal' },
      { name: '重度模型（环境记忆常开）', kind: 'normal' },
    ];
    const model = createModelV2(project, 'G2 标准版续航模型 V1.0（UT 阶段）', columnDefs);
    scenarioDefs.forEach((def, i) => {
      const row = model.rows[i];
      model.columns.forEach((col, ci) => {
        row.cells[col.id].d = String(def[2 + ci]);
      });
    });
    project.models = [model];
    project.activeModelId = model.id;
    return { activeProjectId: project.id, projects: [project] };
  }

  // 旧版本数据归一化：补齐新字段、迁移 v1 模型、移除已下线电池字段
  function normalizeData(data) {
    let changed = false;
    (data.projects || []).forEach((project) => {
      const b = project.battery || (project.battery = defaultBattery());
      if ('envStandbyCurrent' in b) { delete b.envStandbyCurrent; changed = true; }
      if ('wakeCurrent' in b) { delete b.wakeCurrent; changed = true; }
      // 关机电压 / 低电量门限 → 项目配置（v60 迁移，值不丢）
      let freshConfig = false;
      if (!project.config || typeof project.config !== 'object') {
        project.config = defaultProjectConfig();
        freshConfig = true;
        changed = true;
      }
      if (freshConfig) {
        project.config.shutdownVoltage = (b.shutdownVoltage === undefined || b.shutdownVoltage === null) ? '' : b.shutdownVoltage;
        project.config.lowThreshold = (b.lowThreshold === undefined || b.lowThreshold === null) ? '' : b.lowThreshold;
      }
      if ('shutdownVoltage' in b) { delete b.shutdownVoltage; changed = true; }
      if ('lowThreshold' in b) { delete b.lowThreshold; changed = true; }
      if (!Array.isArray(b.dischargeTable)) { b.dischargeTable = []; changed = true; }
      if (!b.dischargeMatrix || typeof b.dischargeMatrix !== 'object' || !Array.isArray(b.dischargeMatrix.rows)) {
        // 老项目：旧点表迁成单行（25度 / 0.2C）；迁移完成后删除旧表，避免清空矩阵后旧数据又冒出来
        b.dischargeMatrix = matrixFromLegacyTable(b.dischargeTable)
          || { voltages: DEFAULT_CUTOFF_VOLTAGES.slice(), rows: [] };
        b.dischargeTable = [];
        changed = true;
      } else if (!Array.isArray(b.dischargeMatrix.voltages) || !b.dischargeMatrix.voltages.length) {
        b.dischargeMatrix.voltages = DEFAULT_CUTOFF_VOLTAGES.slice();
        changed = true;
      }
      (b.dischargeMatrix.rows || []).forEach((row) => {
        if (!row.id) { row.id = uid('dt'); changed = true; }
        if (!row.mah || typeof row.mah !== 'object') { row.mah = {}; changed = true; }
      });
      if (!Array.isArray(project.scenarioFields)) { project.scenarioFields = []; changed = true; }
      (project.scenarios || []).forEach((scenario) => {
        if (!('standard' in scenario)) { scenario.standard = ''; changed = true; }
        if (!('customCurrent' in scenario)) { scenario.customCurrent = ''; changed = true; }
        if (!('singleLife' in scenario)) { scenario.singleLife = ''; changed = true; }
        if (!scenario.fields || typeof scenario.fields !== 'object') { scenario.fields = {}; changed = true; }
        if (!Array.isArray(scenario.records)) scenario.records = [];
      });
      if (Array.isArray(project.models)) {
        project.models = project.models.map((model) => {
          // v1（无 v 或 v=1）先迁移为 v2 结构；v2/v3 原样保留
          const migrated = model.v === 2 || model.v === 3 ? model : migrateModelV1(model);
          // 一次性迁移：旧默认耗电公式升级为带 0 时长保护的 IF 公式（用户自定义公式保持不动）
          (migrated.rows || []).forEach((row) => {
            // v52 全开放：补齐自由电流（B 列覆盖）与自由行标题字段
            if (!('currentRaw' in row)) { row.currentRaw = ''; changed = true; }
            if (!('label' in row)) { row.label = ''; changed = true; }
            Object.keys(row.cells || {}).forEach((colId) => {
              if (row.cells[colId] && row.cells[colId].m === LEGACY_MAH_FORMULA) {
                row.cells[colId].m = DEFAULT_MAH_FORMULA;
                changed = true;
              }
            });
          });
          // v2 虚拟地址 → v3 连续物理地址（重写裸 A1 引用；命名公式无裸地址，仅升版本号）
          if (migrated.v !== 3) {
            migrateModelAddressesV2ToV3(migrated);
            changed = true;
          }
          // 自由编辑改造：补齐汇总脚注，老数据可能整个 foot 项缺失或没有 avg（自定义平均电流）字段
          if (!migrated.foot || typeof migrated.foot !== 'object') { migrated.foot = {}; changed = true; }
          (migrated.columns || []).forEach((col) => {
            if (!migrated.foot[col.id]) { migrated.foot[col.id] = defaultFoot(col.kind); changed = true; }
            else {
              if (!('avg' in migrated.foot[col.id])) { migrated.foot[col.id].avg = ''; changed = true; }
              ['noteTotal', 'noteAvg', 'noteTyp', 'noteMin', 'blankAvg', 'blankTyp', 'blankMin'].forEach((key) => {
                if (!(key in migrated.foot[col.id])) { migrated.foot[col.id][key] = ''; changed = true; }
              });
            }
          });
          const defaultFooter = defaultModelFooter();
          if (!migrated.footer || typeof migrated.footer !== 'object') {
            migrated.footer = defaultFooter;
            changed = true;
          } else {
            if (!migrated.footer.labels || typeof migrated.footer.labels !== 'object') {
              migrated.footer.labels = { ...defaultFooter.labels };
              changed = true;
            } else {
              Object.keys(defaultFooter.labels).forEach((key) => {
                if (!(key in migrated.footer.labels)) { migrated.footer.labels[key] = defaultFooter.labels[key]; changed = true; }
                else if ((LEGACY_FOOTER_LABELS[key] || []).indexOf(migrated.footer.labels[key]) >= 0) {
                  migrated.footer.labels[key] = defaultFooter.labels[key];
                  changed = true;
                }
              });
            }
            if (!migrated.footer.current || typeof migrated.footer.current !== 'object') {
              migrated.footer.current = { ...defaultFooter.current };
              changed = true;
            } else {
              Object.keys(defaultFooter.current).forEach((key) => {
                if (!(key in migrated.footer.current)) { migrated.footer.current[key] = ''; changed = true; }
              });
            }
          }
          if (migrated.footerVersion !== 2) {
            migrateFooterAddressesV52ToV53(migrated);
            migrated.footerVersion = 2;
            changed = true;
          }
          return migrated;
        });
      } else {
        project.models = [];
        changed = true;
      }
    });
    if (changed) {
      try { localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); } catch (error) { /* ignore */ }
    }
    return data;
  }

  // 查询表按电压降序排序，并对同电压点去重（保留后出现的行）；仅处理电压/容量均有效的点，保留行 id
  function sortDischargeTable(rows) {
    const cleaned = (rows || []).filter((r) => num(r.voltage) !== null && num(r.capacityMah) !== null);
    const map = new Map();
    cleaned.forEach((r) => {
      map.set(String(num(r.voltage)), { id: r.id, voltage: num(r.voltage), capacityMah: num(r.capacityMah) });
    });
    return Array.from(map.values()).sort((a, b) => b.voltage - a.voltage);
  }

  // 按关机电压在查询表中夹取 + 线性插值，返回累计可放电容量 mAh；表为空或无有效点返回 null
  function lookupDischargeCapacity(rows, voltage) {
    const v = num(voltage);
    if (v === null) return null;
    const pts = sortDischargeTable(rows);
    if (!pts.length) return null;
    if (v >= pts[0].voltage) return pts[0].capacityMah;
    if (v <= pts[pts.length - 1].voltage) return pts[pts.length - 1].capacityMah;
    for (let i = 0; i < pts.length - 1; i += 1) {
      const hi = pts[i];
      const lo = pts[i + 1];
      if (v <= hi.voltage && v >= lo.voltage) {
        const span = hi.voltage - lo.voltage;
        const t = span === 0 ? 0 : (hi.voltage - v) / span;
        return hi.capacityMah + (lo.capacityMah - hi.capacityMah) * t;
      }
    }
    return null;
  }

  // 依据关机电压 + 查询表 + Typ 容量推导可放电容量比例；信息不足时返回 null 并带原因
  function deriveDischargeRatio(project) {
    const b = (project && project.battery) || {};
    const cfg = projectConfig(project);
    const table = dischargeCurve(b);
    if (!table.length) return { ok: false, reason: 'empty-table' };
    const shutdown = num(cfg.shutdownVoltage);
    if (shutdown === null) return { ok: false, reason: 'no-shutdown-voltage' };
    const capacityMah = lookupDischargeCapacity(table, shutdown);
    if (capacityMah === null) return { ok: false, reason: 'lookup-failed' };
    const capTyp = num(b.capTyp);
    if (capTyp === null || capTyp <= 0) return { ok: true, capacityMah: capacityMah, ratio: null, reason: 'no-cap-typ' };
    const ratio = capacityMah / capTyp;
    return { ok: true, capacityMah: capacityMah, ratio: ratio, reason: '' };
  }

  const store = {
    data: null,
    load() {
      if (this.data) return this.data;
      try {
        const raw = localStorage.getItem(STORAGE_KEY);
        this.data = raw ? JSON.parse(raw) : null;
      } catch (error) {
        console.warn('续航现状数据读取失败', error);
        this.data = null;
      }
      if (!this.data || !Array.isArray(this.data.projects)) {
        this.data = seedData();
        this.save();
      } else {
        normalizeData(this.data);
      }
      return this.data;
    },
    save() {
      try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(this.data));
      } catch (error) {
        console.warn('续航现状数据保存失败（可能超出浏览器存储上限）', error);
        if (window.ES_MODAL) {
          window.ES_MODAL.alert({ title: '保存失败', icon: 'error', message: '浏览器存储空间可能已满，请删除部分项目或历史记录后重试。' });
        }
      }
    },
    activeProject() {
      const data = this.load();
      return data.projects.find((p) => p.id === data.activeProjectId) || data.projects[0] || null;
    },
    setActive(id) {
      this.data.activeProjectId = id;
      this.save();
    },
    addProject(name) {
      const project = {
        id: uid('p'),
        name: name,
        createdAt: Date.now(),
        scenarioFields: [],
        config: defaultProjectConfig(),
        battery: defaultBattery(),
        // 新项目默认载入云表 12 个标准场景及对应场景电流，开箱即可计算
        scenarios: buildDefaultScenarios(),
        models: [],
        activeModelId: null,
      };
      this.data.projects.push(project);
      this.data.activeProjectId = project.id;
      this.save();
      return project;
    },
    removeProject(id) {
      this.data.projects = this.data.projects.filter((p) => p.id !== id);
      if (this.data.activeProjectId === id) {
        this.data.activeProjectId = this.data.projects[0] ? this.data.projects[0].id : null;
      }
      this.save();
    },
    // 导入一个项目：先全量重生成内部 id（保证与现有项目零碰撞、可重复导入为独立副本），
    // 再按同名规则改名，最后复用 normalizeData 补齐字段 / 迁移 v1 模型；绝不静默覆盖现有数据。
    importProject(projectLike) {
      const data = this.load();
      const reidentified = reidentifyProject(projectLike);
      reidentified.name = uniqueProjectName(reidentified.name, data.projects.map((p) => p.name));
      data.projects.push(reidentified);
      data.activeProjectId = reidentified.id;
      normalizeData(data);
      this.save();
      return reidentified;
    },
  };

  // ---------- 项目导入 / 导出（纯本地、零云端依赖，用于多设备之间手工迁移）----------
  const EXPORT_APP = 'power-analysis';
  const EXPORT_KIND = 'endurance-project';
  const EXPORT_FORMAT = 1;

  function deepClone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  // 文件名清洗：去除文件系统非法字符与控制符、折叠空白；为空时回退到 fallback
  function sanitizeFileName(name, fallback) {
    const cleaned = String(name === null || name === undefined ? '' : name)
      .replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .slice(0, 80);
    return cleaned || fallback || '续航项目';
  }

  // 导出文件名时间戳：YYYYMMDD-HHmm
  function fileTimestamp(date) {
    const d = date || new Date();
    const pad = (x) => String(x).padStart(2, '0');
    return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate())
      + '-' + pad(d.getHours()) + pad(d.getMinutes());
  }

  // 构造导出载荷：深拷贝项目并附带格式标识 / 时间，导入端据此校验
  function buildProjectExport(project) {
    return {
      app: EXPORT_APP,
      module: 'endurance',
      kind: EXPORT_KIND,
      formatVersion: EXPORT_FORMAT,
      exportedAt: new Date().toISOString(),
      project: deepClone(project),
    };
  }

  // 项目结构基本校验：必须有名称，场景 / 模型 / 自定义字段若存在必须是数组
  function isValidProjectShape(project) {
    if (!project || typeof project !== 'object') return false;
    if (!project.name || typeof project.name !== 'string') return false;
    if (project.scenarios !== undefined && !Array.isArray(project.scenarios)) return false;
    if (project.models !== undefined && !Array.isArray(project.models)) return false;
    if (project.scenarioFields !== undefined && !Array.isArray(project.scenarioFields)) return false;
    return true;
  }

  // 解析并严格校验导入文本；兼容「包装载荷」与「裸 project」两种形态
  function parseProjectImport(text) {
    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch (error) {
      return { ok: false, error: '文件不是合法的 JSON，请确认选择的是本工具导出的项目文件（.json）。' };
    }
    if (!parsed || typeof parsed !== 'object') {
      return { ok: false, error: '文件内容结构不正确，未找到续航项目数据。' };
    }
    const project = parsed.project && typeof parsed.project === 'object' ? parsed.project : parsed;
    if (!isValidProjectShape(project)) {
      return { ok: false, error: '未在文件中找到有效的续航项目（缺少项目名称，或场景 / 模型数据结构不正确）。' };
    }
    return { ok: true, project: deepClone(project) };
  }

  // 同名去重：已存在则追加「（导入 n）」后缀，保证项目名可区分
  function uniqueProjectName(baseName, existingNames) {
    const base = String(baseName || '').trim() || '导入项目';
    const used = new Set((existingNames || []).map((n) => String(n)));
    if (!used.has(base)) return base;
    let n = 1;
    while (used.has(base + '（导入 ' + n + '）')) n += 1;
    return base + '（导入 ' + n + '）';
  }

  // 全量重生成项目内部 id 并维护引用映射：
  // 场景 id 被模型行 scenarioId 引用；列 id 被 cells / foot 的 key 引用；自定义字段 id 被 scenario.fields 的 key 引用。
  // 重生成后项目内部引用保持完整，且与任何现有项目零 id 碰撞，同一文件可重复导入为互不影响的副本。
  function reidentifyProject(input) {
    const p = deepClone(input);
    p.id = uid('p');

    // 1) 自定义字段 id（被各 scenario.fields 的 key 引用）
    const fieldIdMap = {};
    p.scenarioFields = (Array.isArray(p.scenarioFields) ? p.scenarioFields : []).map((f) => {
      const next = Object.assign({}, f, { id: uid('sf') });
      if (f.id !== undefined && f.id !== null) fieldIdMap[String(f.id)] = next.id;
      return next;
    });

    // 2) 电池可放电容量查询表点 / 矩阵行 id（无外部引用，重生成保持 id 体系统一）
    if (p.battery && Array.isArray(p.battery.dischargeTable)) {
      p.battery.dischargeTable = p.battery.dischargeTable.map((d) => Object.assign({}, d, { id: uid('dt') }));
    }
    if (p.battery && p.battery.dischargeMatrix && Array.isArray(p.battery.dischargeMatrix.rows)) {
      p.battery.dischargeMatrix = Object.assign({}, p.battery.dischargeMatrix, {
        rows: p.battery.dischargeMatrix.rows.map((r) => Object.assign({}, r, { id: uid('dt'), mah: Object.assign({}, r.mah) })),
      });
    }

    // 3) 场景 id + 测试记录 id + 场景内自定义字段 key 重映射
    const scenarioIdMap = {};
    p.scenarios = (Array.isArray(p.scenarios) ? p.scenarios : []).map((s) => {
      const next = Object.assign({}, s);
      next.id = uid('sc');
      if (s.id !== undefined && s.id !== null) scenarioIdMap[String(s.id)] = next.id;
      next.records = (Array.isArray(s.records) ? s.records : []).map((r) => Object.assign({}, r, { id: uid('rec') }));
      if (s.fields && typeof s.fields === 'object') {
        const remapped = {};
        Object.keys(s.fields).forEach((oldKey) => {
          const newKey = Object.prototype.hasOwnProperty.call(fieldIdMap, oldKey) ? fieldIdMap[oldKey] : oldKey;
          remapped[newKey] = s.fields[oldKey];
        });
        next.fields = remapped;
      }
      return next;
    });

    // 4) 模型：model / 列（或 v1 的 mode）/ 行 id，foot 与 cells、v1 durations 的 key、行 scenarioId 全部重映射
    const oldModels = Array.isArray(input.models) ? input.models : [];
    p.models = (Array.isArray(p.models) ? p.models : []).map((m) => {
      const model = Object.assign({}, m);
      model.id = uid('m');
      const colIdMap = {};
      if (Array.isArray(m.columns)) {
        // v2：重生成对比列 id
        model.columns = m.columns.map((c) => {
          const next = Object.assign({}, c, { id: uid('md') });
          colIdMap[String(c.id)] = next.id;
          return next;
        });
      } else if (Array.isArray(m.modes)) {
        // v1：重生成 mode id（迁移为 v2 后即列 id），并同步 rows.durations 的 key
        model.modes = m.modes.map((mode) => {
          const next = Object.assign({}, mode, { id: uid('md') });
          colIdMap[String(mode.id)] = next.id;
          return next;
        });
      }
      const mapKey = (map, oldKey) => (Object.prototype.hasOwnProperty.call(map, oldKey) ? map[oldKey] : oldKey);
      if (model.foot && typeof model.foot === 'object') {
        const remappedFoot = {};
        Object.keys(model.foot).forEach((oldKey) => { remappedFoot[mapKey(colIdMap, oldKey)] = model.foot[oldKey]; });
        model.foot = remappedFoot;
      }
      model.rows = (Array.isArray(m.rows) ? m.rows : []).map((r) => {
        const next = { id: uid('mr'), note: r.note || '', currentRaw: r.currentRaw || '', label: r.label || '' };
        next.scenarioId = mapKey(scenarioIdMap, String(r.scenarioId)) !== String(r.scenarioId)
          ? scenarioIdMap[String(r.scenarioId)] : r.scenarioId;
        if (r.cells && typeof r.cells === 'object') {
          const cells = {};
          Object.keys(r.cells).forEach((oldKey) => { cells[mapKey(colIdMap, oldKey)] = r.cells[oldKey]; });
          next.cells = cells;
        }
        // v1 时长矩阵：保留 durations 并重映射其 mode key，供后续 migrateModelV1 取数
        if (r.durations && typeof r.durations === 'object') {
          const durations = {};
          Object.keys(r.durations).forEach((oldKey) => { durations[mapKey(colIdMap, oldKey)] = r.durations[oldKey]; });
          next.durations = durations;
        }
        return next;
      });
      return model;
    });

    // 5) activeModelId 按旧数组的同位置映射到新模型 id；找不到则回退首个模型
    const oldActiveIndex = oldModels.findIndex((m) => String(m.id) === String(input.activeModelId));
    p.activeModelId = oldActiveIndex >= 0 && p.models[oldActiveIndex]
      ? p.models[oldActiveIndex].id
      : (p.models[0] ? p.models[0].id : null);

    return p;
  }

  // 触发浏览器下载一个文本文件（JSON），用完即释放对象 URL
  function downloadTextFile(filename, text, mime) {
    const blob = new Blob([text], { type: (mime || 'application/json') + ';charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    if (anchor.parentNode) anchor.parentNode.removeChild(anchor);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  // ---------- 业务查询 ----------
  // 场景当前生效电流：自定义电流优先，否则取最新一条记录（按日期，同日取最后录入）
  function latestRecord(scenario) {
    if (!scenario || !Array.isArray(scenario.records) || !scenario.records.length) return null;
    return scenario.records.slice().sort((a, b) => {
      const cmp = String(a.date).localeCompare(String(b.date));
      return cmp !== 0 ? cmp : 0;
    }).slice(-1)[0];
  }

  function scenarioCurrent(project, scenarioId) {
    const scenario = project.scenarios.find((s) => s.id === scenarioId);
    if (!scenario) return null;
    const custom = num(scenario.customCurrent);
    if (custom !== null) return custom;
    const rec = latestRecord(scenario);
    return rec ? num(rec.current) : null;
  }

  // ---------- A1 网格映射（v3：物理列与公式地址统一，列标连续无空白）----------
  // 物理/公式列：1=A 场景名，2=B 电流；第 i 个对比列（0-based）模块占三列：
  //   时长列 (3+3i) / 耗电列 (4+3i) / 备注列 (5+3i，可被引用，文本参与数值运算时记 #VALUE!)。
  // 第 1 行表头，场景行从第 2 行开始，其后依次为 Total、平均电流、Typ、Min 四个连续汇总行。
  function colLetterFor(colIndex, cellType) {
    return window.ES_FORMULA.colToLetter(3 + colIndex * 3 + (cellType === 'm' ? 1 : 0));
  }
  // 备注列字母（v52 起备注可被 A1 引用，也支持数字或 = 公式）
  function noteLetterFor(colIndex) {
    return window.ES_FORMULA.colToLetter(5 + colIndex * 3);
  }
  function rowNumberFor(rowIndex) { return 2 + rowIndex; }
  function totalRowNumber(model) { return 2 + model.rows.length; }
  function avgRowNumber(model) { return 3 + model.rows.length; }
  function douRowNumber(model) { return 4 + model.rows.length; } // Typ 行
  function douMinRowNumber(model) { return 5 + model.rows.length; }
  function footerRowNumber(model, rowKey) {
    if (rowKey === 'total') return totalRowNumber(model);
    if (rowKey === 'avg') return avgRowNumber(model);
    if (rowKey === 'douMin') return douMinRowNumber(model);
    return douRowNumber(model);
  }
  function addressOf(model, desc) {
    if (desc.kind === 'current') {
      const rowIndex = model.rows.findIndex((r) => r.id === desc.rowId);
      return rowIndex < 0 ? '' : 'B' + rowNumberFor(rowIndex);
    }
    if (desc.kind === 'note') {
      const colIndex = model.columns.findIndex((c) => c.id === desc.colId);
      const rowIndex = model.rows.findIndex((r) => r.id === desc.rowId);
      return (colIndex < 0 || rowIndex < 0) ? '' : noteLetterFor(colIndex) + rowNumberFor(rowIndex);
    }
    if (desc.kind === 'footerLabel') return 'A' + footerRowNumber(model, desc.rowKey);
    if (desc.kind === 'footerCurrent') return 'B' + footerRowNumber(model, desc.rowKey);
    const colIndex = model.columns.findIndex((c) => c.id === desc.colId);
    if (colIndex < 0) return '';
    if (desc.kind === 'cell') {
      const rowIndex = model.rows.findIndex((r) => r.id === desc.rowId);
      return colLetterFor(colIndex, desc.t) + rowNumberFor(rowIndex);
    }
    const footRowKey = {
      td: 'total', tm: 'total', noteTotal: 'total',
      blankAvg: 'avg', avg: 'avg', noteAvg: 'avg',
      blankTyp: 'douTyp', douTyp: 'douTyp', noteTyp: 'douTyp',
      blankMin: 'douMin', douMin: 'douMin', noteMin: 'douMin',
    }[desc.t];
    if (!footRowKey) return '';
    const row = footerRowNumber(model, footRowKey);
    if (desc.t.indexOf('note') === 0) return noteLetterFor(colIndex) + row;
    // 平均电流 / Typ / Min 三行把「时长 + 耗电」两列合并，A1 地址取合并区左上角（时长列 C/F/I…）；
    // 右侧耗电列地址（D/G/J…）由 refToDesc 解析到同一结果，兼容历史公式。
    return colLetterFor(colIndex, desc.t === 'tm' ? 'm' : 'd') + row;
  }

  // ---------- 旧地址（v2）→ 新地址（v3）迁移 ----------
  // v2 虚拟地址：模型 i 时长 col=2+2i、耗电 col=3+2i（电流/备注无地址）；
  // v3 连续地址：模型 i 时长 col=3+3i、耗电 col=4+3i（电流占 B=2，备注占模块第三列）。
  // 仅迁移模型区（col>=2）；场景 A=col1 不变，行号不变。
  function remapLegacyCol(oldCol) {
    if (oldCol < 2) return oldCol;
    const i = Math.floor((oldCol - 2) / 2);
    const isM = (oldCol - 2) % 2 === 1;
    return 3 + i * 3 + (isM ? 1 : 0);
  }
  // 重写公式字符串中的 A1 引用列字母，保留 $ 绝对标记与行号；非公式（数字/中文名公式）原样返回。
  // token 边界保护：
  //  - 前置负向回顾：列字母前不得是标识符/数字/小数点/中文，避免科学计数法 1E5、名称延续被误判；
  //  - 行号首位禁 0（Excel 无 0 行），避免 V0.1 这类命名被当成引用；
  //  - 后置负向先行：行号后不得紧跟标识符/中文，避免 B2场景 这类命名内片段被迁移。
  function remapFormulaRefs(raw) {
    if (typeof raw !== 'string' || raw.charAt(0) !== '=') return raw;
    return raw.replace(/(?<![A-Za-z0-9_.\u4e00-\u9fff])(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,4})(?![A-Za-z0-9_.\u4e00-\u9fff])/g, (m, d1, letters, d2, rowDigits) => {
      const oldCol = window.ES_FORMULA.letterToCol(letters);
      const newCol = remapLegacyCol(oldCol);
      if (newCol === oldCol) return m;
      return d1 + window.ES_FORMULA.colToLetter(newCol) + d2 + rowDigits;
    });
  }
  // 把 v2 模型内所有裸地址公式迁移到 v3 连续地址，并标记 v=3
  function migrateModelAddressesV2ToV3(model) {
    const walk = (raw) => remapFormulaRefs(raw);
    (model.rows || []).forEach((row) => {
      Object.keys(row.cells || {}).forEach((colId) => {
        const cell = row.cells[colId];
        if (cell) { cell.d = walk(cell.d); cell.m = walk(cell.m); }
      });
    });
    Object.keys(model.foot || {}).forEach((colId) => {
      const f = model.foot[colId];
      if (f) ['td', 'tm', 'avg', 'noteTotal', 'noteAvg', 'noteTyp', 'noteMin', 'blankAvg', 'blankTyp', 'blankMin', 'douTyp', 'douMin'].forEach((k) => { if (typeof f[k] === 'string') f[k] = walk(f[k]); });
    });
    if (model.footer) {
      Object.keys(model.footer.labels || {}).forEach((k) => {
        if (typeof model.footer.labels[k] === 'string') model.footer.labels[k] = walk(model.footer.labels[k]);
      });
      Object.keys(model.footer.current || {}).forEach((k) => {
        if (typeof model.footer.current[k] === 'string') model.footer.current[k] = walk(model.footer.current[k]);
      });
    }
    model.v = 3;
    return model;
  }

  // ---------- 自由编辑（v52）：结构变更时的 A1 引用 Excel 式平移 ----------
  // 与 v2→v3 迁移共用同一套带边界保护的 A1 token 正则（防科学计数 / 名称片段 / 0 行误判）。
  const A1_TOKEN_RE = /(?<![A-Za-z0-9_.一-龥])(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,4})(?![A-Za-z0-9_.一-龥])/g;
  // mapper 接收 {col,row}，返回新 {col,row}；返回 null 表示该引用指向被删除区域，改写为显式错误 (1/0) → #DIV/0!
  function rewriteA1Tokens(raw, mapper) {
    if (typeof raw !== 'string' || raw.charAt(0) !== '=') return raw;
    return raw.replace(A1_TOKEN_RE, (m, d1, letters, d2, rowDigits) => {
      const ref = { col: window.ES_FORMULA.letterToCol(letters), row: parseInt(rowDigits, 10) };
      const out = mapper(ref);
      if (!out) return '(1/0)';
      if (out.col === ref.col && out.row === ref.row) return m;
      return d1 + window.ES_FORMULA.colToLetter(out.col) + d2 + out.row;
    });
  }
  // 遍历模型内所有可能含公式的原始字符串（时长 / 耗电 / 自由电流 / 备注 / 汇总脚注）
  function walkModelFormulaRaws(model, walker) {
    (model.rows || []).forEach((row) => {
      if (typeof row.currentRaw === 'string') row.currentRaw = walker(row.currentRaw);
      if (typeof row.note === 'string') row.note = walker(row.note);
      Object.keys(row.cells || {}).forEach((colId) => {
        const cell = row.cells[colId];
        if (cell) { cell.d = walker(cell.d); cell.m = walker(cell.m); }
      });
    });
    Object.keys(model.foot || {}).forEach((colId) => {
      const f = model.foot[colId];
      if (f) ['td', 'tm', 'avg', 'noteTotal', 'noteAvg', 'noteTyp', 'noteMin', 'blankAvg', 'blankTyp', 'blankMin', 'douTyp', 'douMin'].forEach((k) => { if (typeof f[k] === 'string') f[k] = walker(f[k]); });
    });
    if (model.footer) {
      Object.values(model.footer.labels || {}).forEach((v, i) => {
        if (typeof v === 'string') model.footer.labels[Object.keys(model.footer.labels)[i]] = walker(v);
      });
      Object.values(model.footer.current || {}).forEach((v, i) => {
        if (typeof v === 'string') model.footer.current[Object.keys(model.footer.current)[i]] = walker(v);
      });
    }
  }
  // 删除第 deletedDataIndex 个数据行（0-based）后：
  //   指向被删行 → (1/0)；指向其下方的行号整体 -1；上方不动。列字母保持不变。
  function rewriteRefsAfterRowDelete(model, deletedDataIndex) {
    const deletedExcelRow = deletedDataIndex + 2; // 第 1 行为表头，数据行从第 2 行起
    walkModelFormulaRaws(model, (raw) => rewriteA1Tokens(raw, (ref) => {
      if (ref.row === deletedExcelRow) return null;
      if (ref.row > deletedExcelRow) return { col: ref.col, row: ref.row - 1 };
      return ref;
    }));
  }
  // 在第 insertedDataIndex 个数据行（0-based）前插入一行后：
  //   指向插入位置及其下方（含 Total / DOU 虚拟行）的行号整体 +1；列字母保持不变。
  function rewriteRefsAfterRowInsert(model, insertedDataIndex) {
    const insertedExcelRow = insertedDataIndex + 2; // 第 1 行为表头，数据行从第 2 行起
    walkModelFormulaRaws(model, (raw) => rewriteA1Tokens(raw, (ref) => {
      if (ref.row >= insertedExcelRow) return { col: ref.col, row: ref.row + 1 };
      return ref;
    }));
  }
  // v52 → v56 汇总区地址迁移：
  // 旧：Total=n+2，DOU 虚拟行=n+3（C=Typ、D=Min、E 不可用）；新：Total=n+2、平均=n+3、Typ=n+4、Min=n+5。
  // 纯 Excel 单元格形态下 D 列是独立空格，因此旧 Min 的 D 地址迁移到新 Min 的值列（C/F/I…）。
  function migrateFooterAddressesV52ToV53(model) {
    const oldDouRow = model.rows.length + 3;
    walkModelFormulaRaws(model, (raw) => rewriteA1Tokens(raw, (ref) => {
      if (ref.col === 1 || ref.col === 2) return ref.row === oldDouRow ? null : ref; // 旧 A/B 汇总格无有效值
      if (ref.row < oldDouRow) return ref;
      if (ref.row > oldDouRow) return { col: ref.col, row: ref.row + 2 };
      const offset = (ref.col - 3) % 3; // 模块列：C/F/I=0，D/G/J=1，E/H/K=2
      if (offset === 0) return { col: ref.col, row: oldDouRow + 1 };     // 旧 Typ C15 → 新 Typ 值列 C16
      if (offset === 1) return { col: ref.col - 1, row: oldDouRow + 2 }; // 旧 Min D15 → 新 Min 值列 C17
      return null;                                                       // 旧 E15 原本不可引用
    }));
  }
  // 删除第 moduleIndex 个对比列模块（每个模块占 3 个物理列：时长/耗电/备注）后：
  //   指向被删模块三列 → (1/0)；指向其右侧模块的列整体 -3；左侧不动。行号保持不变。
  function rewriteRefsAfterModuleDelete(model, moduleIndex) {
    const base = 3 + moduleIndex * 3; // 该模块时长列的物理列号（1-based）
    walkModelFormulaRaws(model, (raw) => rewriteA1Tokens(raw, (ref) => {
      if (ref.col < 3) return ref; // A 场景名 / B 电流列不参与模块位移
      if (ref.col >= base && ref.col <= base + 2) return null;
      if (ref.col >= base + 3) return { col: ref.col - 3, row: ref.row };
      return ref;
    }));
  }
  // 新建一个与场景解耦的「空白自由行」：A 标题、B 电流、各模块时长/耗电/备注全部留空，参与地址与 SUM（空值不计）
  function createBlankRow(model) {
    const cells = {};
    (model.columns || []).forEach((col) => { cells[col.id] = { d: '', m: '' }; });
    return { id: uid('mr'), scenarioId: null, label: '', note: '', currentRaw: '', cells: cells };
  }
  // 新增场景行：汇总区会整体下推一行，先平移 A1 引用，再追加默认时长/耗电公式。
  function appendRowForScenario(model, scenarioId) {
    if (model.rows.some((r) => r.scenarioId === scenarioId)) return;
    rewriteRefsAfterRowInsert(model, model.rows.length);
    const cells = {};
    model.columns.forEach((col) => { cells[col.id] = { d: '0', m: DEFAULT_MAH_FORMULA }; });
    model.rows.push({ id: uid('mr'), scenarioId: scenarioId, label: '', currentRaw: '', note: '', cells: cells });
  }

  // 电池参数命名表：供 DOU 网格与单场景续航公式共用，随电池卡片实时联动、不硬编码容量。
  // 缺容量/关机电压等返回 #VALUE! 占位对象，比例空值按 1、门限空值按 0。
  function batteryNameValues(project) {
    const ET = window.ES_FORMULA.ERROR_TYPES;
    const b = (project && project.battery) || {};
    const capTyp = num(b.capTyp);
    const capMin = num(b.capMin);
    const cfg = projectConfig(project);
    const ratioRaw = num(b.dischargeRatio);
    const thresholdRaw = num(cfg.lowThreshold);
    const ratio = ratioRaw === null ? 1 : ratioRaw;
    const threshold = thresholdRaw === null ? 0 : thresholdRaw;
    const need = (v) => (v === null ? { type: ET.VALUE } : v);
    const available = capTyp === null ? { type: ET.VALUE } : capTyp * ratio;
    const derated = capTyp === null ? { type: ET.VALUE } : capTyp * ratio * (1 - threshold);
    return {
      'Typ': need(capTyp), 'Typ容量': need(capTyp), '电池容量': need(capTyp),
      'Min': need(capMin), 'Min容量': need(capMin),
      '比例': ratio, '可放电比例': ratio, '可放电容量比例': ratio,
      '门限': threshold, '低电量门限': threshold,
      '可用容量': available, '可用电量': available,
      '折减容量': derated,
      '关机电压': need(num(cfg.shutdownVoltage)),
      'Wh': need(num(b.wh)), '电压': need(num(b.voltage)),
    };
  }

  // 单场景续航求值：每个场景一条独立公式（scenario.singleLife），无 A1 网格地址，
  // 仅支持电池参数命名 + 本场景电流命名；空公式返回 empty，纯数字直接取值，出错返回错误码。
  function evaluateSingleLife(project, scenario) {
    const F = window.ES_FORMULA;
    const ET = F.ERROR_TYPES;
    const raw = scenario && scenario.singleLife;
    if (raw === null || raw === undefined) return { ok: true, empty: true, value: null };
    const s = String(raw).trim();
    if (s === '') return { ok: true, empty: true, value: null };
    const names = batteryNameValues(project);
    const current = scenarioCurrent(project, scenario.id);
    const ctx = {
      getRef() { return { type: ET.REF }; },
      getRange() { return { type: ET.REF }; },
      getName(name) {
        if (Object.prototype.hasOwnProperty.call(names, name)) return names[name];
        if (name === '电流' || name === '场景电流' || name === '当前电流' || name === '本行电流') {
          return current === null ? { type: ET.REF } : current;
        }
        return { type: ET.NAME };
      },
      cellKey(ref) { return 'single:' + (ref && ref.label ? ref.label : ''); },
    };
    if (s.charAt(0) === '=') {
      const r = F.evaluate(s.slice(1), ctx);
      return r.ok ? { ok: true, value: r.value } : { ok: false, error: r.error };
    }
    const v = num(s);
    if (v === null) return { ok: false, error: ET.VALUE };
    return { ok: true, value: v };
  }

  // ---------- 公式网格求值 ----------
  // 返回 { result(desc), stats, errorOf(desc) }；一次求值内带 memo 与循环检测
  function evaluateModel(project, model) {
    const F = window.ES_FORMULA;
    const ET = F.ERROR_TYPES;
    const cache = new Map();
    const n = model.rows.length;
    const totalRow = totalRowNumber(model);
    const avgRow = avgRowNumber(model);
    const douRow = douRowNumber(model);
    const minRow = douMinRowNumber(model);

    function keyOf(desc) {
      return desc.kind + '|' + (desc.rowId || desc.rowKey || '') + '|' + desc.colId + '|' + (desc.t || '');
    }

    // A1 地址 -> 单元格描述（v3 连续地址；v52 起 B 电流、备注列也可被引用）
    // col1=A 场景名（文本标签，不参与公式），col2=B 电流（自由：数字/文字/公式），
    // 模型模块 i 占 3+3i 起三列：时长 / 耗电 / 备注（备注可引用，文本在数值运算中记 #VALUE!）。
    function refToDesc(ref) {
      if (ref.col === 1) {
        const footerKey = { [totalRow]: 'total', [avgRow]: 'avg', [douRow]: 'douTyp', [minRow]: 'douMin' }[ref.row];
        return footerKey ? { kind: 'footerLabel', rowKey: footerKey } : null;
      }
      if (ref.col === 2) {
        // B 电流列：数据行为场景/自由行电流；四个汇总行也各自可编辑
        if (ref.row >= 2 && ref.row <= 1 + n) return { kind: 'current', rowId: model.rows[ref.row - 2].id };
        const footerKey = { [totalRow]: 'total', [avgRow]: 'avg', [douRow]: 'douTyp', [minRow]: 'douMin' }[ref.row];
        return footerKey ? { kind: 'footerCurrent', rowKey: footerKey } : null;
      }
      const offset = (ref.col - 3) % 3;
      const ci = Math.floor((ref.col - 3) / 3);
      if (ci >= model.columns.length) return null;
      const colId = model.columns[ci].id;
      if (ref.row >= 2 && ref.row <= 1 + n) {
        const row = model.rows[ref.row - 2];
        if (offset === 2) return { kind: 'note', rowId: row.id, colId: colId }; // 备注列（可引用）
        return { kind: 'cell', rowId: row.id, colId: colId, t: offset === 1 ? 'm' : 'd' };
      }
      const footerTypes = {
        [totalRow]: ['td', 'tm', 'noteTotal'],
        // 平均 / Typ / Min 为「时长+耗电」合并区：左上角 C 与右侧 D 都解析到同一个结果
        [avgRow]: ['avg', 'avg', 'noteAvg'],
        [douRow]: ['douTyp', 'douTyp', 'noteTyp'],
        [minRow]: ['douMin', 'douMin', 'noteMin'],
      };
      return footerTypes[ref.row] ? { kind: 'foot', colId: colId, t: footerTypes[ref.row][offset] } : null;
    }

    function evalCell(desc, visiting) {
      const key = keyOf(desc);
      if (cache.has(key)) return cache.get(key);
      if (visiting[key]) return { ok: false, error: ET.CYCLE };
      visiting[key] = true;

      let raw = '';
      let scopeRow = null;
      if (desc.kind === 'current' || desc.kind === 'note') {
        scopeRow = model.rows.find((r) => r.id === desc.rowId) || null;
        raw = scopeRow ? (desc.kind === 'current' ? (scopeRow.currentRaw || '') : (scopeRow.note || '')) : '';
      } else if (desc.kind === 'footerLabel') {
        raw = (model.footer && model.footer.labels && model.footer.labels[desc.rowKey]) || '';
      } else if (desc.kind === 'footerCurrent') {
        raw = (model.footer && model.footer.current && model.footer.current[desc.rowKey]) || '';
      } else if (desc.kind === 'cell') {
        const row = model.rows.find((r) => r.id === desc.rowId);
        const cell = row && row.cells[desc.colId];
        raw = cell ? (cell[desc.t] === undefined || cell[desc.t] === null ? '' : cell[desc.t]) : '';
        scopeRow = row;
      } else {
        raw = (model.foot[desc.colId] || {})[desc.t] || '';
      }

      let result;
      if (desc.kind === 'foot' && desc.t === 'avg' && String(raw).trim() === '') {
        // 平均电流留空时，UI 显示自动值；A1 引用也必须得到同一个自动值，避免“看得见但公式引用不到”。
        const tdResult = evalCell({ kind: 'foot', colId: desc.colId, t: 'td' }, visiting);
        const tmResult = evalCell({ kind: 'foot', colId: desc.colId, t: 'tm' }, visiting);
        const canAuto = tdResult.ok && !tdResult.empty && tdResult.value !== 0
          && tmResult.ok && !tmResult.empty;
        result = canAuto ? { ok: true, value: tmResult.value / tdResult.value } : { ok: true, empty: true, value: null };
      } else if (desc.kind === 'current' && String(raw).trim() === '') {
        // B 电流留空：回退到场景共享电流（自定义电流优先，否则最新测试记录），与「电流」命名口径一致
        const fallback = scopeRow ? scenarioCurrent(project, scopeRow.scenarioId) : null;
        result = fallback === null ? { ok: true, empty: true, value: null } : { ok: true, value: fallback };
      } else {
        const ctx = makeCtx(desc, scopeRow, visiting);
        result = evalRaw(raw, ctx);
      }
      delete visiting[key];
      cache.set(key, result);
      return result;
    }

    function evalRaw(raw, ctx) {
      if (raw === null || raw === undefined) return { ok: true, empty: true, value: null };
      const s = String(raw).trim();
      if (s === '') return { ok: true, empty: true, value: null };
      if (s.charAt(0) === '=') {
        const r = F.evaluate(s.slice(1), ctx);
        return r.ok ? { ok: true, value: r.value } : { ok: false, error: r.error };
      }
      const v = num(s);
      if (v === null) return { ok: false, error: ET.VALUE };
      return { ok: true, value: v };
    }

    function makeCtx(desc, scopeRow, visiting) {
      const names = batteryNameValues(project);

      function rowResult(row, cellType) {
        return evalCell({ kind: 'cell', rowId: row.id, colId: desc.colId, t: cellType }, visiting);
      }
      function footResult(t) {
        return evalCell({ kind: 'foot', colId: desc.colId, t: t }, visiting);
      }
      function valueOrError(r) {
        if (r.ok) return r.empty ? null : r.value;
        return { type: r.error };
      }

      return {
        getRef(ref) {
          const target = refToDesc(ref);
          if (!target) return { type: ET.REF };
          const r = evalCell(target, visiting);
          return r.ok ? (r.empty ? { type: ET.REF } : r.value) : { type: r.error };
        },
        getRange(a, b) {
          const refs = [];
          for (let rr = a.row; rr <= b.row; rr += 1) {
            for (let cc = a.col; cc <= b.col; cc += 1) {
              // 合并区：单点引用右侧仍取同一个值，但区间聚合时右侧按 Excel 合并语义视为空，避免重复计数
              const mergedRight = (rr === avgRow || rr === douRow || rr === minRow)
                && cc >= 3 && (cc - 3) % 3 === 1;
              if (mergedRight) { refs.push(null); continue; }
              const target = refToDesc({ col: cc, row: rr, label: '' });
              if (!target) { refs.push({ type: ET.REF }); continue; }
              const r = evalCell(target, visiting);
              refs.push(r.ok ? (r.empty ? null : r.value) : { type: r.error });
            }
          }
          return { __range: true, refs: refs };
        },
        getName(name) {
          if (Object.prototype.hasOwnProperty.call(names, name)) return names[name];
          // 结构化名称
          if (name === '时长' || name === '本行时长') {
            if (desc.kind !== 'cell' || !scopeRow) return { type: ET.REF };
            return valueOrError(rowResult(scopeRow, 'd'));
          }
          if (name === '电流' || name === '本行电流' || name === '场景电流') {
            if (!scopeRow) return { type: ET.REF };
            // 在 B 电流格自身的公式里，「电流」取场景共享电流（避免自引用循环，便于写 =电流*1.1 这类公式）
            if (desc.kind === 'current') {
              const base = scenarioCurrent(project, scopeRow.scenarioId);
              return base === null ? { type: ET.REF } : base;
            }
            // 其它格引用「电流」：B 列有自由覆盖（数字/文字/公式）时取覆盖值，否则回退场景共享电流
            if (String(scopeRow.currentRaw || '').trim() !== '') {
              return valueOrError(evalCell({ kind: 'current', rowId: scopeRow.id }, visiting));
            }
            const current = scenarioCurrent(project, scopeRow.scenarioId);
            return current === null ? { type: ET.REF } : current;
          }
          if (name === '本列时长' || name === '本列耗电') {
            const cellType = name === '本列时长' ? 'd' : 'm';
            const refs = model.rows.map((row) => {
              const r = rowResult(row, cellType);
              return r.ok ? (r.empty ? null : r.value) : { type: r.error };
            });
            return { __range: true, refs: refs };
          }
          if (name === '总时长') return valueOrError(footResult('td'));
          if (name === '总耗电') return valueOrError(footResult('tm'));
          return { type: ET.NAME };
        },
        cellKey(ref) {
          const target = refToDesc(ref);
          return target ? keyOf(target) : ('ref?' + ref.label);
        },
      };
    }

    // 预热全表
    const stats = {};
    model.columns.forEach((col) => {
      const td = evalCell({ kind: 'foot', colId: col.id, t: 'td' }, {});
      const tm = evalCell({ kind: 'foot', colId: col.id, t: 'tm' }, {});
      const douTyp = evalCell({ kind: 'foot', colId: col.id, t: 'douTyp' }, {});
      const douMin = evalCell({ kind: 'foot', colId: col.id, t: 'douMin' }, {});
      // 平均电流留空时也预热：结果为自动总耗电 / 总时长，保证 A1 引用与界面显示一致。
      evalCell({ kind: 'foot', colId: col.id, t: 'avg' }, {});
      ['noteTotal', 'noteAvg', 'noteTyp', 'noteMin', 'blankAvg', 'blankTyp', 'blankMin'].forEach((t) => {
        evalCell({ kind: 'foot', colId: col.id, t: t }, {});
      });
      const totalHours = td.ok ? td.value : null;
      const totalMah = tm.ok ? tm.value : null;
      stats[col.id] = {
        totalHours: totalHours,
        totalMah: totalMah,
        avgCurrent: (totalHours && totalMah !== null) ? totalMah / totalHours : null,
        douTyp: douTyp.ok ? douTyp.value : null,
        douMin: douMin.ok ? douMin.value : null,
        douTypError: douTyp.ok ? null : douTyp.error,
        douMinError: douMin.ok ? null : douMin.error,
        tdError: td.ok ? null : td.error,
        tmError: tm.ok ? null : tm.error,
      };
    });
    ['total', 'avg', 'douTyp', 'douMin'].forEach((rowKey) => {
      evalCell({ kind: 'footerLabel', rowKey }, {});
      evalCell({ kind: 'footerCurrent', rowKey }, {});
    });
    model.rows.forEach((row) => {
      evalCell({ kind: 'current', rowId: row.id }, {}); // 预热 B 电流（含留空回退），供网格直接 result()
      model.columns.forEach((col) => {
        evalCell({ kind: 'cell', rowId: row.id, colId: col.id, t: 'd' }, {});
        evalCell({ kind: 'cell', rowId: row.id, colId: col.id, t: 'm' }, {});
        evalCell({ kind: 'note', rowId: row.id, colId: col.id }, {}); // 预热备注，使其可被 result()/A1 引用
      });
    });

    return {
      result(desc) { return cache.get(keyOf(desc)) || { ok: true, empty: true, value: null }; },
      stats: stats,
    };
  }

  // 兼容旧接口：单个对比列（使用模式）的统计
  function computeModeStats(project, model, column) {
    const ev = evaluateModel(project, model);
    const s = ev.stats[column.id] || {};
    const b = project.battery || {};
    const cfg = projectConfig(project);
    const ratio = num(b.dischargeRatio);
    const threshold = num(cfg.lowThreshold);
    const derate = (ratio === null ? 1 : ratio) * (1 - (threshold === null ? 0 : threshold));
    const isFull = column.kind === 'full';
    return {
      totalHours: s.totalHours, totalMah: s.totalMah, avgCurrent: s.avgCurrent,
      fullHours: isFull ? s.douTyp : null,
      douTyp: isFull ? null : s.douTyp,
      douMin: isFull ? null : s.douMin,
      derate: derate,
      errors: { td: s.tdError, tm: s.tmError, douTyp: s.douTypError, douMin: s.douMinError },
    };
  }

  function computeModelStats(project, model) {
    const out = {};
    model.columns.forEach((col) => { out[col.id] = computeModeStats(project, model, col); });
    return out;
  }

  window.ES_ENDURANCE = {
    STORAGE_KEY, CHART_COLORS, chartInstances, destroyCharts,
    uid, escapeHtml, num, fmt, todayStr, debounce,
    store, defaultBattery, latestRecord, scenarioCurrent,
    computeModeStats, computeModelStats, evaluateModel, evaluateSingleLife,
    sortDischargeTable, lookupDischargeCapacity, deriveDischargeRatio,
    DEFAULT_CUTOFF_VOLTAGES, matrixVoltages, matrixBaseVoltage, matrixCells, matrixCellValue,
    matrixCurve, matrixFromLegacyTable, matrixSetCell, dischargeCurve, sampleDischargeMatrix,
    defaultProjectConfig, projectConfig,
    DEFAULT_MAH_FORMULA, defaultFoot, defaultModelFooter, createColumn, createModelV2, migrateModelV1,
    defaultScenarioDefs, buildDefaultScenarios, mergeDefaultScenarios, currentCompare,
    sanitizeFileName, fileTimestamp, buildProjectExport, parseProjectImport,
    isValidProjectShape, uniqueProjectName, reidentifyProject, downloadTextFile,
    GRID_LAYOUT_KEY, GRID_COL_ROLES, GRID_ROW_H, GRID_DEFAULT_ROWH,
    defaultGridLayout, loadGridLayout, saveGridLayout, resetGridLayout, clampGridColWidth,
    colLetterFor, noteLetterFor, rowNumberFor, totalRowNumber, avgRowNumber, douRowNumber, douMinRowNumber,
    footerRowNumber, addressOf,
    remapLegacyCol, remapFormulaRefs, migrateModelAddressesV2ToV3, migrateFooterAddressesV52ToV53, normalizeData,
    rewriteA1Tokens, rewriteRefsAfterRowDelete, rewriteRefsAfterRowInsert,
    rewriteRefsAfterModuleDelete, createBlankRow, appendRowForScenario,
  };
})();

/* ================================================================
 * 公共模态框 ES_MODAL：屏幕居中的美化弹窗，替代浏览器原生 prompt/confirm/alert。
 * 全部 Promise 化：prompt 返回 string|null、confirm 返回 boolean、
 * alert 关闭后 resolve、form 返回字段对象|null（取消）。
 * ================================================================ */
(function () {
  'use strict';
  const E = window.ES_ENDURANCE;
  const { escapeHtml } = E;
  let activeOverlay = null;

  function closeActive() {
    if (!activeOverlay) return;
    const overlay = activeOverlay;
    const previous = overlay._previousFocus || null;
    if (typeof overlay._reject === 'function') overlay._reject();
    if (overlay.parentNode) overlay.parentNode.removeChild(overlay);
    activeOverlay = null;
    document.removeEventListener('keydown', onKeydown, true);
    document.body.classList.remove('es-modal-open');
    if (previous && typeof previous.focus === 'function') {
      try { previous.focus({ preventScroll: true }); } catch (err) { previous.focus(); }
    }
  }
  function onKeydown(event) {
    if (!activeOverlay) return;
    if (event.key === 'Escape') {
      event.stopPropagation();
      activeOverlay._reject(null);
      return;
    }
    // 简单焦点陷阱：Tab 在弹窗内循环
    if (event.key === 'Tab') {
      const focusables = activeOverlay.querySelectorAll('input, select, textarea, button:not([disabled])');
      if (!focusables.length) return;
      const first = focusables[0];
      const last = focusables[focusables.length - 1];
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    }
  }

  function openModal({ title, icon, bodyHtml, confirmText, cancelText, danger, onMount, onSubmit, hideCancel }) {
    return new Promise((resolve) => {
      // 先关闭并 resolve 旧弹窗，避免旧 Promise 永久挂起
      closeActive();
      const previousFocus = document.activeElement;
      const overlay = document.createElement('div');
      overlay.className = 'es-modal-overlay';
      overlay.innerHTML = '<div class="es-modal-dialog panel rounded-3xl" role="dialog" aria-modal="true">'
        + '<div class="es-modal-head"><h4><span class="material-symbols-outlined">' + escapeHtml(icon || 'edit_note') + '</span>'
        + escapeHtml(title || '') + '</h4>'
        + '<button type="button" class="es-icon-btn es-modal-close" title="关闭（Esc）"><span class="material-symbols-outlined">close</span></button></div>'
        + '<div class="es-modal-body">' + bodyHtml + '</div>'
        + '<p class="es-modal-error" hidden></p>'
        + '<div class="es-modal-foot">'
        + (hideCancel ? '' : '<button type="button" class="template-btn es-modal-cancel">' + escapeHtml(cancelText || '取消') + '</button>')
        + '<button type="button" class="template-btn es-modal-btn ' + (danger ? 'es-modal-confirm-danger' : 'es-modal-confirm-primary') + ' es-modal-confirm">'
        + escapeHtml(confirmText || '确定') + '</button>'
        + '</div></div>';
      let settled = false;
      overlay._previousFocus = previousFocus;
      overlay._reject = () => {
        if (settled) return;
        settled = true;
        if (activeOverlay === overlay) {
          activeOverlay = null;
          document.removeEventListener('keydown', onKeydown, true);
          document.body.classList.remove('es-modal-open');
          if (overlay.parentNode) overlay.parentNode.removeChild(overlay);
          if (previousFocus && typeof previousFocus.focus === 'function') {
            try { previousFocus.focus({ preventScroll: true }); } catch (err) { previousFocus.focus(); }
          }
        }
        resolve(null);
      };

      const showError = (msg) => {
        const errEl = overlay.querySelector('.es-modal-error');
        if (!msg) { errEl.hidden = true; errEl.textContent = ''; return; }
        errEl.hidden = false;
        errEl.textContent = msg;
      };
      const done = (value) => {
        if (settled) return;
        settled = true;
        if (activeOverlay === overlay) {
          activeOverlay = null;
          document.removeEventListener('keydown', onKeydown, true);
          document.body.classList.remove('es-modal-open');
          if (overlay.parentNode) overlay.parentNode.removeChild(overlay);
          if (previousFocus && typeof previousFocus.focus === 'function') {
            try { previousFocus.focus({ preventScroll: true }); } catch (err) { previousFocus.focus(); }
          }
        }
        resolve(value);
      };

      overlay.querySelector('.es-modal-close').addEventListener('click', () => overlay._reject());
      const cancelBtn = overlay.querySelector('.es-modal-cancel');
      if (cancelBtn) cancelBtn.addEventListener('click', () => overlay._reject());
      overlay.addEventListener('mousedown', (event) => { if (event.target === overlay) overlay._reject(); });
      document.addEventListener('keydown', onKeydown, true);

      activeOverlay = overlay;
      document.body.classList.add('es-modal-open');
      document.body.appendChild(overlay);
      const ctx = { overlay, showError, done, close: () => overlay._reject() };
      if (onMount) onMount(ctx);
      overlay.querySelector('.es-modal-confirm').addEventListener('click', () => onSubmit(ctx));
      const firstInput = overlay.querySelector('input, select, textarea');
      if (firstInput) {
        firstInput.focus();
        if (typeof firstInput.select === 'function' && firstInput.type !== 'date') firstInput.select();
      } else {
        overlay.querySelector('.es-modal-confirm').focus();
      }
    });
  }

  function fieldHtml(f) {
    const val = f.value === null || f.value === undefined ? '' : f.value;
    const common = ' class="es-modal-input" autocomplete="off"'
      + (f.placeholder ? ' placeholder="' + escapeHtml(f.placeholder) + '"' : '')
      + (f.inputmode ? ' inputmode="' + escapeHtml(f.inputmode) + '"' : '');
    let control;
    if (f.type === 'select') {
      control = '<select' + common + '>' + f.options.map((o) => {
        const ov = typeof o === 'object' ? o.value : o;
        const ol = typeof o === 'object' ? o.label : o;
        return '<option value="' + escapeHtml(ov) + '"' + (String(ov) === String(val) ? ' selected' : '') + '>' + escapeHtml(ol) + '</option>';
      }).join('') + '</select>';
    } else if (f.type === 'textarea') {
      control = '<textarea rows="3"' + common + '>' + escapeHtml(val) + '</textarea>';
    } else {
      control = '<input type="' + escapeHtml(f.type || 'text') + '"' + common + ' value="' + escapeHtml(val) + '" />';
    }
    return '<label class="es-modal-field"><span>' + escapeHtml(f.label) + '</span>' + control + '</label>';
  }

  const ES_MODAL = {
    prompt(opts) {
      const f = {
        key: 'value', label: opts.label || '名称', value: opts.value || '',
        placeholder: opts.placeholder || '', type: opts.type || 'text', inputmode: opts.inputmode,
      };
      return openModal({
        title: opts.title, icon: opts.icon,
        confirmText: opts.confirmText, cancelText: opts.cancelText, danger: opts.danger,
        bodyHtml: fieldHtml(f),
        onMount: (ctx) => {
          const input = ctx.overlay.querySelector('.es-modal-input');
          input.addEventListener('keydown', (event) => { if (event.key === 'Enter') { event.preventDefault(); ctx.overlay.querySelector('.es-modal-confirm').click(); } });
        },
        onSubmit: (ctx) => {
          const raw = ctx.overlay.querySelector('.es-modal-input').value;
          if (opts.validate) {
            const msg = opts.validate(raw);
            if (msg) { ctx.showError(msg); return; }
          }
          ctx.done(raw);
        },
      });
    },
    confirm(opts) {
      return openModal({
        title: opts.title || '请确认', icon: opts.icon || 'help', danger: true,
        confirmText: opts.confirmText || '确认删除', cancelText: opts.cancelText || '取消',
        bodyHtml: '<p class="es-modal-message">' + escapeHtml(opts.message || '').replace(/\n/g, '<br>') + '</p>',
        onSubmit: (ctx) => ctx.done(true),
      }).then((v) => v === true);
    },
    alert(opts) {
      const body = opts.html
        ? '<div class="es-modal-message es-modal-html">' + opts.html + '</div>'
        : '<p class="es-modal-message">' + escapeHtml(opts.message || '').replace(/\n/g, '<br>') + '</p>';
      return openModal({
        title: opts.title || '提示', icon: opts.icon || 'info', hideCancel: true,
        confirmText: opts.confirmText || '我知道了',
        bodyHtml: body,
        onSubmit: (ctx) => ctx.done(true),
      }).then(() => undefined);
    },
    form(opts) {
      const fields = opts.fields || [];
      return openModal({
        title: opts.title, icon: opts.icon,
        confirmText: opts.confirmText, cancelText: opts.cancelText, danger: opts.danger,
        bodyHtml: fields.map(fieldHtml).join(''),
        onMount: (ctx) => {
          ctx.overlay.querySelectorAll('.es-modal-input').forEach((input) => {
            input.addEventListener('keydown', (event) => {
              if (event.key === 'Enter' && input.tagName !== 'TEXTAREA') {
                event.preventDefault();
                ctx.overlay.querySelector('.es-modal-confirm').click();
              }
            });
          });
          if (opts.onMount) opts.onMount(ctx);
        },
        onSubmit: (ctx) => {
          const values = {};
          fields.forEach((f, i) => {
            values[f.key] = ctx.overlay.querySelectorAll('.es-modal-input')[i].value;
          });
          if (opts.validate) {
            const msg = opts.validate(values);
            if (msg) { ctx.showError(msg); return; }
          }
          ctx.done(values);
        },
      });
    },
  };
  window.ES_MODAL = ES_MODAL;
})();

/* ================================================================
 * 第二部分：页面骨架、项目管理、电池参数、可放电容量查询表
 * ================================================================ */
(function () {
  'use strict';
  const E = window.ES_ENDURANCE;
  const {
    escapeHtml, uid, num, fmt, debounce,
    projectConfig, defaultProjectConfig, matrixVoltages, matrixCells, matrixCellValue, matrixSetCell,
    dischargeCurve, DEFAULT_CUTOFF_VOLTAGES,
  } = E;

  const root = document.getElementById('enduranceStatusRoot');

  // ---------- 顶层渲染 ----------
  function render() {
    if (!root) return;
    E.destroyCharts();
    const data = E.store.load();
    const project = E.store.activeProject();
    root.innerHTML = ''
      + renderProjectBar(data, project)
      + (project ? renderBoard(project) : renderEmpty());
    bindProjectBar(data, project);
    if (project) {
      bindBattery(project);
      bindDischargeCard(project);
      bindProjectConfig(project);
      window.ES_ENDURANCE_SCENARIO.bindScenarios(project);
      window.ES_ENDURANCE_MODEL.bindModels(project);
      window.ES_ENDURANCE_UI.bindCompare(project);
    }
  }

  function renderEmpty() {
    return '<div class="panel rounded-3xl es-empty">'
      + '<span class="material-symbols-outlined es-empty-icon">battery_saver</span>'
      + '<h3>还没有项目</h3>'
      + '<p>新增一个项目后即可维护电池参数、场景电流与 DOU 续航模型，并进行多项目横向对比。</p>'
      + '<button type="button" class="template-btn template-btn-primary es-add-project-btn">'
      + '<span class="material-symbols-outlined text-sm">add</span>新增第一个项目</button>'
      + '</div>';
  }

  // ---------- 项目管理栏 ----------
  function renderProjectBar(data, project) {
    const chips = data.projects.map((p) => {
      const active = project && p.id === project.id;
      return '<button type="button" class="es-project-chip' + (active ? ' is-active' : '') + '" data-es-project="' + escapeHtml(p.id) + '">'
        + '<span class="material-symbols-outlined">folder' + (active ? '_special' : '') + '</span>'
        + '<span>' + escapeHtml(p.name) + '</span>'
        + '</button>';
    }).join('');
    return '<div class="panel rounded-3xl es-project-bar">'
      + '<div class="es-project-chips" role="group" aria-label="项目切换">' + chips + '</div>'
      + '<div class="es-project-actions">'
      + '<button type="button" class="template-btn template-btn-primary es-add-project-btn"><span class="material-symbols-outlined text-sm">add</span>新增项目</button>'
      + '<button type="button" class="template-btn es-import-project-btn"><span class="material-symbols-outlined text-sm">upload_file</span>导入项目</button>'
      + (project ? '<button type="button" class="template-btn es-export-project-btn"><span class="material-symbols-outlined text-sm">download</span>导出项目</button>'
        + '<button type="button" class="template-btn es-rename-project-btn"><span class="material-symbols-outlined text-sm">drive_file_rename_outline</span>重命名</button>'
        + '<button type="button" class="template-btn template-btn-danger es-delete-project-btn"><span class="material-symbols-outlined text-sm">delete</span>删除项目</button>' : '')
      + '<input type="file" class="es-import-file-input" accept=".json,application/json" hidden />'
      + '</div></div>';
  }

  function bindProjectBar(data, project) {
    root.querySelectorAll('[data-es-project]').forEach((btn) => {
      btn.addEventListener('click', () => {
        E.store.setActive(btn.dataset.esProject);
        render();
      });
    });
    root.querySelectorAll('.es-add-project-btn').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const name = await window.ES_MODAL.prompt({
          title: '新建项目', icon: 'create_new_folder', label: '项目名称',
          value: '新项目 ' + (data.projects.length + 1),
          validate: (v) => (v.trim() ? null : '项目名称不能为空'),
        });
        if (name === null) return;
        E.store.addProject(name.trim());
        render();
      });
    });
    // 导出当前项目为 .json：纯本地下载，不经过任何云端
    const exportBtn = root.querySelector('.es-export-project-btn');
    if (exportBtn && project) {
      exportBtn.addEventListener('click', () => {
        try {
          const payload = E.buildProjectExport(project);
          const fileName = '续航项目_' + E.sanitizeFileName(project.name) + '_' + E.fileTimestamp() + '.json';
          E.downloadTextFile(fileName, JSON.stringify(payload, null, 2));
        } catch (error) {
          window.ES_MODAL.alert({
            title: '导出失败', icon: 'error',
            message: '生成项目文件失败：' + (error && error.message ? error.message : '未知错误'),
          });
        }
      });
    }
    // 导入 .json：隐藏 file input 选文件 → FileReader 读取 → 校验 → 作为独立副本写入
    const importBtn = root.querySelector('.es-import-project-btn');
    const fileInput = root.querySelector('.es-import-file-input');
    if (importBtn && fileInput) {
      importBtn.addEventListener('click', () => fileInput.click());
      fileInput.addEventListener('change', () => {
        const file = fileInput.files && fileInput.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = () => handleImportText(String(reader.result || ''));
        reader.onerror = () => window.ES_MODAL.alert({ title: '导入失败', icon: 'error', message: '读取文件失败，请重试。' });
        reader.readAsText(file);
        fileInput.value = ''; // 复位，允许再次选择同一个文件
      });
    }
    const renameBtn = root.querySelector('.es-rename-project-btn');
    if (renameBtn && project) {
      renameBtn.addEventListener('click', async () => {
        const name = await window.ES_MODAL.prompt({
          title: '重命名项目', icon: 'drive_file_rename_outline', label: '项目名称',
          value: project.name,
          validate: (v) => (v.trim() ? null : '项目名称不能为空'),
        });
        if (name === null) return;
        project.name = name.trim();
        E.store.save();
        render();
      });
    }
    const deleteBtn = root.querySelector('.es-delete-project-btn');
    if (deleteBtn && project) {
      deleteBtn.addEventListener('click', async () => {
        const ok = await window.ES_MODAL.confirm({
          title: '删除项目', icon: 'delete',
          message: '确认删除项目「' + project.name + '」？\n该项目下的电池参数、场景电流与全部续航模型都会被删除，且不可恢复。',
        });
        if (!ok) return;
        E.store.removeProject(project.id);
        render();
      });
    }
  }

  // 解析导入文本并落库；任何非法输入都通过弹窗提示，绝不破坏现有数据
  async function handleImportText(text) {
    const parsed = E.parseProjectImport(text);
    if (!parsed.ok) {
      await window.ES_MODAL.alert({ title: '无法导入', icon: 'error', message: parsed.error });
      return;
    }
    let imported;
    try {
      imported = E.store.importProject(parsed.project);
    } catch (error) {
      await window.ES_MODAL.alert({
        title: '导入失败', icon: 'error',
        message: '写入项目数据失败：' + (error && error.message ? error.message : '未知错误'),
      });
      return;
    }
    const scenarioCount = Array.isArray(imported.scenarios) ? imported.scenarios.length : 0;
    const modelCount = Array.isArray(imported.models) ? imported.models.length : 0;
    await window.ES_MODAL.alert({
      title: '导入成功', icon: 'check_circle',
      message: '已作为独立副本导入项目「' + imported.name + '」\n场景 ' + scenarioCount + ' 个、模型 '
        + modelCount + ' 个。\n原有项目与本地数据均未改动。',
    });
    render();
  }

  // ---------- 项目看板骨架 ----------
  function renderBoard(project) {
    return '<div class="es-board">'
      + renderBatteryCard(project)
      + window.ES_ENDURANCE_SCENARIO.renderScenarioCard(project)
      + window.ES_ENDURANCE_MODEL.renderModelCard(project)
      + window.ES_ENDURANCE_UI.renderCompareCard(project)
      + '</div>';
  }

  // ---------- 电池参数 ----------
  const BATTERY_FIELDS = [
    { key: 'capTyp', label: '电池容量 mAh（Typ）', placeholder: '示例：246' },
    { key: 'capMin', label: '电池容量 Min（mAh）', placeholder: '示例：240' },
    { key: 'wh', label: '额定容量 Wh', placeholder: '示例：0.95' },
    { key: 'voltage', label: '额定电压 V', placeholder: '示例：3.87' },
    { key: 'dischargeRatio', label: '可放电容量比例(0.2C)', placeholder: '0~1，示例：0.95（可由查询表自动回填）' },
  ];

  function batteryDerateLine(project) {
    const b = project.battery || {};
    const cfg = projectConfig(project);
    const ratio = num(b.dischargeRatio);
    const threshold = num(cfg.lowThreshold);
    const capTyp = num(b.capTyp);
    const derate = (ratio === null ? 1 : ratio) * (1 - (threshold === null ? 0 : threshold));
    const available = capTyp === null ? null : capTyp * (ratio === null ? 1 : ratio);
    return '可用容量（公式名「可用容量」= Typ × 可放电比例）= <b>' + fmt(available, 2)
      + ' mAh</b>；综合可用电量系数 = 可放电比例 × (1 − 低电量门限) = <b>' + fmt(derate, 4) + '</b>';
  }

  function renderBatteryCard(project) {
    const b = project.battery || {};
    const fields = BATTERY_FIELDS.map((f) => {
      const val = b[f.key] === undefined || b[f.key] === null ? '' : b[f.key];
      return '<label class="es-battery-field">'
        + '<span>' + escapeHtml(f.label) + '</span>'
        + '<input type="text" inputmode="decimal" data-es-battery="' + f.key + '"'
        + ' value="' + escapeHtml(val) + '" placeholder="' + escapeHtml(f.placeholder) + '" autocomplete="off" />'
        + '</label>';
    }).join('');
    return '<section class="panel rounded-3xl es-card">'
      + '<header class="es-card-head"><h3><span class="material-symbols-outlined">battery_full</span>电池参数</h3>'
      + '<small>续航公式基于这些参数自动计算，修改后实时保存</small></header>'
      + '<div class="es-battery-grid">' + fields + '</div>'
      + '<p class="es-battery-derate">' + batteryDerateLine(project) + '</p>'
      + renderDischargeCard(project)
      + renderProjectConfigCard(project)
      + '</section>';
  }

  const saveBattery = debounce(() => E.store.save(), 150);

  function bindBattery(project) {
    root.querySelectorAll('[data-es-battery]').forEach((input) => {
      input.addEventListener('input', () => {
        const key = input.dataset.esBattery;
        const raw = input.value.trim();
        const parsed = raw === '' ? '' : num(raw);
        // 非数字、负数不允许写入；比例还必须落在 0~1
        let invalid = raw !== '' && (parsed === null || parsed < 0);
        if (!invalid && key === 'dischargeRatio' && raw !== '' && parsed > 1) {
          invalid = true;
        }
        if (invalid) {
          input.classList.add('is-invalid');
          return;
        }
        input.classList.remove('is-invalid');
        project.battery[key] = raw;
        saveBattery();
        const derateEl = root.querySelector('.es-battery-derate');
        if (derateEl) derateEl.innerHTML = batteryDerateLine(project);
        // Typ 容量变化时，查询表占 Typ 比例与可放电比例联动重算
        if (key === 'capTyp') {
          syncRatioFromTable();
        }
        // 电池参数变化会影响 DOU，局部刷新模型与对比区
        refreshModelAndCompare();
      });
    });
  }

  // 模型 / 对比区局部刷新（避免输入焦点被电池区重渲染打断）；防抖，连续输入只在停顿后重算一次
  const refreshModelAndCompare = debounce(function () {
    const project = E.store.activeProject();
    if (!project) return;
    const modelCard = root.querySelector('.es-model-card');
    const compareCard = root.querySelector('.es-compare-card');
    if (modelCard) {
      const next = document.createElement('div');
      next.innerHTML = window.ES_ENDURANCE_MODEL.renderModelCard(project);
      modelCard.replaceWith(next.firstElementChild);
      window.ES_ENDURANCE_MODEL.bindModels(project);
    }
    if (compareCard) {
      window.ES_ENDURANCE_UI.refreshCompareQuiet && window.ES_ENDURANCE_UI.refreshCompareQuiet();
    }
  }, 220);

  // ---------- 项目配置（关机电压 / 低电量门限） ----------
  function renderProjectConfigCard(project) {
    const cfg = projectConfig(project);
    return '<div class="es-submodule es-project-config">'
      + '<div class="es-submodule-head"><h4><span class="material-symbols-outlined">tune</span>项目配置</h4>'
      + '<small>整机口径，不随电池规格书变化</small></div>'
      + '<div class="es-config-grid">'
      + '<label class="es-battery-field"><span>关机电压 V</span>'
      + '<input type="text" inputmode="decimal" data-es-config="shutdownVoltage" value="' + escapeHtml(cfg.shutdownVoltage)
      + '" placeholder="示例：3.4" autocomplete="off" /></label>'
      + '<label class="es-battery-field"><span>低电量门限</span>'
      + '<input type="text" inputmode="decimal" data-es-config="lowThreshold" value="' + escapeHtml(cfg.lowThreshold)
      + '" placeholder="0~1，示例：0.02" autocomplete="off" /></label>'
      + '</div>'
      + '<p class="es-config-hint">关机电压用于在查询表中插值出可放电容量并回填「可放电容量比例」；低电量门限只参与综合可用电量系数，不改写电池容量。</p>'
      + '</div>';
  }

  function bindProjectConfig(project) {
    const cfg = projectConfig(project);
    root.querySelectorAll('[data-es-config]').forEach((input) => {
      input.addEventListener('input', () => {
        const key = input.dataset.esConfig;
        const raw = input.value.trim();
        const parsed = raw === '' ? '' : num(raw);
        let invalid = raw !== '' && (parsed === null || parsed < 0);
        if (!invalid && key === 'lowThreshold' && raw !== '' && parsed > 1) invalid = true;
        if (invalid) {
          input.classList.add('is-invalid');
          return;
        }
        input.classList.remove('is-invalid');
        cfg[key] = raw;
        saveBattery();
        const derateEl = root.querySelector('.es-battery-derate');
        if (derateEl) derateEl.innerHTML = batteryDerateLine(project);
        if (key === 'shutdownVoltage') syncRatioFromTable();
        refreshModelAndCompare();
      });
    });
  }

  // ---------- 电池可放电容量查询表 ----------
  // 依据关机电压 + 查询表自动回填可放电容量比例(0.2C)；只更新输入框值，不整卡重绘以避免丢焦点。
  // 防抖回调内重新取当前激活项目，避免快速切换项目后 pending 闭包把旧项目数据写进新 DOM。
  const syncRatioFromTable = debounce(function () {
    const project = E.store.activeProject();
    if (!project) return;
    const derived = E.deriveDischargeRatio(project);
    const syncEl = root.querySelector('.es-discharge-sync');
    const b = project.battery || {};
    // Typ 容量变化时，就地刷新矩阵每格的「占 Typ 比例」
    const capTypNow = num(b.capTyp);
    root.querySelectorAll('[data-es-dm-row]').forEach((tr) => {
      const row = ((b.dischargeMatrix || {}).rows || []).find((r) => r.id === tr.dataset.esDmRow);
      if (!row) return;
      tr.querySelectorAll('[data-es-dm-cell]').forEach((input) => {
        const pct = tr.querySelector('.es-dm-pct[data-es-dm-pct="' + input.dataset.esDmCell + '"]');
        if (!pct) return;
        const c = num(input.value);
        pct.textContent = (capTypNow && capTypNow > 0 && c !== null) ? fmt(c / capTypNow * 100, 1) + '%' : '—';
      });
    });
    if (!derived.ok) {
      if (syncEl) syncEl.innerHTML = dischargeSyncHint(project, derived.reason);
      return;
    }
    // 可放电容量超过 Typ（比例 >100%）通常意味着查询表 / Typ 填写有误，不回填以免污染 DOU
    if (derived.ratio !== null && derived.ratio > 1) {
      if (syncEl) syncEl.innerHTML = dischargeSyncHint(project, 'ratio-overflow', derived);
      return;
    }
    if (derived.ratio !== null) {
      const rounded = Math.round(derived.ratio * 10000) / 10000;
      b.dischargeRatio = String(rounded);
      E.store.save();
      const ratioInput = root.querySelector('[data-es-battery="dischargeRatio"]');
      if (ratioInput) {
        ratioInput.value = rounded;
        ratioInput.classList.remove('is-invalid');
      }
      const derateEl = root.querySelector('.es-battery-derate');
      if (derateEl) derateEl.innerHTML = batteryDerateLine(project);
      refreshModelAndCompare();
    }
    if (syncEl) syncEl.innerHTML = dischargeSyncHint(project, derived.ratio === null ? 'no-cap-typ' : '', derived);
  }, 180);

  function dischargeSyncHint(project, reason, derived) {
    const b = project.battery || {};
    const cfg = projectConfig(project);
    const matrix = b.dischargeMatrix || { rows: [] };
    const curve = dischargeCurve(b);
    const rowCount = (matrix.rows || []).length;
    const hint = (icon, html) => '<span class="material-symbols-outlined">' + icon
      + '</span><span class="es-hint-text">' + html + '</span>';
    if (!curve.length) {
      return hint('info', '尚未维护查询表：可导入电池规格书自动识别，或点「新增一行」按温度 / 放电倍率维护各截止电压的可放电容量。');
    }
    if (reason === 'no-shutdown-voltage') {
      return hint('hourglass_top', '已维护 ' + rowCount + ' 组温度 / 倍率数据、' + curve.length
        + ' 个截止电压；填写「项目配置 → 关机电压」后将自动插值计算可放电容量比例(0.2C)。');
    }
    const d = derived || E.deriveDischargeRatio(project);
    if (reason === 'no-cap-typ' || d.reason === 'no-cap-typ') {
      return hint('edit_note', '关机电压 <b>' + escapeHtml(cfg.shutdownVoltage) + ' V</b> → 可放电容量 <b>' + fmt(d.capacityMah, 2)
        + ' mAh</b>；填写电池容量 Typ 后将自动换算并回填比例。');
    }
    if (reason === 'ratio-overflow') {
      return hint('warning', '关机电压 <b>' + escapeHtml(cfg.shutdownVoltage) + ' V</b> → 可放电容量 <b>' + fmt(d.capacityMah, 2)
        + ' mAh</b>，占 Typ 容量 <b>' + fmt(d.ratio * 100, 2) + '%</b>，超过 100%，未自动回填；请检查查询表容量或 Typ 容量是否填写正确。');
    }
    return hint('check_circle', '关机电压 <b>' + escapeHtml(cfg.shutdownVoltage) + ' V</b> → 可放电容量 <b>' + fmt(d.capacityMah, 2)
      + ' mAh</b>（线性插值）→ 占 Typ 容量 <b>' + fmt(d.ratio * 100, 2) + '%</b>，已自动回填「可放电容量比例(0.2C)」。');
  }


  function renderDischargeCard(project) {
    const b = project.battery || {};
    const matrix = b.dischargeMatrix || { voltages: DEFAULT_CUTOFF_VOLTAGES.slice(), rows: [] };
    const voltages = matrixVoltages(matrix);
    const capTyp = num(b.capTyp);
    const derived = E.deriveDischargeRatio(project);
    const rows = matrix.rows || [];
    const pctOf = (v) => {
      const c = num(v);
      return (capTyp && capTyp > 0 && c !== null) ? fmt(c / capTyp * 100, 1) + '%' : '—';
    };
    const body = rows.length ? rows.map((row) => {
      const cells = voltages.map((v) => {
        const val = matrixCellValue(row, v);
        return '<td class="es-dm-cell">'
          + '<input type="text" inputmode="decimal" class="es-dt-input" data-es-dm-cell="' + v + '"'
          + ' value="' + escapeHtml(val === null ? '' : val) + '" placeholder="mAh" autocomplete="off" />'
          + '<span class="es-dm-pct" data-es-dm-pct="' + v + '">' + pctOf(val) + '</span>'
          + '</td>';
      }).join('');
      return '<tr data-es-dm-row="' + escapeHtml(row.id) + '">'
        + '<td><input type="text" class="es-dt-input" data-es-dm-col="temp" value="'
        + escapeHtml(row.temp === undefined || row.temp === null ? '' : row.temp) + '" placeholder="示例：25度" autocomplete="off" /></td>'
        + '<td><input type="text" class="es-dt-input" data-es-dm-col="cRate" value="'
        + escapeHtml(row.cRate === undefined || row.cRate === null ? '' : row.cRate) + '" placeholder="示例：0.2C" autocomplete="off" /></td>'
        + cells
        + '<td><button type="button" class="es-icon-btn es-icon-danger" data-es-dm-delete="' + escapeHtml(row.id)
        + '" title="删除该行"><span class="material-symbols-outlined">close</span></button></td>'
        + '</tr>';
    }).join('') : '<tr><td colspan="' + (voltages.length + 3)
      + '" class="es-table-empty">暂无数据，导入电池规格书或点「新增一行」维护温度 / 放电倍率下的可放电容量</td></tr>';

    const headVoltages = voltages.map((v) => '<th class="es-dm-vhead">'
      + '<span class="es-dm-vwrap"><input type="text" inputmode="decimal" class="es-dm-vinput" data-es-dm-voltage="' + v
      + '" value="' + escapeHtml(v) + '" autocomplete="off" /><small>V</small>'
      + '<button type="button" class="es-icon-btn es-icon-danger es-dm-vdel" data-es-dm-vdel="' + v
      + '" title="删除该电压列"><span class="material-symbols-outlined">close</span></button></span></th>').join('');

    return '<div class="es-submodule es-discharge-sub">'
      + '<div class="es-submodule-head"><h4><span class="material-symbols-outlined">table_chart</span>电池可放电容量查询表</h4>'
      + '<div class="es-head-actions">'
      + '<button type="button" class="template-btn template-btn-sm" data-es-dt-import><span class="material-symbols-outlined text-sm">upload_file</span>导入文档自动识别</button>'
      + '<input type="file" accept=".xlsx,.xls,.csv" hidden data-es-dt-file />'
      + '<button type="button" class="template-btn template-btn-sm" data-es-dt-add><span class="material-symbols-outlined text-sm">add</span>新增一行</button>'
      + '<button type="button" class="template-btn template-btn-sm" data-es-dt-add-v><span class="material-symbols-outlined text-sm">add_column_right</span>新增电压列</button>'
      + '<button type="button" class="template-btn template-btn-sm template-btn-danger" data-es-dt-clear><span class="material-symbols-outlined text-sm">delete_sweep</span>清空</button>'
      + '</div></div>'
      + '<p class="es-discharge-tip"><span class="es-hint-text">按电池规格书格式维护：每行一个「温度 + 放电倍率」，各截止电压列填<b>从满电放至该电压的累计可放电容量（mAh）</b>，格内小字为占 Typ 容量比例。可放电比例按 0.2C 行插值，支持 .xlsx / .xls / .csv 导入。</span></p>'
      + '<div class="es-table-wrap"><table class="es-table es-discharge-table es-matrix-table"><thead>'
      + '<tr><th rowspan="2">温度</th><th rowspan="2">放电倍率</th>'
      + '<th colspan="' + voltages.length + '">截止电压（可放电 mAh / 占 Typ）</th><th rowspan="2" style="width:56px">操作</th></tr>'
      + '<tr>' + headVoltages + '</tr>'
      + '</thead><tbody>' + body + '</tbody></table></div>'
      + '<p class="es-discharge-sync">' + dischargeSyncHint(project, derived.reason, derived.ok ? derived : null) + '</p>'
      + '</div>';
  }

  function bindDischargeCard(project) {
    const b = project.battery || (project.battery = E.defaultBattery());
    if (!Array.isArray(b.dischargeTable)) b.dischargeTable = [];
    if (!b.dischargeMatrix || typeof b.dischargeMatrix !== 'object') {
      b.dischargeMatrix = { voltages: DEFAULT_CUTOFF_VOLTAGES.slice(), rows: [] };
    }
    if (!Array.isArray(b.dischargeMatrix.rows)) b.dischargeMatrix.rows = [];
    if (!Array.isArray(b.dischargeMatrix.voltages) || !b.dischargeMatrix.voltages.length) {
      b.dischargeMatrix.voltages = DEFAULT_CUTOFF_VOLTAGES.slice();
    }
    const matrix = b.dischargeMatrix;

    const fileInput = root.querySelector('[data-es-dt-file]');
    root.querySelector('[data-es-dt-import]').addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', async (event) => {
      const file = event.target.files && event.target.files[0];
      event.target.value = '';
      if (!file) return;
      try {
        const parsed = await parseDischargeFile(file, num(b.capTyp));
        // 导入的电压 - 容量点作为「25度 / 0.2C」一行并入矩阵，并对齐矩阵电压列
        const voltages = parsed.rows.map((r) => r.voltage);
        matrixVoltages(matrix).forEach((v) => { if (voltages.indexOf(v) === -1) voltages.push(v); });
        matrix.voltages = voltages.sort((x, y) => x - y);
        const mah = {};
        parsed.rows.forEach((r) => { mah[String(r.voltage)] = r.capacityMah; });
        matrix.rows = [{ id: uid('dt'), temp: '25度', cRate: '0.2C', mah: mah }];
        E.store.save();
        render();
        await window.ES_MODAL.alert({
          title: '导入成功', icon: 'check_circle',
          message: '从「' + parsed.sheetName + '」导入 ' + parsed.rows.length + ' 个截止电压容量，作为「25度 / 0.2C」一行。\n' + parsed.note,
        });
        syncRatioFromTable();
      } catch (error) {
        await window.ES_MODAL.alert({
          title: '导入失败', icon: 'error',
          message: error && error.message ? error.message : '无法识别文件内容',
        });
      }
    });

    root.querySelector('[data-es-dt-add]').addEventListener('click', () => {
      const cells = {};
      matrixVoltages(matrix).forEach((v) => { cells[String(v)] = ''; });
      const last = matrix.rows[matrix.rows.length - 1];
      matrix.rows.push({ id: uid('dt'), temp: last ? last.temp : '25度', cRate: last ? last.cRate : '0.2C', mah: cells });
      E.store.save();
      render();
    });

    root.querySelector('[data-es-dt-add-v]').addEventListener('click', () => {
      const vs = matrixVoltages(matrix);
      const next = vs.length ? Math.round((vs[vs.length - 1] + 0.1) * 100) / 100 : DEFAULT_CUTOFF_VOLTAGES[0];
      matrix.voltages = vs.concat([next]);
      matrix.rows.forEach((row) => { matrixSetCell(row, next, ''); });
      E.store.save();
      render();
    });

    root.querySelector('[data-es-dt-clear]').addEventListener('click', async () => {
      if (!matrix.rows.length) return;
      const ok = await window.ES_MODAL.confirm({
        title: '清空查询表', icon: 'delete_sweep',
        message: '确认清空电池可放电容量查询表？（电压列会保留）',
        confirmText: '清空',
      });
      if (!ok) return;
      matrix.rows = [];
      E.store.save();
      render();
      syncRatioFromTable();
    });

    root.querySelectorAll('[data-es-dm-delete]').forEach((btn) => {
      btn.addEventListener('click', () => {
        matrix.rows = matrix.rows.filter((r) => r.id !== btn.dataset.esDmDelete);
        E.store.save();
        render();
        syncRatioFromTable();
      });
    });

    // 删除某一电压列（同时删除所有行里对应的容量）
    root.querySelectorAll('[data-es-dm-vdel]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const v = num(btn.dataset.esDmVdel);
        if (v === null) return;
        matrix.voltages = matrixVoltages(matrix).filter((x) => x !== v);
        matrix.rows.forEach((row) => {
          const cells = matrixCells(row);
          Object.keys(cells).forEach((k) => { if (num(k) === v) delete cells[k]; });
        });
        E.store.save();
        render();
        syncRatioFromTable();
      });
    });

    // 电压列表头编辑：改完（回车 / 失焦）后重映射所有行的 key
    root.querySelectorAll('[data-es-dm-voltage]').forEach((input) => {
      input.addEventListener('change', () => {
        const from = num(input.dataset.esDmVoltage);
        const to = num(input.value);
        const list = matrixVoltages(matrix);
        if (to === null || to <= 0 || (from !== null && list.indexOf(to) !== -1 && to !== from)) {
          input.classList.add('is-invalid');
          input.value = input.dataset.esDmVoltage;
          return;
        }
        input.classList.remove('is-invalid');
        matrix.voltages = list.map((v) => (v === from ? to : v));
        if (from !== null) {
          matrix.rows.forEach((row) => {
            const cells = matrixCells(row);
            Object.keys(cells).forEach((k) => {
              if (num(k) === from) {
                const val = cells[k];
                delete cells[k];
                cells[String(to)] = val;
              }
            });
          });
        }
        E.store.save();
        render();
        syncRatioFromTable();
      });
    });

    // 行内编辑：温度 / 倍率 / 各截止电压容量；非法值标红不入库
    root.querySelectorAll('[data-es-dm-col]').forEach((input) => {
      input.addEventListener('input', () => {
        const tr = input.closest('[data-es-dm-row]');
        const row = matrix.rows.find((r) => r.id === tr.dataset.esDmRow);
        if (!row) return;
        row[input.dataset.esDmCol] = input.value;
        E.store.save();
      });
    });

    root.querySelectorAll('[data-es-dm-cell]').forEach((input) => {
      input.addEventListener('input', () => {
        const tr = input.closest('[data-es-dm-row]');
        const row = matrix.rows.find((r) => r.id === tr.dataset.esDmRow);
        if (!row) return;
        const raw = input.value.trim();
        const parsed = raw === '' ? '' : num(raw);
        if (raw !== '' && (parsed === null || parsed < 0)) {
          input.classList.add('is-invalid');
          return;
        }
        input.classList.remove('is-invalid');
        matrixSetCell(row, input.dataset.esDmCell, raw);
        E.store.save();
        const capTyp = num(b.capTyp);
        const pct = tr.querySelector('.es-dm-pct[data-es-dm-pct="' + input.dataset.esDmCell + '"]');
        if (pct) pct.textContent = (capTyp && capTyp > 0 && parsed !== '') ? fmt(parsed / capTyp * 100, 1) + '%' : '—';
        syncRatioFromTable();
      });
    });
  }

  // 解析电池规格书：智能识别电压列与容量列（支持可放电容量 / 剩余容量 / 百分比 SOC）
  async function parseDischargeFile(file, capTyp) {
    if (typeof XLSX === 'undefined' || !XLSX.utils) {
      throw new Error('表格解析组件未加载，请检查网络后重试');
    }
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array' });
    let lastError = '未找到同时包含「电压」与「容量」列的工作表';
    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      try {
        const parsed = parseDischargeAoa(aoa, capTyp);
        return { rows: parsed.rows, note: parsed.note, sheetName: sheetName };
      } catch (error) {
        lastError = error.message;
      }
    }
    throw new Error(lastError + '。请确认表头包含「电压 / V」与「容量 / mAh（或 SOC 百分比）」字样。');
  }

  function parseDischargeAoa(aoa, capTyp) {
    const cellText = (v) => String(v == null ? '' : v).trim().toLowerCase();
    const isVoltageHeader = (v) => {
      const s = cellText(v);
      if (!s) return false;
      if (s.includes('电压') || s.includes('voltage')) return true;
      return s === 'v' || s === '(v)' || s === '（v）';
    };
    const scanLimit = Math.min(aoa.length, 25);
    let headerRowIdx = -1;
    let voltageCol = -1;
    let capacityCol = -1;
    let capacityKind = ''; // discharge-mah / remaining-mah / discharge-pct / remaining-pct
    for (let r = 0; r < scanLimit; r += 1) {
      const row = aoa[r] || [];
      const vCol = row.findIndex(isVoltageHeader);
      if (vCol === -1) continue;
      let best = -1;
      let bestKind = '';
      let bestScore = -1;
      row.forEach((cell, c) => {
        if (c === vCol) return;
        const s = cellText(cell);
        if (!s) return;
        const isPct = /%|百分比|百分|比例|percent|ratio|\bsoc\b/.test(s);
        const isRemaining = /剩余|余量|soc|remain|left/.test(s);
        const isCapacity = /容量|mah|capacity|放电/.test(s);
        if (!isPct && !isCapacity) return;
        // 优先级：可放电 mAh > 剩余 mAh > 可放电百分比 > 剩余百分比
        const kind = isPct
          ? (isRemaining ? 'remaining-pct' : 'discharge-pct')
          : (isRemaining ? 'remaining-mah' : 'discharge-mah');
        const score = { 'discharge-mah': 4, 'remaining-mah': 3, 'discharge-pct': 2, 'remaining-pct': 1 }[kind];
        if (score > bestScore) { bestScore = score; best = c; bestKind = kind; }
      });
      if (best !== -1) {
        headerRowIdx = r; voltageCol = vCol; capacityCol = best; capacityKind = bestKind;
        break;
      }
    }
    if (headerRowIdx === -1) throw new Error('该 Sheet 未识别到电压 / 容量表头');

    const needTyp = capacityKind !== 'discharge-mah';
    if (needTyp && (!capTyp || capTyp <= 0)) {
      throw new Error('该表格为剩余容量或百分比口径，需先在电池参数中填写「电池容量 mAh（Typ）」才能换算');
    }
    const rows = [];
    const seen = new Set();
    for (let r = headerRowIdx + 1; r < aoa.length; r += 1) {
      const row = aoa[r] || [];
      const voltage = num(row[voltageCol]);
      const rawCap = num(row[capacityCol]);
      if (voltage === null || rawCap === null) continue;
      if (voltage <= 0 || voltage > 10 || rawCap < 0) continue;
      let capacityMah;
      if (capacityKind === 'discharge-mah') {
        capacityMah = rawCap;
      } else if (capacityKind === 'remaining-mah') {
        capacityMah = capTyp - rawCap;
      } else {
        const ratio = rawCap > 1 ? rawCap / 100 : rawCap;
        capacityMah = (capacityKind === 'remaining-pct' ? 1 - ratio : ratio) * capTyp;
      }
      if (capacityMah < 0) capacityMah = 0;
      const key = String(voltage);
      if (seen.has(key)) continue;
      seen.add(key);
      rows.push({ voltage: voltage, capacityMah: Math.round(capacityMah * 1000) / 1000 });
    }
    if (!rows.length) throw new Error('表头下未读到有效的电压 / 容量数值行');
    rows.sort((a, b) => b.voltage - a.voltage);
    const kindText = {
      'discharge-mah': '容量列为可放电容量（mAh），直接采用',
      'remaining-mah': '容量列为剩余容量，已按 Typ 容量换算为可放电容量',
      'discharge-pct': '容量列为可放电比例，已按 Typ 容量换算为 mAh',
      'remaining-pct': '容量列为剩余 SOC，已按 (1 − SOC) × Typ 容量换算',
    }[capacityKind];
    return { rows: rows, note: kindText };
  }

  window.ES_ENDURANCE_UI = { render, refreshModelAndCompare, syncRatioFromTable };
})();

/* ================================================================
 * 第三部分：场景电流（功耗标准 / 自定义电流 / 版本日期 / 备注 / 自定义列）
 * ================================================================ */
(function () {
  'use strict';
  const E = window.ES_ENDURANCE;
  const { escapeHtml, uid, num, fmt, todayStr } = E;
  const root = document.getElementById('enduranceStatusRoot');

  function renderScenarioCard(project) {
    const customFields = project.scenarioFields || [];
    const head = '<th>场景</th><th>功耗标准</th><th>当前电流</th><th>单场景续航</th><th>测试版本</th><th>更新日期</th><th>备注/说明</th>'
      + customFields.map((f) => '<th class="es-custom-th">' + escapeHtml(f.name)
        + '<button type="button" class="es-field-del" data-es-field-delete="' + escapeHtml(f.id) + '" title="删除该列"><span class="material-symbols-outlined">close</span></button></th>').join('')
      + '<th class="es-col-ops">管理</th>';

    const rows = project.scenarios.map((scenario) => {
      const rec = E.latestRecord(scenario);
      const custom = num(scenario.customCurrent);
      const isCustom = custom !== null;
      const current = isCustom ? custom : (rec ? num(rec.current) : null);
      const currentText = current === null ? '' : String(current);
      const versionText = rec && rec.version ? String(rec.version) : '';
      const dateText = rec && rec.date ? String(rec.date) : '';
      // 方案 A：当前电流相对功耗标准的达标徽标（无标准 / 无电流则不渲染）
      const cmp = E.currentCompare(current, scenario.standard);
      const cmpBadge = cmp
        ? '<em class="es-cmp-badge es-cmp-' + cmp.level + '" title="' + escapeHtml(cmp.title) + '">' + escapeHtml(cmp.text) + '</em>'
        : '';
      const currentCell = '<td class="es-sc-current">'
        + '<div class="es-inline-input-wrap">'
        + '<input type="text" inputmode="decimal" class="es-field-input es-current-input" data-es-current="' + escapeHtml(scenario.id) + '"'
        + ' data-es-current-orig="' + escapeHtml(isCustom ? currentText : '') + '"'
        + ' value="' + escapeHtml(currentText) + '" placeholder="自定义电流 mA" autocomplete="off" />'
        + cmpBadge
        + (isCustom ? '<em class="es-custom-chip" title="自定义电流，不来自测试记录">自定义</em>'
          + '<button type="button" class="es-icon-btn es-inline-clear" data-es-action="clear-custom" title="清除自定义，恢复测试记录电流"><span class="material-symbols-outlined">restart_alt</span></button>' : '')
        + '</div></td>';
      // 单场景续航：每场景一条独立公式（默认留空）。空显示「未设置」，错误显示错误码，fx 打开专属编辑器
      const sl = E.evaluateSingleLife(project, scenario);
      const slIsFormula = String(scenario.singleLife || '').trim().charAt(0) === '=';
      const slCls = 'es-singlelife-cell' + (slIsFormula ? ' is-formula' : '') + (sl.ok ? '' : ' is-error');
      let slInner;
      if (!sl.ok) {
        slInner = '<span class="es-cell-text es-cell-text-error" title="' + escapeHtml(sl.error) + '">' + escapeHtml(sl.error) + '</span>';
      } else if (sl.empty) {
        slInner = '<span class="es-cell-empty">未设置</span>';
      } else {
        slInner = '<span class="es-cell-text">' + escapeHtml(fmt(sl.value, 4)) + '</span>';
      }
      const slCell = '<td class="' + slCls + '">' + slInner
        + '<button type="button" class="es-fx-btn" data-es-singlelife-fx="' + escapeHtml(scenario.id)
        + '" title="编辑单场景续航公式 / 数值">fx</button></td>';
      const customCells = customFields.map((f) => '<td><input type="text" class="es-field-input" data-es-field-row="' + escapeHtml(scenario.id)
        + '" data-es-field-col="' + escapeHtml(f.id) + '" value="' + escapeHtml(scenario.fields && scenario.fields[f.id] || '') + '" autocomplete="off" /></td>').join('');
      return '<tr data-es-scenario="' + escapeHtml(scenario.id) + '">'
        + '<td class="es-sc-name">' + escapeHtml(scenario.name)
        + '<button type="button" class="es-icon-btn es-inline-edit" data-es-action="rename-scenario" title="重命名场景"><span class="material-symbols-outlined">drive_file_rename_outline</span></button></td>'
        + '<td><input type="text" class="es-field-input es-standard-input" data-es-standard="' + escapeHtml(scenario.id)
        + '" value="' + escapeHtml(scenario.standard || '') + '" placeholder="功耗标准 / 门限" autocomplete="off" /></td>'
        + currentCell
        + slCell
        + '<td><input type="text" class="es-field-input es-version-input" data-es-version="' + escapeHtml(scenario.id)
        + '" data-es-version-orig="' + escapeHtml(versionText)
        + '" value="' + escapeHtml(versionText) + '" placeholder="测试版本" autocomplete="off" /></td>'
        + '<td><input type="text" class="es-field-input es-date-input" data-es-date="' + escapeHtml(scenario.id)
        + '" data-es-date-orig="' + escapeHtml(dateText)
        + '" value="' + escapeHtml(dateText) + '" placeholder="YYYY-MM-DD" autocomplete="off" /></td>'
        + '<td><input type="text" class="es-field-input es-remark-input" data-es-remark="' + escapeHtml(scenario.id)
        + '" value="' + escapeHtml(scenario.remark || '') + '" placeholder="备注 / 说明" autocomplete="off" /></td>'
        + customCells
        + '<td class="es-sc-ops">'
        + '<button type="button" class="es-mini-btn" data-es-action="update-current" title="按日期/版本录入测试电流（留痕）"><span class="material-symbols-outlined">edit_note</span>更新</button>'
        + '<button type="button" class="es-mini-btn" data-es-action="history" title="查看历史记录"><span class="material-symbols-outlined">history</span>历史</button>'
        + '<button type="button" class="es-mini-btn es-mini-danger" data-es-action="delete-scenario" title="删除场景"><span class="material-symbols-outlined">close</span></button>'
        + '</td></tr>';
    }).join('');

    const colCount = 8 + customFields.length;
    const emptyHtml = '<tr><td colspan="' + colCount + '" class="es-table-empty es-scenario-empty">'
      + '<div class="es-empty-msg">还没有场景，可载入云表默认的 12 个标准场景（含 UT 阶段场景电流），也可以手动新增场景</div>'
      + '<div class="es-empty-actions">'
      + '<button type="button" class="template-btn template-btn-primary template-btn-sm" data-es-action="load-default-scenarios">'
      + '<span class="material-symbols-outlined text-sm">download</span>载入默认 12 场景（含电流）</button>'
      + '<button type="button" class="template-btn template-btn-sm es-add-scenario-btn"><span class="material-symbols-outlined text-sm">add</span>手动新增场景</button>'
      + '</div></td></tr>';
    return '<section class="panel rounded-3xl es-card">'
      + '<header class="es-card-head"><h3><span class="material-symbols-outlined">bolt</span>场景电流</h3>'
      + '<div class="es-head-actions">'
      + '<button type="button" class="template-btn template-btn-sm" data-es-action="load-default-scenarios"><span class="material-symbols-outlined text-sm">download</span>载入参考场景</button>'
      + '<button type="button" class="template-btn template-btn-sm" data-es-add-field><span class="material-symbols-outlined text-sm">view_column</span>新增列</button>'
      + '<button type="button" class="template-btn template-btn-sm es-add-scenario-btn"><span class="material-symbols-outlined text-sm">add</span>新增场景</button>'
      + '</div></header>'
      + '<div class="es-table-wrap"><table class="es-table es-scenario-table"><thead><tr>' + head + '</tr></thead><tbody>'
      + (rows || emptyHtml)
      + '</tbody></table></div></section>';
  }

  function isValidDateText(text) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) return false;
    const y = Number(text.slice(0, 4));
    const m = Number(text.slice(5, 7));
    const d = Number(text.slice(8, 10));
    if (m < 1 || m > 12 || d < 1) return false;
    const dt = new Date(y, m - 1, d);
    return dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d;
  }

  // 方案 A 徽标就地刷新：直接读行内输入框当前值，标准 / 电流编辑时即时更新，无需整卡重绘
  function refreshRowBadge(tr) {
    if (!tr) return;
    const stdInput = tr.querySelector('.es-standard-input');
    const curInput = tr.querySelector('.es-current-input');
    const wrap = tr.querySelector('.es-sc-current .es-inline-input-wrap');
    if (!stdInput || !curInput || !wrap) return;
    const cmp = E.currentCompare(curInput.value, stdInput.value);
    let badge = wrap.querySelector('.es-cmp-badge');
    if (!cmp) {
      if (badge) badge.remove();
      return;
    }
    if (!badge) {
      badge = document.createElement('em');
      badge.className = 'es-cmp-badge';
      const input = wrap.querySelector('.es-current-input');
      wrap.insertBefore(badge, input ? input.nextSibling : wrap.firstChild);
    }
    badge.className = 'es-cmp-badge es-cmp-' + cmp.level;
    badge.title = cmp.title;
    badge.textContent = cmp.text;
  }

  function bindScenarios(project) {
    root.querySelectorAll('.es-add-scenario-btn').forEach((addBtn) => {
      addBtn.addEventListener('click', () => openScenarioEditor(project, null));
    });

    // 载入云表参考场景：任意项目可用，同名场景原样保留、仅补齐缺失项（不覆盖已有电流）
    root.querySelectorAll('[data-es-action="load-default-scenarios"]').forEach((loadBtn) => {
      loadBtn.addEventListener('click', async () => {
        const existingNames = new Set(project.scenarios.map((s) => s.name));
        const defs = E.defaultScenarioDefs();
        const missing = defs.filter((d) => !existingNames.has(d[0]));
        const skipCount = defs.length - missing.length;
        if (!missing.length) {
          await window.ES_MODAL.alert({
            title: '参考场景已齐全', icon: 'download',
            message: '云表 12 个参考场景已全部存在，无需重复载入；如需更新电流可直接在「当前电流」中修改。',
          });
          return;
        }
        const message = '将补齐 ' + missing.length + ' 个缺失的参考场景，并附上 UT 阶段场景电流（mA）：\n'
          + missing.map((d) => d[0]).join('、')
          + (skipCount ? '\n另有 ' + skipCount + ' 个同名场景已存在，会原样保留、不会覆盖其电流。' : '')
          + '\n载入后仍可随时修改或删除。';
        const ok = await window.ES_MODAL.confirm({
          title: '载入参考场景', icon: 'download', message: message, confirmText: '载入',
        });
        if (!ok) return;
        const result = E.mergeDefaultScenarios(project);
        E.store.save();
        window.ES_ENDURANCE_UI.render();
        if (result.added.length) {
          await window.ES_MODAL.alert({
            title: '载入完成', icon: 'check_circle',
            message: '已新增 ' + result.added.length + ' 个参考场景'
              + (result.skipped.length ? '，保留 ' + result.skipped.length + ' 个同名场景。' : '。'),
          });
        }
      });
    });

    root.querySelector('[data-es-add-field]').addEventListener('click', async () => {
      if (!Array.isArray(project.scenarioFields)) project.scenarioFields = [];
      const name = await window.ES_MODAL.prompt({
        title: '新增自定义列', icon: 'view_column', label: '列名称',
        placeholder: '如：功耗来源、测试人', confirmText: '新增列',
        validate: (raw) => {
          const trimmed = (raw || '').trim();
          if (!trimmed) return '列名不能为空';
          if (project.scenarioFields.some((f) => f.name === trimmed)) return '已存在同名列';
          return null;
        },
      });
      if (name === null) return;
      project.scenarioFields.push({ id: uid('sf'), name: name.trim() });
      E.store.save();
      window.ES_ENDURANCE_UI.render();
    });

    root.querySelectorAll('[data-es-field-delete]').forEach((btn) => {
      btn.addEventListener('click', async () => {
        const fid = btn.dataset.esFieldDelete;
        const field = (project.scenarioFields || []).find((f) => f.id === fid);
        if (!field) return;
        const ok = await window.ES_MODAL.confirm({
          title: '删除自定义列', icon: 'delete',
          message: '确认删除列「' + field.name + '」？各场景在该列填写的内容会一并删除。',
        });
        if (!ok) return;
        project.scenarioFields = project.scenarioFields.filter((f) => f.id !== fid);
        project.scenarios.forEach((s) => { if (s.fields) delete s.fields[fid]; });
        E.store.save();
        window.ES_ENDURANCE_UI.render();
      });
    });

    // 行内文本字段：失焦保存，不重绘卡片以保留操作焦点
    root.querySelectorAll('.es-standard-input').forEach((input) => {
      input.addEventListener('blur', () => {
        const scenario = project.scenarios.find((s) => s.id === input.dataset.esStandard);
        if (!scenario) return;
        if (scenario.standard !== input.value) {
          scenario.standard = input.value.trim();
          E.store.save();
        }
        refreshRowBadge(input.closest('tr'));
      });
    });
    // 方案 A：功耗标准 / 当前电流输入时，徽标就地实时联动（不打断输入焦点）
    root.querySelectorAll('.es-scenario-table tbody tr[data-es-scenario]').forEach((tr) => {
      const std = tr.querySelector('.es-standard-input');
      const cur = tr.querySelector('.es-current-input');
      if (std) std.addEventListener('input', () => refreshRowBadge(tr));
      if (cur) cur.addEventListener('input', () => refreshRowBadge(tr));
    });
    root.querySelectorAll('.es-remark-input').forEach((input) => {
      input.addEventListener('blur', () => {
        const scenario = project.scenarios.find((s) => s.id === input.dataset.esRemark);
        if (!scenario || scenario.remark === input.value) return;
        scenario.remark = input.value.trim();
        E.store.save();
      });
    });
    // 当前电流：内联编辑自定义电流。
    // 仅当值相对渲染时的原始值发生变化才写库，避免未编辑就把测试记录电流误转为自定义；
    // 留空且原本是自定义时清除自定义并完整重绘，恢复测试记录值与「自定义」标记。
    root.querySelectorAll('.es-current-input').forEach((input) => {
      input.addEventListener('blur', () => {
        const scenario = project.scenarios.find((s) => s.id === input.dataset.esCurrent);
        if (!scenario) return;
        const trimmed = input.value.trim();
        const orig = input.dataset.esCurrentOrig || '';
        if (trimmed === orig) { input.classList.remove('is-error'); return; }
        if (trimmed === '') {
          input.classList.remove('is-error');
          if (scenario.customCurrent !== '') {
            scenario.customCurrent = '';
            E.store.save();
            window.ES_ENDURANCE_UI.render();
          }
          return;
        }
        const v = num(trimmed);
        if (v === null || v < 0) {
          input.classList.add('is-error');
          return;
        }
        input.classList.remove('is-error');
        scenario.customCurrent = String(v);
        E.store.save();
        window.ES_ENDURANCE_UI.render();
      });
    });
    // 单场景续航：fx 打开专属公式编辑器（每场景一条独立公式，默认留空）
    root.querySelectorAll('[data-es-singlelife-fx]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const target = project.scenarios.find((s) => s.id === btn.dataset.esSinglelifeFx);
        if (target) openSingleLifeEditor(project, target);
      });
    });
    // 测试版本 / 更新日期：内联编辑最新一条测试记录。
    // 无记录且填写了内容才创建记录并立即落库；空值不产生空记录。
    function upsertLatestRecord(scenario) {
      let rec = E.latestRecord(scenario);
      if (rec) return rec;
      rec = {
        id: uid('rec'),
        date: E.todayStr(),
        version: '',
        current: E.scenarioCurrent(project, scenario.id) === null ? '' : String(E.scenarioCurrent(project, scenario.id)),
        note: '',
      };
      scenario.records.push(rec);
      E.store.save();
      return rec;
    }
    root.querySelectorAll('.es-version-input').forEach((input) => {
      input.addEventListener('blur', () => {
        const scenario = project.scenarios.find((s) => s.id === input.dataset.esVersion);
        if (!scenario) return;
        const trimmed = input.value.trim();
        if (trimmed === (input.dataset.esVersionOrig || '')) { input.classList.remove('is-error'); return; }
        if (trimmed === '' && !E.latestRecord(scenario)) return;
        const rec = upsertLatestRecord(scenario);
        rec.version = trimmed;
        E.store.save();
        window.ES_ENDURANCE_UI.render();
      });
    });
    root.querySelectorAll('.es-date-input').forEach((input) => {
      input.addEventListener('blur', () => {
        const scenario = project.scenarios.find((s) => s.id === input.dataset.esDate);
        if (!scenario) return;
        const trimmed = input.value.trim();
        if (trimmed === (input.dataset.esDateOrig || '')) { input.classList.remove('is-error'); return; }
        if (trimmed === '') {
          input.classList.remove('is-error');
          const rec0 = E.latestRecord(scenario);
          if (rec0 && rec0.date) { rec0.date = ''; E.store.save(); window.ES_ENDURANCE_UI.render(); }
          return;
        }
        if (!isValidDateText(trimmed)) {
          input.classList.add('is-error');
          return;
        }
        input.classList.remove('is-error');
        const rec = upsertLatestRecord(scenario);
        rec.date = trimmed;
        scenario.records.sort((a, b) => String(a.date).localeCompare(String(b.date)));
        E.store.save();
        window.ES_ENDURANCE_UI.render();
      });
    });
    root.querySelectorAll('.es-field-input[data-es-field-row]').forEach((input) => {
      input.addEventListener('blur', () => {
        const scenario = project.scenarios.find((s) => s.id === input.dataset.esFieldRow);
        if (!scenario) return;
        if (!scenario.fields || typeof scenario.fields !== 'object') scenario.fields = {};
        if ((scenario.fields[input.dataset.esFieldCol] || '') === input.value) return;
        scenario.fields[input.dataset.esFieldCol] = input.value;
        E.store.save();
      });
    });

    root.querySelectorAll('[data-es-scenario]').forEach((tr) => {
      const id = tr.dataset.esScenario;
      tr.querySelectorAll('[data-es-action]').forEach((btn) => {
        btn.addEventListener('click', () => {
          const action = btn.dataset.esAction;
          if (action === 'update-current') openCurrentEditor(project, id);
          else if (action === 'history') openHistory(project, id);
          else if (action === 'rename-scenario') openScenarioEditor(project, project.scenarios.find((s) => s.id === id));
          else if (action === 'clear-custom') clearCustomCurrent(project, id);
          else if (action === 'delete-scenario') deleteScenario(project, id);
        });
      });
    });
  }

  async function deleteScenario(project, id) {
    const scenario = project.scenarios.find((s) => s.id === id);
    if (!scenario) return;
    const ok = await window.ES_MODAL.confirm({
      title: '删除场景', icon: 'delete',
      message: '确认删除场景「' + scenario.name + '」及其全部电流记录？续航模型中引用该场景的行会一并删除。',
    });
    if (!ok) return;
    project.scenarios = project.scenarios.filter((s) => s.id !== id);
    project.models.forEach((model) => {
      const rowIndex = model.rows.findIndex((r) => r.scenarioId === id);
      if (rowIndex >= 0) E.rewriteRefsAfterRowDelete(model, rowIndex);
      model.rows = model.rows.filter((r) => r.scenarioId !== id);
    });
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  // 新增 / 重命名场景；新增时同步给该项目每个模型加一行
  async function openScenarioEditor(project, scenario) {
    const name = await window.ES_MODAL.prompt({
      title: scenario ? '重命名场景' : '新增场景',
      icon: scenario ? 'drive_file_rename_outline' : 'add',
      label: '场景名称',
      value: scenario ? scenario.name : '',
      placeholder: scenario ? '' : '请输入场景名称',
      confirmText: scenario ? '保存' : '新增',
      validate: (raw) => ((raw || '').trim() ? null : '场景名称不能为空'),
    });
    if (name === null) return;
    const trimmed = name.trim();
    if (scenario) {
      scenario.name = trimmed;
    } else {
      const created = {
        id: uid('sc'), name: trimmed, remark: '', standard: '',
        customCurrent: '', fields: {}, singleLife: '', records: [],
      };
      project.scenarios.push(created);
      project.models.forEach((model) => E.appendRowForScenario(model, created.id));
    }
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  function clearCustomCurrent(project, scenarioId) {
    const scenario = project.scenarios.find((s) => s.id === scenarioId);
    if (!scenario) return;
    scenario.customCurrent = '';
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  // 电流更新弹窗：一个表单同时录入日期 / 版本 / 电流 / 备注，确认后追加一条记录（历史不覆盖）
  async function openCurrentEditor(project, scenarioId) {
    const scenario = project.scenarios.find((s) => s.id === scenarioId);
    if (!scenario) return;
    const last = E.latestRecord(scenario);
    const values = await window.ES_MODAL.form({
      title: '更新测试电流', icon: 'edit_note', confirmText: '追加记录',
      fields: [
        { key: 'date', label: '更新日期', value: last ? (last.date || '') : todayStr(), placeholder: 'YYYY-MM-DD' },
        { key: 'version', label: '测试版本号', value: last ? (last.version || '') : '', placeholder: '如：EVT1' },
        { key: 'current', label: '场景电流（mA）', value: last ? String(last.current) : '', placeholder: '不小于 0 的数字', inputmode: 'decimal' },
        { key: 'note', label: '备注 / 说明（可留空）', type: 'textarea', value: last ? (last.note || '') : '' },
      ],
      validate: (v) => {
        if (!isValidDateText((v.date || '').trim())) return '更新日期格式应为真实日期 YYYY-MM-DD';
        const c = num(v.current);
        if (c === null || c < 0) return '场景电流需为不小于 0 的数字';
        return null;
      },
    });
    if (values === null) return;
    scenario.records.push({
      id: uid('rec'), date: values.date.trim(), version: values.version.trim(),
      current: num(values.current), note: values.note.trim(),
    });
    scenario.records.sort((a, b) => String(a.date).localeCompare(String(b.date)));
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  // 单场景续航公式编辑器：每场景一条独立公式（默认留空），命名空间仅 电流 + 电池参数，无 A1 网格
  const SINGLE_LIFE_INSERTS = ['电流', '场景电流', '可用容量', 'Typ', 'Min', '比例', '门限', 'IF(', 'ROUND(', 'POWER(', 'SQRT('];
  function openSingleLifeEditor(project, scenario) {
    const overlay = document.createElement('div');
    overlay.className = 'es-formula-overlay';
    overlay.innerHTML = '<div class="es-formula-dialog panel rounded-3xl">'
      + '<div class="es-formula-head"><h4><span class="material-symbols-outlined">function</span>单场景续航 · <b>' + escapeHtml(scenario.name) + '</b></h4>'
      + '<button type="button" class="es-icon-btn es-icon-danger" data-es-sl-close><span class="material-symbols-outlined">close</span></button></div>'
      + '<p class="es-formula-sub">直接填数字，或以 <b>=</b> 开头编写公式；留空表示不计算。可命名引用 <b>电流</b>（本场景当前电流，mA）与电池参数（可用容量 / Typ / Min / 比例 / 门限 等）。常用写法：<b>=可用容量/电流</b>（容量 mAh ÷ 电流 mA = 小时）。</p>'
      + '<textarea class="es-formula-input" rows="3" autocomplete="off"></textarea>'
      + '<div class="es-formula-quick">' + SINGLE_LIFE_INSERTS.map((q) => '<button type="button" class="es-quick-chip" data-es-sl-insert="' + escapeHtml(q) + '">' + escapeHtml(q) + '</button>').join('') + '</div>'
      + '<p class="es-formula-trial">试算：<span class="es-formula-trial-value">—</span></p>'
      + '<div class="es-formula-foot">'
      + '<button type="button" class="template-btn" data-es-sl-clear title="清空公式，恢复为不计算">清空</button>'
      + '<div><button type="button" class="template-btn" data-es-sl-cancel>取消</button>'
      + '<button type="button" class="template-btn template-btn-primary" data-es-sl-save>保存</button></div></div>'
      + '</div>';
    document.body.appendChild(overlay);
    const textarea = overlay.querySelector('.es-formula-input');
    const trialValue = overlay.querySelector('.es-formula-trial-value');
    textarea.value = scenario.singleLife || '';

    function close() { overlay.remove(); document.removeEventListener('keydown', onKey); }
    function onKey(e) { if (e.key === 'Escape') close(); }
    document.addEventListener('keydown', onKey);

    function trial() {
      const text = textarea.value.trim();
      if (text === '') { trialValue.textContent = '不计算（空）'; trialValue.classList.remove('is-error'); return; }
      const trialScenario = JSON.parse(JSON.stringify(scenario));
      trialScenario.singleLife = text;
      const r = E.evaluateSingleLife(project, trialScenario);
      if (r.ok) {
        trialValue.textContent = r.empty ? '不计算（空）' : fmt(r.value, 4);
        trialValue.classList.remove('is-error');
      } else {
        trialValue.textContent = r.error;
        trialValue.classList.add('is-error');
      }
    }
    textarea.addEventListener('input', trial);
    trial();

    overlay.querySelector('[data-es-sl-close]').addEventListener('click', close);
    overlay.querySelector('[data-es-sl-cancel]').addEventListener('click', close);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
    overlay.querySelectorAll('[data-es-sl-insert]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const piece = btn.dataset.esSlInsert;
        const start = textarea.selectionStart || textarea.value.length;
        const end = textarea.selectionEnd || textarea.value.length;
        textarea.value = textarea.value.slice(0, start) + piece + textarea.value.slice(end);
        textarea.focus();
        textarea.selectionStart = textarea.selectionEnd = start + piece.length;
        trial();
      });
    });
    overlay.querySelector('[data-es-sl-clear]').addEventListener('click', () => { textarea.value = ''; trial(); });
    overlay.querySelector('[data-es-sl-save]').addEventListener('click', async () => {
      const text = textarea.value.trim();
      if (text.charAt(0) === '=') {
        const check = window.ES_FORMULA.checkSyntax(text.slice(1));
        if (!check.ok) { await window.ES_MODAL.alert({ title: '公式语法错误', icon: 'error', message: check.error }); return; }
      } else if (text !== '' && num(text) === null) {
        await window.ES_MODAL.alert({ title: '无法保存', icon: 'error', message: '请输入数字，或以 = 开头的公式' });
        return;
      }
      scenario.singleLife = text;
      E.store.save();
      close();
      window.ES_ENDURANCE_UI.render();
      window.ES_ENDURANCE_UI.refreshCompareQuiet && window.ES_ENDURANCE_UI.refreshCompareQuiet();
    });
  }

  // 历史记录查看：居中弹窗 + 表格展示
  async function openHistory(project, scenarioId) {
    const scenario = project.scenarios.find((s) => s.id === scenarioId);
    if (!scenario || !scenario.records.length) {
      await window.ES_MODAL.alert({ title: '电流更新历史', icon: 'history', message: '该场景还没有测试电流记录' });
      return;
    }
    const isCustom = num(scenario.customCurrent) !== null;
    const customBanner = isCustom
      ? '<p class="es-history-banner">当前生效为自定义电流 ' + fmt(num(scenario.customCurrent), 3)
        + ' mA，测试记录暂不参与计算。</p>'
      : '';
    const rows = scenario.records.slice().reverse().map((r, i) => {
      const activeTag = (i === 0 && !isCustom) ? '<em class="es-history-tag">当前生效</em>' : '';
      return '<tr><td>' + escapeHtml(r.date || '—') + '</td>'
        + '<td>' + escapeHtml(r.version || '未填版本') + '</td>'
        + '<td class="es-history-num">' + escapeHtml(fmt(num(r.current), 3)) + ' mA</td>'
        + '<td>' + escapeHtml(r.note || '') + activeTag + '</td></tr>';
    }).join('');
    const html = customBanner
      + '<table class="es-modal-table es-history-table"><thead><tr><th>更新日期</th><th>测试版本</th><th>场景电流</th><th>备注</th></tr></thead>'
      + '<tbody>' + rows + '</tbody></table>'
      + '<p class="es-modal-hint">共 ' + scenario.records.length + ' 条记录，按日期倒序排列</p>';
    await window.ES_MODAL.alert({
      title: '「' + scenario.name + '」电流更新历史',
      icon: 'history', html: html, confirmText: '关闭',
    });
  }

  window.ES_ENDURANCE_SCENARIO = { renderScenarioCard, bindScenarios };
})();

/* ================================================================
 * 第四部分：公式驱动续航模型（对比列 × 场景行网格，单元格公式可编辑）
 * ================================================================ */
(function () {
  'use strict';
  const E = window.ES_ENDURANCE;
  const { escapeHtml, num, fmt, uid, debounce } = E;
  const root = document.getElementById('enduranceStatusRoot');

  function activeModel(project) {
    return project.models.find((m) => m.id === project.activeModelId) || project.models[0] || null;
  }

  function descKey(desc) {
    return JSON.stringify(desc);
  }
  function parseDesc(encoded) {
    try { return JSON.parse(decodeURIComponent(encoded)); } catch (e) { return null; }
  }
  // data-es-fx 属性里存的是 HTML 实体编码后的 JSON，CSS 属性选择器不会解码 &quot;，
  // 因此不能用 querySelector 拼 JSON；改为遍历按钮、比对 dataset 解码后的真实值。
  function findFxCell(scope, desc) {
    const key = descKey(desc);
    let matched = null;
    scope.querySelectorAll('[data-es-fx]').forEach((btn) => {
      if (!matched && btn.dataset.esFx === key) matched = btn;
    });
    return matched ? matched.closest('.es-formula-cell') : null;
  }
  function getRaw(model, desc) {
    if (desc.kind === 'current' || desc.kind === 'note') {
      const row = model.rows.find((r) => r.id === desc.rowId);
      if (!row) return '';
      const v = desc.kind === 'current' ? row.currentRaw : row.note;
      return v === undefined || v === null ? '' : String(v);
    }
    if (desc.kind === 'footerLabel' || desc.kind === 'footerCurrent') {
      const bucket = desc.kind === 'footerLabel'
        ? (model.footer && model.footer.labels)
        : (model.footer && model.footer.current);
      const v = bucket ? bucket[desc.rowKey] : '';
      return v === undefined || v === null ? '' : String(v);
    }
    if (desc.kind === 'cell') {
      const row = model.rows.find((r) => r.id === desc.rowId);
      const cell = row && row.cells[desc.colId];
      return cell && cell[desc.t] !== undefined && cell[desc.t] !== null ? String(cell[desc.t]) : '';
    }
    return (model.foot[desc.colId] && model.foot[desc.colId][desc.t]) || '';
  }
  function setRaw(model, desc, raw) {
    if (desc.kind === 'current' || desc.kind === 'note') {
      const row = model.rows.find((r) => r.id === desc.rowId);
      if (!row) return;
      if (desc.kind === 'current') row.currentRaw = raw; else row.note = raw;
      return;
    }
    if (desc.kind === 'footerLabel' || desc.kind === 'footerCurrent') {
      if (!model.footer) model.footer = E.defaultModelFooter();
      const bucketKey = desc.kind === 'footerLabel' ? 'labels' : 'current';
      if (!model.footer[bucketKey]) model.footer[bucketKey] = E.defaultModelFooter()[bucketKey];
      model.footer[bucketKey][desc.rowKey] = raw;
      return;
    }
    if (desc.kind === 'cell') {
      const row = model.rows.find((r) => r.id === desc.rowId);
      if (!row) return;
      if (!row.cells[desc.colId]) row.cells[desc.colId] = { d: '0', m: E.DEFAULT_MAH_FORMULA };
      row.cells[desc.colId][desc.t] = raw;
    } else {
      if (!model.foot[desc.colId]) model.foot[desc.colId] = E.defaultFoot('normal');
      model.foot[desc.colId][desc.t] = raw;
    }
  }

  // ---------- 卡片骨架 ----------
  function renderModelCard(project) {
    const model = activeModel(project);
    const tabs = project.models.map((m) => {
      const active = model && m.id === model.id;
      return '<button type="button" class="es-model-tab' + (active ? ' is-active' : '') + '" data-es-model="' + escapeHtml(m.id) + '">'
        + '<span class="material-symbols-outlined">view_day</span>' + escapeHtml(m.name) + '</button>';
    }).join('');

    return '<section class="panel rounded-3xl es-card es-model-card">'
      + '<header class="es-card-head"><h3><span class="material-symbols-outlined">modeling</span>DOU续航模型</h3>'
      + '<div class="es-head-actions">'
      + '<button type="button" class="template-btn template-btn-sm" data-es-model-action="add"><span class="material-symbols-outlined text-sm">add</span>新建模型</button>'
      + '<button type="button" class="template-btn template-btn-sm" data-es-model-action="copy"><span class="material-symbols-outlined text-sm">content_copy</span>复制当前</button>'
      + '</div></header>'
      + '<div class="es-model-tabs" role="group" aria-label="续航模型切换">' + tabs + '</div>'
      + (model ? renderModelBody(project, model) : renderModelEmpty())
      + '</section>';
  }

  function renderModelEmpty() {
    return '<div class="es-model-empty"><p>当前项目还没有续航模型，点「新建模型」创建；每个项目可保存多套模型，互不影响。</p>'
      + '<button type="button" class="template-btn template-btn-primary" data-es-model-action="add"><span class="material-symbols-outlined text-sm">add</span>新建模型</button></div>';
  }

  // ---------- 网格渲染 ----------
  function resultHtml(ev, desc) {
    const r = ev.result(desc);
    if (!r.ok) return '<span class="es-cell-error" title="' + escapeHtml(r.error) + '">' + escapeHtml(r.error) + '</span>';
    if (r.empty) return '<span class="es-cell-empty">—</span>';
    return '<span class="es-cell-value">' + fmt(r.value, 2) + '</span>';
  }

  // 汇总区备注格：与数据行备注完全同一种形态（文本框 + 「备注/说明」占位），只接受纯文本、不提供 fx / 不解析公式
  function footerNoteCell(model, col, t) {
    const colIndex = model.columns.findIndex((c) => c.id === col.id);
    const rowKey = t === 'noteTotal' ? 'total' : (t === 'noteAvg' ? 'avg' : (t === 'noteTyp' ? 'douTyp' : 'douMin'));
    const addr = E.noteLetterFor(colIndex) + E.footerRowNumber(model, rowKey);
    const v = (model.foot[col.id] && model.foot[col.id][t]) || '';
    return '<td class="es-note-cell"><input type="text" class="es-note-input" data-es-foot-note="'
      + escapeHtml(col.id) + '" data-es-foot-note-field="' + escapeHtml(t) + '" value="' + escapeHtml(v)
      + '" placeholder="备注/说明" autocomplete="off" title="地址 ' + escapeHtml(addr) + '：纯文本备注（不解析公式）" /></td>';
  }

  // 汇总/数据通用格：默认一个物理列一个格；只有平均电流 / Typ / Min 结果格用 colspan=2 合并「时长+耗电」两列
  function formulaCellHtml(ev, model, desc, digits, colspan) {
    const addr = E.addressOf(model, desc);
    const raw = getRaw(model, desc);
    const r = ev.result(desc);
    const isFormula = String(raw).trim().charAt(0) === '=';
    const showTextError = !r.ok && (isFormula || r.error !== '#VALUE!');
    const spanAttr = colspan ? ' colspan="' + colspan + '"' : '';
    const mergeCls = colspan ? ' es-dou-merge' : '';
    const cls = 'es-formula-cell es-editable-cell' + mergeCls + (isFormula ? ' is-formula' : '') + (showTextError ? ' is-error' : '');
    // 汇总区备注列空值显示「备注/说明」轻提示，与数据行保持一致
    const hintText = (desc.kind === 'foot' && String(desc.t || '').indexOf('note') === 0) ? '备注/说明' : '—';
    const text = !r.ok ? (showTextError ? r.error : String(raw).trim()) : (r.empty ? hintText : fmt(r.value, digits === undefined ? 2 : digits));
    const addrText = addr ? '地址 ' + addr + '：' : '汇总展示格（无固定 A1 地址）：';
    const errHint = r.ok ? '' : (r.error === '#VALUE!'
      ? '（引用的单元格是文字或空值，或电池容量等参数不是数字）'
      : (r.error === '#DIV/0!' ? '（除数为 0：总时长或总耗电为 0）' : ''));
    return '<td' + spanAttr + ' class="' + cls + '" data-es-cell="' + escapeHtml(descKey(desc))
      + '" title="' + escapeHtml(addrText + errHint) + '点击直接编辑（数字 / 文字 / = 公式），fx 为高级试算编辑器">'
      + '<span class="es-cell-text' + (showTextError ? ' es-cell-text-error' : '') + '">' + escapeHtml(text) + '</span>'
      + '<button type="button" class="es-fx-btn" data-es-fx="' + escapeHtml(descKey(desc)) + '" title="在公式编辑器中打开 ' + escapeHtml(addr) + '">fx</button></td>';
  }

  // B 电流格（v52 全开放）：默认回退场景共享电流（灰色自动态）；填入数字 / 文字 / 公式后按本模型自由覆盖
  function currentCellHtml(ev, model, row, rowIndex) {
    const desc = { kind: 'current', rowId: row.id };
    const addr = 'B' + E.rowNumberFor(rowIndex);
    const raw = getRaw(model, desc);
    const trimmed = String(raw || '').trim();
    if (trimmed === '') {
      const r = ev.result(desc);
      const text = !r.ok ? r.error : (r.empty ? '—' : fmt(r.value, 2));
      return '<td class="es-current-cell es-formula-cell es-editable-cell es-avg-auto" data-es-cell="'
        + escapeHtml(descKey(desc)) + '" title="地址 ' + addr + '：当前取场景共享电流；点击可在本模型内改写为数字 / 文字 / = 公式，留空恢复场景电流">'
        + '<span class="es-cell-text es-avg-auto-text' + (!r.ok ? ' es-cell-text-error' : '') + '">' + escapeHtml(text) + '</span>'
        + '<button type="button" class="es-fx-btn" data-es-fx="' + escapeHtml(descKey(desc)) + '" title="编辑电流 ' + addr + '">fx</button></td>';
    }
    const r = ev.result(desc);
    const isFormula = trimmed.charAt(0) === '=';
    const showTextError = !r.ok && (isFormula || r.error !== '#VALUE!');
    const cls = 'es-current-cell es-formula-cell es-editable-cell' + (isFormula ? ' is-formula' : '') + (showTextError ? ' is-error' : '');
    const text = !r.ok ? (showTextError ? r.error : trimmed) : (r.empty ? '—' : fmt(r.value, 2));
    return '<td class="' + cls + '" data-es-cell="' + escapeHtml(descKey(desc))
      + '" title="地址 ' + addr + '：本模型自由电流（数字 / 文字 / = 公式），留空恢复场景共享电流">'
      + '<span class="es-cell-text' + (showTextError ? ' es-cell-text-error' : '') + '">' + escapeHtml(text) + '</span>'
      + '<button type="button" class="es-fx-btn" data-es-fx="' + escapeHtml(descKey(desc)) + '" title="编辑电流 ' + addr + '">fx</button></td>';
  }

  // 平均电流格：按需求把「时长 + 耗电」两列（C:D / F:G / I:J）合并成一格，A1 地址取左上角 C；foot.avg 留空时显示自动值
  function avgCellHtml(ev, model, col) {
    const desc = { kind: 'foot', colId: col.id, t: 'avg' };
    const raw = getRaw(model, desc);
    const trimmed = String(raw || '').trim();
    const addr = E.addressOf(model, desc);
    if (trimmed === '') {
      const auto = ev.stats[col.id] ? ev.stats[col.id].avgCurrent : null;
      const text = auto === null || auto === undefined ? '—' : fmt(auto, 2);
      return '<td colspan="2" class="es-stat-mah es-dou-merge es-formula-cell es-editable-cell es-avg-auto" data-es-cell="'
        + escapeHtml(descKey(desc)) + '" title="合并「时长 + 耗电」两列，地址 ' + escapeHtml(addr) + '；平均电流（自动 = 总耗电 ÷ 总时长），点击可改写为固定值或公式，留空恢复自动">'
        + '<span class="es-cell-text es-avg-auto-text">' + escapeHtml(text) + '</span>'
        + '<button type="button" class="es-fx-btn" data-es-fx="' + escapeHtml(descKey(desc)) + '" title="自定义平均电流">fx</button></td>';
    }
    const r = ev.result(desc);
    const isFormula = trimmed.charAt(0) === '=';
    const showTextError = !r.ok && (isFormula || r.error !== '#VALUE!');
    const cls = 'es-stat-mah es-dou-merge es-formula-cell es-editable-cell' + (isFormula ? ' is-formula' : '') + (showTextError ? ' is-error' : '');
    const text = !r.ok ? (showTextError ? r.error : trimmed) : (r.empty ? '—' : fmt(r.value, 2));
    return '<td colspan="2" class="' + cls + '" data-es-cell="' + escapeHtml(descKey(desc))
      + '" title="合并「时长 + 耗电」两列，地址 ' + escapeHtml(addr) + '；自定义平均电流：点击编辑（数字、文字或 = 公式），清空则恢复自动">'
      + '<span class="es-cell-text' + (showTextError ? ' es-cell-text-error' : '') + '">' + escapeHtml(text) + '</span>'
      + '<button type="button" class="es-fx-btn" data-es-fx="' + escapeHtml(descKey(desc)) + '" title="编辑平均电流">fx</button></td>';
  }

  function renderModelBody(project, model) {
    const ev = E.evaluateModel(project, model);
    const layout = E.loadGridLayout();
    const dataRowH = E.GRID_ROW_H[layout.rowHeight] || E.GRID_ROW_H[E.GRID_DEFAULT_ROWH];

    // 列宽 / 行高可自定义：宽度按角色内联到 col，行高经 CSS 变量下发；表格按内容收缩、不撑满容器
    const colEl = (role, cls) => '<col class="' + cls + '" data-col-role="' + role
      + '" style="width:' + layout.colWidths[role] + 'px; min-width:' + layout.colWidths[role] + 'px" />';

    // Excel 风格列标行（v3 连续地址，无空白列）：A=场景、B=电流；每个模型模块按时长 / 耗电 / 备注连续给字母
    const resizeHandle = (role) => '<span class="es-col-resize" data-col-role="' + role
      + '" title="拖拽调整列宽，双击恢复默认"></span>';
    const colBar = '<tr class="es-excel-colbar">'
      + '<th class="es-excel-rowhead"></th>'
      + '<th class="es-excel-colhead">' + resizeHandle('scenario') + 'A</th>'
      + '<th class="es-excel-colhead">' + resizeHandle('current') + 'B</th>'
      + model.columns.map((col, ci) => {
        const dLetter = E.colLetterFor(ci, 'd');
        const mLetter = E.colLetterFor(ci, 'm');
        const nLetter = E.noteLetterFor(ci);
        return '<th class="es-excel-colhead">' + resizeHandle('duration') + dLetter + '</th>'
          + '<th class="es-excel-colhead">' + resizeHandle('mah') + mLetter + '</th>'
          + '<th class="es-excel-colhead">' + resizeHandle('note') + nLetter + '</th>';
      }).join('')
      + '</tr>';

    // colgroup 精确控宽（table-layout:fixed 下列标行与数据行严格对齐）：
    // 行号 / 场景(A) / 电流(B) + 每个模型模块的 时长·耗电·备注；宽度来自布局偏好，用户拖拽后持久化
    const colGroup = '<colgroup>' + colEl('rownum', 'es-col-rownum') + colEl('scenario', 'es-col-scenario')
      + colEl('current', 'es-col-current')
      + model.columns.map(() => colEl('duration', 'es-col-duration') + colEl('mah', 'es-col-mah') + colEl('note', 'es-col-note')).join('')
      + '</colgroup>';
    // 表格显式宽度 = 各列宽之和：fixed 布局下严格按设定列宽渲染、不被内容撑宽，也不撑满容器
    const tableWidth = layout.colWidths.rownum + layout.colWidths.scenario + layout.colWidths.current
      + model.columns.length * (layout.colWidths.duration + layout.colWidths.mah + layout.colWidths.note);

    // 两行分组表头：行号 / 场景 / 电流 跨 2 行；每个模型模块 colspan=3（时长 / 耗电量 / 备注）
    const headTop = '<th class="es-corner es-excel-rowhead" rowspan="2" title="行号（公式地址中的行，如 C2）"></th><th class="es-corner" rowspan="2">场景</th><th rowspan="2">电流 mA</th>'
      + model.columns.map((col, ci) => {
        const tag = col.kind === 'full' ? '<em class="es-mode-tag">满电</em>' : '';
        const win = col.window ? '<small>' + escapeHtml(col.window) + '</small>' : '';
        const need = (col.needMah !== null && col.needMah !== undefined && col.needMah !== '')
          ? '<small>目标 ' + fmt(num(col.needMah), 0) + ' mAh</small>' : '';
        const note = col.note ? '<small class="es-mode-note">' + escapeHtml(col.note) + '</small>' : '';
        const ops = '<span class="es-mode-ops">'
          + '<button type="button" class="es-icon-btn" data-es-col-edit="' + escapeHtml(col.id) + '" title="编辑对比列"><span class="material-symbols-outlined">edit</span></button>'
          + '<button type="button" class="es-icon-btn es-icon-danger" data-es-col-delete="' + escapeHtml(col.id) + '" title="删除对比列"><span class="material-symbols-outlined">delete</span></button>'
          + '</span>';
        const dLetter = E.colLetterFor(ci, 'd');
        const mLetter = E.colLetterFor(ci, 'm');
        const nLetter = E.noteLetterFor(ci);
        return '<th class="es-mode-head" colspan="3"><div class="es-mode-title"><span class="es-colname-edit es-editable-cell" data-es-colname-edit="'
          + escapeHtml(col.id) + '" title="点击重命名该对比列">' + escapeHtml(col.name) + '</span>' + tag + ops + '</div>'
          + '<small class="es-col-addr">时长 ' + dLetter + ' / 耗电 ' + mLetter + ' / 备注/说明 ' + nLetter + '</small>' + win + need + note + '</th>';
      }).join('');
    const headSub = model.columns.map(() => '<th class="es-sub">时长 h</th><th class="es-sub">耗电量 mAh</th><th class="es-sub es-note-sub" title="本行备注 / 说明：自由文本或 = 公式，可被按 A1 地址引用">备注/说明</th>').join('');

    const body = model.rows.map((row, rowIndex) => {
      const scenario = project.scenarios.find((s) => s.id === row.scenarioId);
      const isBlankRow = !scenario;
      const scenarioName = scenario ? scenario.name : (String(row.label || '').trim() || '（未命名自由行）');
      const cells = model.columns.map((col) => {
        const cell = row.cells[col.id] || { d: '', m: E.DEFAULT_MAH_FORMULA };
        const dDesc = { kind: 'cell', rowId: row.id, colId: col.id, t: 'd' };
        const mDesc = { kind: 'cell', rowId: row.id, colId: col.id, t: 'm' };
        const nDesc = { kind: 'note', rowId: row.id, colId: col.id };
        const dAddr = E.addressOf(model, dDesc);
        const nAddr = E.addressOf(model, nDesc);
        const dResult = ev.result(dDesc);
        const dInvalid = !dResult.ok;
        const isFormulaD = String(cell.d || '').trim().charAt(0) === '=';
        const dInput = '<td class="es-duration-wrap' + (dInvalid ? ' is-error' : '') + (isFormulaD ? ' is-formula' : '') + '">'
          + '<input type="text" class="es-duration-input" data-es-row="' + escapeHtml(row.id) + '" data-es-col="' + escapeHtml(col.id) + '"'
          + ' value="' + escapeHtml(cell.d === undefined || cell.d === null ? '' : cell.d) + '"'
          + ' title="地址 ' + escapeHtml(dAddr) + '；可填数字 / 文字，或以 = 开头写公式" autocomplete="off" />'
          + '<small class="es-formula-result' + (dInvalid ? ' is-error' : '') + '"' + (isFormulaD ? '' : ' hidden') + '>'
          + (isFormulaD ? (dInvalid ? escapeHtml(dResult.error) : fmt(dResult.value, 2)) : '') + '</small>'
          + '</td>';
        const noteCell = '<td class="es-note-cell"><input type="text" class="es-note-input" data-es-row-note="'
          + escapeHtml(row.id) + '" value="' + escapeHtml(row.note || '') + '" placeholder="备注/说明" autocomplete="off"'
          + ' title="地址 ' + escapeHtml(nAddr) + '：自由文本或 = 公式，可被其它单元格按 A1 地址引用（文本参与数值运算时记 #VALUE!）" /></td>';
        return dInput + formulaCellHtml(ev, model, mDesc) + noteCell;
      }).join('');
      const rowHead = '<th class="es-excel-rownum' + (isBlankRow ? ' is-free-row' : '') + '" title="行号（公式地址中的行）"><span class="es-rownum-num">' + E.rowNumberFor(rowIndex) + '</span>'
        + '<button type="button" class="es-row-insert-btn" data-es-row-insert="' + escapeHtml(row.id)
        + '" title="在这一行上方插入空白自由行"><span class="material-symbols-outlined">add</span></button>'
        + '<button type="button" class="es-row-del-btn" data-es-row-delete="' + escapeHtml(row.id)
        + '" title="从当前模型删除该行（不删除场景及其电流数据）"><span class="material-symbols-outlined">close</span></button></th>';
      const nameCell = isBlankRow
        ? '<td class="es-sc-name es-editable-cell" data-es-label-edit="' + escapeHtml(row.id) + '" title="点击编辑自由行标题">' + escapeHtml(scenarioName) + '</td>'
        : '<td class="es-sc-name es-editable-cell" data-es-scenario-edit="' + escapeHtml(row.scenarioId)
          + '" title="点击重命名场景">' + escapeHtml(scenarioName) + '</td>';
      return '<tr data-es-model-row="' + escapeHtml(row.id) + '">' + rowHead
        + nameCell + currentCellHtml(ev, model, row, rowIndex) + cells + '</tr>';
    }).join('');

    // tfoot：Total 行逐列求和（C=时长 SUM、D=耗电 SUM、E=备注）；
    // 平均电流 / Typ / Min 三行按需求把「时长 + 耗电」两列合并（C:D / F:G / I:J），A1 地址取左上角 C，E/H/K 备注独立；
    // 四个汇总行的 A:B 也合并成一格，只展示 A 标签（v59 起汇总行不再有 B 列电流格）
    const footerLabel = (rowKey) => formulaCellHtml(ev, model, { kind: 'footerLabel', rowKey }, 2, 2);
    const totalRowHtml = '<tr class="es-stat-row"><th class="es-excel-rownum">' + E.totalRowNumber(model) + '</th>'
      + footerLabel('total')
      + model.columns.map((col) => formulaCellHtml(ev, model, { kind: 'foot', colId: col.id, t: 'td' }, 2)
        + formulaCellHtml(ev, model, { kind: 'foot', colId: col.id, t: 'tm' }, 2)
        + footerNoteCell(model, col, 'noteTotal')).join('') + '</tr>';

    const avgRow = '<tr class="es-stat-row es-stat-auto"><th class="es-excel-rownum">' + E.avgRowNumber(model) + '</th>'
      + footerLabel('avg')
      + model.columns.map((col) => avgCellHtml(ev, model, col)
        + footerNoteCell(model, col, 'noteAvg')).join('') + '</tr>';

    const douRow = (rowKey, valueType, noteType) => '<tr class="es-stat-row"><th class="es-excel-rownum">'
      + E.footerRowNumber(model, rowKey) + '</th>'
      + footerLabel(rowKey)
      + model.columns.map((col) => formulaCellHtml(ev, model, { kind: 'foot', colId: col.id, t: valueType }, 2, 2)
        + footerNoteCell(model, col, noteType)).join('') + '</tr>';

    const totals = totalRowHtml + avgRow
      + douRow('douTyp', 'douTyp', 'noteTyp')
      + douRow('douMin', 'douMin', 'noteMin');

    const rowHeightBtn = (key, label) => '<button type="button" class="es-mini-btn es-rowh-btn'
      + (layout.rowHeight === key ? ' is-on' : '') + '" data-es-row-h="' + key + '">' + label + '</button>';
    const rowHeightGroup = '<div class="es-rowh-group" role="group" aria-label="行高设置">'
      + '<span class="es-rowh-label">行高</span>'
      + rowHeightBtn('compact', '紧凑') + rowHeightBtn('normal', '标准') + rowHeightBtn('comfortable', '宽松')
      + '</div>';

    return '<div class="es-mode-toolbar">'
      + '<button type="button" class="es-mini-btn" data-es-model-action="add-row"><span class="material-symbols-outlined">playlist_add</span>新增场景行</button>'
      + '<button type="button" class="es-mini-btn" data-es-model-action="add-blank-row"><span class="material-symbols-outlined">post_add</span>插入空白行</button>'
      + '<button type="button" class="es-mini-btn" data-es-model-action="add-col"><span class="material-symbols-outlined">view_column</span>新增对比列</button>'
      + '<button type="button" class="es-mini-btn" data-es-model-action="rename"><span class="material-symbols-outlined">drive_file_rename_outline</span>重命名模型</button>'
      + '<button type="button" class="es-mini-btn" data-es-model-action="help"><span class="material-symbols-outlined">function</span>公式说明</button>'
      + rowHeightGroup
      + '<button type="button" class="es-mini-btn" data-es-model-action="reset-layout"><span class="material-symbols-outlined">restart_alt</span>重置列宽/行高</button>'
      + '<button type="button" class="es-mini-btn es-mini-danger" data-es-model-action="delete-model"><span class="material-symbols-outlined">delete</span>删除模型</button>'
      + '</div>'
      + '<div class="es-model-table-wrap"><table class="es-model-table es-excel-grid" data-row-height="' + layout.rowHeight
      + '" style="width:' + tableWidth + 'px; min-width:' + tableWidth + 'px; --es-data-row-h:' + dataRowH
      + 'px; --es-foot-row-h:' + (dataRowH + 4) + 'px">' + colGroup + '<thead>' + colBar + '<tr>' + headTop + '</tr><tr>' + headSub + '</tr></thead>'
      + '<tbody>' + (body || '<tr><td colspan="' + (3 + model.columns.length * 3) + '" class="es-table-empty">请先在上方「场景电流」中新增场景，或点工具栏「新增场景行」</td></tr>') + '</tbody>'
      + '<tfoot>' + totals + '</tfoot></table></div>'
      + '<p class="es-model-tip"><span class="material-symbols-outlined">info</span>'
      + '默认口径对齐云表：场景耗电 = 时长 × 场景电流（0 时长自动得 0）；可使用时长 = 可用容量（Typ × 可放电比例，不硬编码）÷ 总耗电 × 总时长。'
      + '<b>数据格、汇总区每个实际渲染格、场景/自由行标题与列名均可点按自由编辑</b>：数字、文字或 <b>=</b> 公式（回车 / 失焦保存、Esc 取消；文本格原样显示，参与数值计算时记 #VALUE!）；'
      + 'Total 行逐列求和；平均电流 / Typ / Min 三行按模块把「时长 + 耗电」两列合并（C:D / F:G / I:J），地址取合并区左上角 C；'
      + '<b>B 电流</b>可在本模型内改写为任意数字 / 文字 / 公式（留空回退场景共享电流，地址 B2、B3… 可被引用），<b>备注列</b>同样可被 A1 引用；'
      + '工具栏可「插入空白行」做与场景无关的自由行；自由行行号会常驻显示 + / ✕，可继续插入或直接删除；删除行 / 列时 A1 引用按 Excel 规则自动平移，指向被删区域的公式显式报错；<b>fx</b> 为带试算的高级编辑器。</p>';
  }

  // ---------- 就地刷新（保留输入焦点） ----------
  function updateModelCells(project, model) {
    const ev = E.evaluateModel(project, model);
    const tbody = root.querySelector('.es-model-table tbody');
    if (tbody) {
      tbody.querySelectorAll('tr[data-es-model-row]').forEach((tr) => {
        const rowId = tr.dataset.esModelRow;
        const row = model.rows.find((r) => r.id === rowId);
        if (!row) return;
        model.columns.forEach((col) => {
          const cDesc = { kind: 'current', rowId: rowId };
          const dDesc = { kind: 'cell', rowId: rowId, colId: col.id, t: 'd' };
          const mDesc = { kind: 'cell', rowId: rowId, colId: col.id, t: 'm' };
          const dResult = ev.result(dDesc);
          const currentCell = findFxCell(tr, cDesc);
          if (currentCell) paintFormulaCell(currentCell, ev.result(cDesc), getRaw(model, cDesc));
          const dWrap = tr.querySelector('.es-duration-input[data-es-col="' + col.id + '"]')?.closest('.es-duration-wrap');
          if (dWrap) {
            dWrap.classList.toggle('is-error', !dResult.ok);
            const rawNow = row.cells[col.id] ? String(row.cells[col.id].d || '') : '';
            const isFormula = rawNow.trim().charAt(0) === '=';
            dWrap.classList.toggle('is-formula', isFormula);
            const small = dWrap.querySelector('.es-formula-result');
            if (small) {
              small.hidden = !isFormula;
              small.classList.toggle('is-error', !dResult.ok);
              small.textContent = dResult.ok ? fmt(dResult.value, 2) : dResult.error;
            }
          }
          const mCell = findFxCell(tr, mDesc);
          if (mCell) paintFormulaCell(mCell, ev.result(mDesc), getRaw(model, mDesc));
        });
      });
    }
    const tfoot = root.querySelector('.es-model-table tfoot');
    if (tfoot) {
      ['td', 'tm', 'douTyp', 'douMin',
        'noteTotal', 'noteAvg', 'noteTyp', 'noteMin',
        'blankAvg', 'blankTyp', 'blankMin'].forEach((t) => {
        model.columns.forEach((col) => {
          const desc = { kind: 'foot', colId: col.id, t: t };
          const cell = findFxCell(tfoot, desc);
          if (cell) paintFormulaCell(cell, ev.result(desc), getRaw(model, desc));
        });
      });
      // 平均电流行：foot.avg 留空时只刷新自动值文字；自定义后按公式格结果重绘（不再 textContent 覆盖 fx 按钮）
      model.columns.forEach((col) => {
        const desc = { kind: 'foot', colId: col.id, t: 'avg' };
        const cell = findFxCell(tfoot, desc);
        if (!cell) return;
        const rawAvg = String(((model.foot[col.id] || {}).avg) || '').trim();
        if (rawAvg === '') {
          const auto = (ev.stats[col.id] || {}).avgCurrent;
          const autoText = cell.querySelector('.es-avg-auto-text');
          if (autoText) autoText.textContent = (auto === null || auto === undefined) ? '—' : fmt(auto, 2);
        } else {
          paintFormulaCell(cell, ev.result(desc), rawAvg);
        }
      });
      ['total', 'avg', 'douTyp', 'douMin'].forEach((rowKey) => {
        const labelDesc = { kind: 'footerLabel', rowKey };
        const currentDesc = { kind: 'footerCurrent', rowKey };
        const labelCell = findFxCell(tfoot, labelDesc);
        const currentCell = findFxCell(tfoot, currentDesc);
        if (labelCell) paintFormulaCell(labelCell, ev.result(labelDesc), getRaw(model, labelDesc));
        if (currentCell) paintFormulaCell(currentCell, ev.result(currentDesc), getRaw(model, currentDesc));
      });
    }
  }

  function paintFormulaCell(cell, result, raw) {
    const trimmed = String(raw || '').trim();
    const isFormula = trimmed.charAt(0) === '=';
    const showError = !result.ok && (trimmed === '' || isFormula || result.error !== '#VALUE!');
    cell.classList.toggle('is-error', showError);
    const text = cell.querySelector('.es-cell-text');
    if (!text) return;
    text.classList.toggle('es-cell-text-error', showError);
    text.textContent = !result.ok
      ? (showError ? result.error : trimmed)
      : (result.empty ? '—' : fmt(result.value, 2));
  }

  // ---------- 事件绑定 ----------
  // ---------- 列宽拖拽（Pointer Events，桌面 / 触屏通用；同角色所有列联动，偏好持久化） ----------
  function bindGridColumnResize() {
    const table = root.querySelector('.es-excel-grid');
    if (!table) return;
    const cols = Array.prototype.slice.call(table.querySelectorAll('colgroup col[data-col-role]'));
    const syncTableWidth = () => {
      const total = cols.reduce((sum, c) => sum + (c.getBoundingClientRect().width || 0), 0);
      if (total > 0) {
        table.style.width = Math.round(total) + 'px';
        table.style.minWidth = Math.round(total) + 'px';
      }
    };
    const applyWidth = (role, width) => {
      const w = E.clampGridColWidth(role, width);
      cols.forEach((c) => { if (c.dataset.colRole === role) c.style.width = w + 'px'; });
      syncTableWidth();
      return w;
    };
    table.querySelectorAll('.es-col-resize').forEach((handle) => {
      handle.addEventListener('pointerdown', (event) => {
        event.preventDefault();
        event.stopPropagation();
        const role = handle.dataset.colRole;
        const roleCols = cols.filter((c) => c.dataset.colRole === role);
        const startX = event.clientX;
        const startW = roleCols.length ? roleCols[0].getBoundingClientRect().width : E.loadGridLayout().colWidths[role];
        document.body.classList.add('es-col-resizing');
        const onMove = (ev) => applyWidth(role, startW + (ev.clientX - startX));
        const onUp = () => {
          document.removeEventListener('pointermove', onMove);
          document.removeEventListener('pointerup', onUp);
          document.body.classList.remove('es-col-resizing');
          const finalW = roleCols.length ? Math.round(roleCols[0].getBoundingClientRect().width) : startW;
          E.saveGridLayout({ colWidths: { [role]: E.clampGridColWidth(role, finalW) } });
        };
        document.addEventListener('pointermove', onMove);
        document.addEventListener('pointerup', onUp);
      });
      // 双击手柄：该角色恢复默认列宽
      handle.addEventListener('dblclick', (event) => {
        event.preventDefault();
        event.stopPropagation();
        const role = handle.dataset.colRole;
        const w = E.defaultGridLayout().colWidths[role];
        applyWidth(role, w);
        E.saveGridLayout({ colWidths: { [role]: w } });
      });
    });
  }

  function bindModels(project) {
    const quietRefreshCompare = debounce(() => {
      window.ES_ENDURANCE_UI.refreshCompareQuiet && window.ES_ENDURANCE_UI.refreshCompareQuiet();
    }, 220);

    root.querySelectorAll('[data-es-model]').forEach((btn) => {
      btn.addEventListener('click', () => {
        project.activeModelId = btn.dataset.esModel;
        E.store.save();
        window.ES_ENDURANCE_UI.render();
      });
    });
    root.querySelectorAll('[data-es-model-action]').forEach((btn) => {
      btn.addEventListener('click', () => handleModelAction(project, btn.dataset.esModelAction));
    });
    root.querySelectorAll('[data-es-col-edit]').forEach((btn) => {
      btn.addEventListener('click', () => openColumnEditor(project, activeModel(project), btn.dataset.esColEdit));
    });
    root.querySelectorAll('[data-es-col-delete]').forEach((btn) => {
      btn.addEventListener('click', () => deleteColumn(project, activeModel(project), btn.dataset.esColDelete));
    });

    // 行高三档切换（紧凑 / 标准 / 宽松），偏好持久化
    root.querySelectorAll('[data-es-row-h]').forEach((btn) => {
      btn.addEventListener('click', () => {
        E.saveGridLayout({ rowHeight: btn.dataset.esRowH });
        window.ES_ENDURANCE_UI.render();
      });
    });

    // 列宽拖拽 / 双击恢复默认
    bindGridColumnResize();

    // 汇总区备注格：纯文本，实时写回 foot 并保存（不解析公式、不整卡重绘以保留焦点）
    root.querySelectorAll('[data-es-foot-note]').forEach((input) => {
      input.addEventListener('input', () => {
        const model = activeModel(project);
        if (!model) return;
        const colId = input.dataset.esFootNote;
        const field = input.dataset.esFootNoteField;
        if (!model.foot[colId]) model.foot[colId] = E.defaultFoot('normal');
        model.foot[colId][field] = input.value;
        E.store.save();
        updateModelCells(project, model);
        quietRefreshCompare();
      });
    });

    // 备注/说明列：实时写回模型行并保存、不整卡重绘（保留输入焦点）；v52 起内容也可被其它单元格按 A1 地址引用。
    // 备注为场景行级、与具体模型无关，同一行在各模型模块内的备注框共享同一份 row.note，
    // 因此输入时同步同行其它备注框，避免一处编辑、其余框仍显示旧值。
    root.querySelectorAll('[data-es-row-note]').forEach((input) => {
      input.addEventListener('input', () => {
        const model = activeModel(project);
        if (!model) return;
        const row = model.rows.find((r) => r.id === input.dataset.esRowNote);
        if (!row) return;
        row.note = input.value;
        root.querySelectorAll('[data-es-row-note]').forEach((peer) => {
          if (peer !== input && peer.dataset.esRowNote === input.dataset.esRowNote) peer.value = input.value;
        });
        E.store.save();
        // 备注可作为 E/H/K 地址被公式引用：输入后立即重算依赖格与汇总，同时刷新跨项目对比
        updateModelCells(project, model);
        quietRefreshCompare();
      });
    });

    // fx 高级试算 / 单元格就地编辑 / 行列骨架改名 / 删除模型行：统一在下方 tableWrap 的 click 委托处理

    // 时长矩阵编辑（事件委托，就地更新结果与汇总，不整卡重绘以保留输入焦点）
    const tableWrap = root.querySelector('.es-model-table-wrap');
    if (tableWrap) {
      tableWrap.addEventListener('input', (event) => {
        const input = event.target.closest ? event.target.closest('.es-duration-input') : null;
        if (!input) return;
        const model = activeModel(project);
        if (!model) return;
        const row = model.rows.find((r) => r.id === input.dataset.esRow);
        if (!row) return;
        const wrap = input.closest('.es-duration-wrap');
        const rawText = input.value;
        const trimmed = rawText.trim();
        const isFormula = trimmed.charAt(0) === '=';
        // 自由编辑：数字 / 任意文本 / 公式都入库；文本或非法值的标红交给 updateModelCells 按求值结果统一刷新
        if (!row.cells[input.dataset.esCol]) row.cells[input.dataset.esCol] = { d: '', m: E.DEFAULT_MAH_FORMULA };
        row.cells[input.dataset.esCol].d = rawText;
        // 公式型时长格补充样式标记
        wrap.classList.toggle('is-formula', isFormula);
        E.store.save();
        updateModelCells(project, model);
        quietRefreshCompare();
      });

      // 统一 click 委托：fx 高级试算 / 数据格就地编辑 / 行列改名 / 删除模型行
      tableWrap.addEventListener('click', (event) => {
        const target = event.target;
        if (!target || !target.closest) return;
        const model = activeModel(project);
        if (!model) return;
        const fxBtn = target.closest('[data-es-fx]');
        if (fxBtn) {
          event.stopPropagation();
          const desc = parseDesc(fxBtn.dataset.esFx);
          if (desc) openFormulaEditor(project, model, desc);
          return;
        }
        const rowInsert = target.closest('[data-es-row-insert]');
        if (rowInsert) { event.stopPropagation(); insertBlankRowAbove(project, model, rowInsert.dataset.esRowInsert); return; }
        const rowDel = target.closest('[data-es-row-delete]');
        if (rowDel) { event.stopPropagation(); deleteScenarioRow(project, model, rowDel.dataset.esRowDelete); return; }
        const colName = target.closest('[data-es-colname-edit]');
        if (colName) { event.stopPropagation(); beginColNameEdit(project, model, colName.dataset.esColnameEdit, colName); return; }
        const scName = target.closest('[data-es-scenario-edit]');
        if (scName) { event.stopPropagation(); beginScenarioNameEdit(project, scName.dataset.esScenarioEdit, scName); return; }
        const labelCell = target.closest('[data-es-label-edit]');
        if (labelCell) { event.stopPropagation(); beginRowLabelEdit(model, labelCell.dataset.esLabelEdit, labelCell); return; }
        const cell = target.closest('[data-es-cell]');
        if (cell) {
          const desc = parseDesc(cell.dataset.esCell);
          if (desc) beginCellEdit(project, model, desc, cell);
        }
      });
    }
  }

  // ---------- 模型 / 列 / 行管理 ----------
  async function handleModelAction(project, action) {
    const model = activeModel(project);
    if (action === 'add') {
      const values = await window.ES_MODAL.form({
        title: '新建续航模型', icon: 'add_chart', confirmText: '新建',
        fields: [
          { key: 'projectName', label: '项目名称（必填）', value: project.name, placeholder: '模型归属项目；若不存在将自动新建该项目' },
          { key: 'modelName', label: '模型名称（选填）', value: '', placeholder: '留空则自动命名为「续航模型 N」' },
        ],
        validate: (v) => ((v.projectName || '').trim() ? null : '项目名称不能为空'),
      });
      if (values === null) return;
      const projectName = values.projectName.trim();
      const all = E.store.load();
      let target = all.projects.find((p) => p.name === projectName);
      if (!target) target = E.store.addProject(projectName); // 不存在则新建项目（自带 12 参考场景）并切换为当前
      const modelName = (values.modelName || '').trim() || ('续航模型 ' + (target.models.length + 1));
      const created = E.createModelV2(target, modelName, [{ name: '典型使用', kind: 'normal' }]);
      target.models.push(created);
      target.activeModelId = created.id;
      E.store.save();
      if (target.id !== project.id) E.store.setActive(target.id);
      window.ES_ENDURANCE_UI.render();
      return;
    }
    if (action === 'copy') {
      if (!model) return;
      const name = await window.ES_MODAL.prompt({
        title: '复制续航模型', icon: 'content_copy', label: '模型名称',
        value: model.name + '（副本）', confirmText: '复制',
        validate: (raw) => ((raw || '').trim() ? null : '模型名称不能为空'),
      });
      if (name === null) return;
      const copy = cloneModel(model, name.trim());
      project.models.push(copy);
      project.activeModelId = copy.id;
      E.store.save();
      window.ES_ENDURANCE_UI.render();
      return;
    }
    if (action === 'help') { openFormulaHelp(activeModel(project)); return; }
    if (action === 'reset-layout') {
      E.resetGridLayout();
      window.ES_ENDURANCE_UI.render();
      return;
    }
    if (!model) return;
    if (action === 'rename') {
      const name = await window.ES_MODAL.prompt({
        title: '重命名续航模型', icon: 'drive_file_rename_outline', label: '模型名称',
        value: model.name, confirmText: '保存',
        validate: (raw) => ((raw || '').trim() ? null : '模型名称不能为空'),
      });
      if (name === null) return;
      model.name = name.trim();
      E.store.save();
      window.ES_ENDURANCE_UI.render();
      return;
    }
    if (action === 'delete-model') {
      const ok = await window.ES_MODAL.confirm({
        title: '删除续航模型', icon: 'delete',
        message: '确认删除模型「' + model.name + '」？',
      });
      if (!ok) return;
      project.models = project.models.filter((m) => m.id !== model.id);
      project.activeModelId = project.models[0] ? project.models[0].id : null;
      E.store.save();
      window.ES_ENDURANCE_UI.render();
      return;
    }
    if (action === 'add-col') { await addColumn(project, model); return; }
    if (action === 'add-row') { await addScenarioRow(project, model); return; }
    if (action === 'add-blank-row') {
      appendBlankRow(model);
      E.store.save();
      window.ES_ENDURANCE_UI.render();
      return;
    }
  }

  function cloneModel(model, newName) {
    const copy = JSON.parse(JSON.stringify(model));
    copy.id = uid('m');
    copy.name = newName;
    const colIdMap = {};
    copy.columns.forEach((col) => {
      const oldId = col.id;
      col.id = uid('md');
      colIdMap[oldId] = col.id;
    });
    const newFoot = {};
    Object.keys(copy.foot || {}).forEach((oldId) => {
      newFoot[colIdMap[oldId]] = copy.foot[oldId];
    });
    copy.foot = newFoot;
    copy.rows = copy.rows.map((row) => {
      const cells = {};
      Object.keys(row.cells || {}).forEach((oldId) => { cells[colIdMap[oldId]] = row.cells[oldId]; });
      return { id: uid('mr'), scenarioId: row.scenarioId, label: row.label || '', currentRaw: row.currentRaw || '', note: row.note || '', cells: cells };
    });
    return copy;
  }

  async function addColumn(project, model) {
    const values = await window.ES_MODAL.form({
      title: '新增对比列', icon: 'view_column', confirmText: '新增列',
      fields: [
        { key: 'name', label: '对比列（使用模式）名称', placeholder: '如：中度使用' },
        {
          key: 'kind', label: '容量口径', type: 'select', value: 'normal',
          options: [
            { value: 'normal', label: '普通口径（可用容量 = Typ × 可放电比例）' },
            { value: 'full', label: '满电使用（按 Typ / Min 满容量计算）' },
          ],
        },
      ],
      validate: (v) => ((v.name || '').trim() ? null : '列名称不能为空'),
    });
    if (values === null) return;
    const col = E.createColumn(values.name.trim(), values.kind === 'full' ? 'full' : 'normal');
    model.columns.push(col);
    model.foot[col.id] = E.defaultFoot(col.kind);
    // 只给当前模型已有行补新模块；不能按 project.scenarios 回填，否则会把用户已从本模型删除的行重新带回来。
    // 场景行给默认 0 时长 + 默认耗电公式；自由空白行保持空白，交由用户自行填写。
    model.rows.forEach((row) => {
      row.cells[col.id] = row.scenarioId
        ? { d: '0', m: E.DEFAULT_MAH_FORMULA }
        : { d: '', m: '' };
    });
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  async function addScenarioRow(project, model) {
    const missing = project.scenarios.filter((s) => !model.rows.some((r) => r.scenarioId === s.id));
    const NEW = '__new__';
    const options = missing.map((s) => ({ value: s.id, label: s.name }));
    options.push({ value: NEW, label: '＋ 新建场景' });
    const values = await window.ES_MODAL.form({
      title: '模型新增场景行', icon: 'playlist_add', confirmText: '加入模型',
      fields: [{
        key: 'scenario', label: missing.length ? '选择要加入本模型的场景' : '项目场景已全部加入，只能新建',
        type: 'select', value: missing.length ? missing[0].id : NEW, options: options,
      }],
    });
    if (values === null) return;
    if (values.scenario === NEW) {
      const created = await createScenarioInteractively(project);
      if (!created) return;
      E.appendRowForScenario(model, created.id);
    } else {
      E.appendRowForScenario(model, values.scenario);
    }
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  // 自由空白行：先平移 A1 引用（含 Total / 平均 / DOU 汇总行），再插入，保证旧公式地址不失效。
  function appendBlankRow(model) {
    const insertIndex = model.rows.length;
    E.rewriteRefsAfterRowInsert(model, insertIndex);
    model.rows.push(E.createBlankRow(model));
  }

  function insertBlankRowAbove(project, model, rowId) {
    const insertIndex = model.rows.findIndex((r) => r.id === rowId);
    if (insertIndex < 0) return;
    E.rewriteRefsAfterRowInsert(model, insertIndex);
    model.rows.splice(insertIndex, 0, E.createBlankRow(model));
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  async function createScenarioInteractively(project) {
    const name = await window.ES_MODAL.prompt({
      title: '新建场景', icon: 'add', label: '场景名称',
      placeholder: '请输入场景名称', confirmText: '新建',
      validate: (raw) => ((raw || '').trim() ? null : '场景名称不能为空'),
    });
    if (name === null) return null;
    const created = {
      id: uid('sc'), name: name.trim(), remark: '', standard: '',
      customCurrent: '', fields: {}, records: [],
    };
    project.scenarios.push(created);
    return created;
  }

  async function openColumnEditor(project, model, colId) {
    if (!model) return;
    const col = model.columns.find((c) => c.id === colId);
    if (!col) return;
    const values = await window.ES_MODAL.form({
      title: '编辑对比列', icon: 'tune', confirmText: '保存',
      fields: [
        { key: 'name', label: '对比列名称', value: col.name },
        { key: 'window', label: '使用时间窗口（可留空）', value: col.window || '', placeholder: '如 8am~8pm' },
        { key: 'needMah', label: '目标耗电量 mAh（可留空）', value: col.needMah === null || col.needMah === undefined ? '' : String(col.needMah), placeholder: '不小于 0 的数字', inputmode: 'decimal' },
        { key: 'note', label: '备注（可留空）', type: 'textarea', value: col.note || '' },
        {
          key: 'kind', label: '容量口径', type: 'select', value: col.kind,
          options: [
            { value: 'normal', label: '普通口径（可用容量 = Typ × 可放电比例）' },
            { value: 'full', label: '满电使用（按 Typ / Min 满容量计算）' },
          ],
        },
      ],
      validate: (v) => {
        if (!(v.name || '').trim()) return '列名称不能为空';
        if ((v.needMah || '').trim() !== '' && (num(v.needMah) === null || num(v.needMah) < 0)) return '目标耗电量需为不小于 0 的数字';
        return null;
      },
    });
    if (values === null) return;
    col.name = values.name.trim();
    col.window = values.window.trim();
    col.needMah = values.needMah.trim() === '' ? null : num(values.needMah);
    col.note = values.note.trim();
    const newKind = values.kind === 'full' ? 'full' : 'normal';
    if (newKind !== col.kind) {
      col.kind = newKind;
      model.foot[col.id] = E.defaultFoot(newKind);
    }
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  async function deleteColumn(project, model, colId) {
    if (!model) return;
    const col = model.columns.find((c) => c.id === colId);
    if (!col) return;
    if (model.columns.length <= 1) {
      await window.ES_MODAL.alert({ title: '无法删除', icon: 'info', message: '每个续航模型至少保留一个对比列' });
      return;
    }
    const ok = await window.ES_MODAL.confirm({
      title: '删除对比列', icon: 'delete',
      message: '确认删除对比列「' + col.name + '」？该列下各场景的时长 / 耗电公式会一并删除。',
    });
    if (!ok) return;
    const moduleIndex = model.columns.findIndex((c) => c.id === colId);
    // Excel 式引用平移：指向被删模块的引用显式报错，右侧模块的物理列整体 -3
    if (moduleIndex >= 0) E.rewriteRefsAfterModuleDelete(model, moduleIndex);
    model.columns = model.columns.filter((c) => c.id !== colId);
    delete model.foot[colId];
    model.rows.forEach((row) => { delete row.cells[colId]; });
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  // ---------- Excel 风格就地编辑（click-to-edit）----------
  // 把任意 td / span 变成临时输入框：Enter 或失焦提交、Esc 取消；校验失败保持编辑并标红，不入库、不重绘。
  function beginInlineEdit(el, initial, options) {
    if (!el || el.querySelector('input, textarea')) return;
    el.classList.add('is-editing');
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'es-inline-edit-input ' + (options.inputClass || '');
    input.value = initial === null || initial === undefined ? '' : String(initial);
    input.autocomplete = 'off';
    el.innerHTML = '';
    el.appendChild(input);
    input.focus();
    input.select();
    let finished = false;
    const rerender = () => { finished = true; window.ES_ENDURANCE_UI.render(); };
    input.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
      else if (e.key === 'Escape') { e.preventDefault(); rerender(); }
    });
    input.addEventListener('blur', async () => {
      if (finished) return;
      const verdict = options.validate ? options.validate(input.value) : { ok: true, value: input.value };
      if (!verdict.ok) {
        input.classList.add('is-error');
        input.title = verdict.error || '输入不合法';
        input.focus();
        return; // 保持编辑态，等待用户修正
      }
      await options.commit(verdict.value);
      rerender();
    });
  }

  // 自由编辑格（耗电、Total、平均电流、DOU）：数字、任意文本或以 = 开头的公式均可；
  // 公式经安全引擎语法校验，文本直接入库（参与数值计算时由引擎记为 #VALUE!，与 Excel 一致）
  function beginCellEdit(project, model, desc, td) {
    beginInlineEdit(td, getRaw(model, desc), {
      inputClass: 'es-cell-edit-input',
      validate(text0) {
        const t = String(text0 || '').trim();
        if (t.charAt(0) === '=') {
          const check = window.ES_FORMULA.checkSyntax(t.slice(1));
          return check.ok ? { ok: true, value: t } : { ok: false, error: check.error };
        }
        // 数字或任意文本都允许；留空即清空该单元格
        return { ok: true, value: t };
      },
      commit(value) {
        setRaw(model, desc, value);
        E.store.save();
        if (typeof quietRefreshCompare === 'function') quietRefreshCompare();
      },
    });
  }

  // 场景名（A 列）就地重命名
  function beginScenarioNameEdit(project, scenarioId, td) {
    const scenario = project.scenarios.find((s) => s.id === scenarioId);
    if (!scenario) return;
    beginInlineEdit(td, scenario.name, {
      inputClass: 'es-name-edit-input',
      validate(v) {
        const t = String(v || '').trim();
        return t ? { ok: true, value: t } : { ok: false, error: '场景名称不能为空' };
      },
      commit(v) { scenario.name = v; E.store.save(); },
    });
  }

  // 自由行（无场景）A 列标题就地编辑：允许任意文字，也允许留空（渲染为「未命名自由行」）
  function beginRowLabelEdit(model, rowId, td) {
    const row = model.rows.find((r) => r.id === rowId);
    if (!row) return;
    beginInlineEdit(td, row.label || '', {
      inputClass: 'es-name-edit-input',
      validate(v) { return { ok: true, value: String(v || '').trim() }; },
      commit(v) { row.label = v; E.store.save(); },
    });
  }

  // 对比列名（列头）就地重命名；kind / window / 目标耗电等高级字段仍走列头「编辑」弹窗
  function beginColNameEdit(project, model, colId, span) {
    const col = model.columns.find((c) => c.id === colId);
    if (!col) return;
    beginInlineEdit(span, col.name, {
      inputClass: 'es-colname-edit-input',
      validate(v) {
        const t = String(v || '').trim();
        return t ? { ok: true, value: t } : { ok: false, error: '列名称不能为空' };
      },
      commit(v) { col.name = v; E.store.save(); },
    });
  }

  // 仅从当前模型删除一行（model.rows），不删除场景本身及其电流数据，其它模型不受影响
  async function deleteScenarioRow(project, model, rowId) {
    const rowIndex = model.rows.findIndex((r) => r.id === rowId);
    if (rowIndex < 0) return;
    const row = model.rows[rowIndex];
    const scenario = project.scenarios.find((s) => s.id === row.scenarioId);
    const label = scenario ? scenario.name : (String(row.label || '').trim() || '该自由行');
    const ok = await window.ES_MODAL.confirm({
      title: '删除模型行', icon: 'delete',
      message: '仅从当前模型「' + model.name + '」移除「' + label + '」这一行，不会删除场景及其电流数据，其它模型也不受影响。'
        + '指向该行的公式会变为报错，其下方行的 A1 引用会自动上移一行。确认删除？',
    });
    if (!ok) return;
    // Excel 式引用平移：先按删除前的行序号改写公式，再真正移除该行
    E.rewriteRefsAfterRowDelete(model, rowIndex);
    model.rows = model.rows.filter((r) => r.id !== rowId);
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  // ---------- 公式编辑器 ----------
  const QUICK_INSERTS = ['Typ', 'Min', '比例', '可用容量', '门限', '时长', '电流', '本列时长', '本列耗电', '总时长', '总耗电', 'SUM(', 'AVG(', 'IF(', 'ROUND('];

  function openFormulaEditor(project, model, desc) {
    const addr = E.addressOf(model, desc);
    const initial = getRaw(model, desc);
    const overlay = document.createElement('div');
    overlay.className = 'es-formula-overlay';
    overlay.innerHTML = '<div class="es-formula-dialog panel rounded-3xl">'
      + '<div class="es-formula-head"><h4><span class="material-symbols-outlined">function</span>编辑单元格 <b>' + escapeHtml(addr) + '</b></h4>'
      + '<button type="button" class="es-icon-btn es-icon-danger" data-es-fx-close><span class="material-symbols-outlined">close</span></button></div>'
      + '<p class="es-formula-sub">可填数字、任意文字，或以 <b>=</b> 开头编写公式；留空表示空单元格。文字格在参与数值计算时记为 #VALUE!。</p>'
      + '<textarea class="es-formula-input" rows="3" autocomplete="off"></textarea>'
      + '<div class="es-formula-quick">' + QUICK_INSERTS.map((q) => '<button type="button" class="es-quick-chip" data-es-insert="' + escapeHtml(q) + '">' + escapeHtml(q) + '</button>').join('') + '</div>'
      + '<p class="es-formula-trial">试算：<span class="es-formula-trial-value">—</span></p>'
      + '<div class="es-formula-foot">'
      + '<button type="button" class="template-btn" data-es-fx-reset title="恢复该单元格的默认公式">恢复默认</button>'
      + '<div><button type="button" class="template-btn" data-es-fx-cancel>取消</button>'
      + '<button type="button" class="template-btn template-btn-primary" data-es-fx-save>保存</button></div></div>'
      + '</div>';
    document.body.appendChild(overlay);
    const textarea = overlay.querySelector('.es-formula-input');
    const trialValue = overlay.querySelector('.es-formula-trial-value');
    textarea.value = initial;

    function close() { overlay.remove(); document.removeEventListener('keydown', onKey); }
    function onKey(e) { if (e.key === 'Escape') close(); }
    document.addEventListener('keydown', onKey);

    function trial() {
      const text = textarea.value.trim();
      if (text === '') { trialValue.textContent = '空单元格'; trialValue.classList.remove('is-error'); return; }
      const trialModel = JSON.parse(JSON.stringify(model));
      setRaw(trialModel, desc, text);
      const ev = E.evaluateModel(project, trialModel);
      const r = ev.result(desc);
      if (r.ok) {
        trialValue.textContent = r.empty ? '空单元格' : fmt(r.value, 4);
        trialValue.classList.remove('is-error');
      } else {
        trialValue.textContent = r.error;
        trialValue.classList.add('is-error');
      }
    }
    textarea.addEventListener('input', trial);
    trial();

    overlay.querySelector('[data-es-fx-close]').addEventListener('click', close);
    overlay.querySelector('[data-es-fx-cancel]').addEventListener('click', close);
    overlay.querySelectorAll('[data-es-insert]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const piece = btn.dataset.esInsert;
        const start = textarea.selectionStart || textarea.value.length;
        const end = textarea.selectionEnd || textarea.value.length;
        textarea.value = textarea.value.slice(0, start) + piece + textarea.value.slice(end);
        textarea.focus();
        textarea.selectionStart = textarea.selectionEnd = start + piece.length;
        trial();
      });
    });
    overlay.querySelector('[data-es-fx-reset]').addEventListener('click', () => {
      if (desc.kind === 'current' || desc.kind === 'note') {
        textarea.value = ''; // 自由电流 / 备注没有固定默认公式，恢复默认即清空（电流回退场景值）
      } else if (desc.kind === 'footerLabel' || desc.kind === 'footerCurrent') {
        const footerDefaults = E.defaultModelFooter();
        textarea.value = desc.kind === 'footerLabel'
          ? footerDefaults.labels[desc.rowKey]
          : footerDefaults.current[desc.rowKey];
      } else if (desc.kind === 'cell') {
        textarea.value = desc.t === 'm' ? E.DEFAULT_MAH_FORMULA : '0';
      } else {
        const col = model.columns.find((c) => c.id === desc.colId);
        textarea.value = E.defaultFoot(col ? col.kind : 'normal')[desc.t];
      }
      trial();
    });
    overlay.querySelector('[data-es-fx-save]').addEventListener('click', async () => {
      const text = textarea.value.trim();
      if (text.charAt(0) === '=') {
        const check = window.ES_FORMULA.checkSyntax(text.slice(1));
        if (!check.ok) {
          await window.ES_MODAL.alert({ title: '公式语法错误', icon: 'error', message: check.error });
          return;
        }
      }
      // 自由编辑：公式以外的数字或任意文本都允许保存（文本在数值计算时记为 #VALUE!）
      setRaw(model, desc, text);
      E.store.save();
      close();
      window.ES_ENDURANCE_UI.render();
      window.ES_ENDURANCE_UI.refreshCompareQuiet && window.ES_ENDURANCE_UI.refreshCompareQuiet();
    });
  }

  // ---------- 公式说明 ----------
  function openFormulaHelp(model) {
    const n = model ? model.rows.length : 0;
    const overlay = document.createElement('div');
    overlay.className = 'es-formula-overlay';
    overlay.innerHTML = '<div class="es-formula-dialog es-formula-help panel rounded-3xl">'
      + '<div class="es-formula-head"><h4><span class="material-symbols-outlined">function</span>公式语法说明</h4>'
      + '<button type="button" class="es-icon-btn es-icon-danger" data-es-help-close><span class="material-symbols-outlined">close</span></button></div>'
      + '<div class="es-help-body">'
      + '<h5>一、单元格地址（A1）</h5>'
      + '<p>A 列为场景 / 自由行标题（数据区按文本处理）；第 1 行表头，场景行从第 2 行开始'
      + (model ? '（本表共 ' + n + ' 个数据行，Total 在第 ' + E.totalRowNumber(model) + ' 行，平均电流在第 ' + E.avgRowNumber(model)
        + ' 行，Typ 在第 ' + E.douRowNumber(model) + ' 行，Min 在第 ' + E.douMinRowNumber(model) + ' 行）' : '') + '；'
      + 'B=电流（数据区与四个汇总行均可自由编辑，支持 B2 等裸地址引用）；每个对比列占连续三列「时长 / 耗电 / 备注」，'
      + '数据区 B 列留空时回退场景共享电流；第 1 个对比列为 C（时长）/ D（耗电）/ E（备注），第 2 个为 F / G / H，依次类推；备注列同样可被 A1 引用（文本参与数值运算时记 #VALUE!）。鼠标悬停单元格可见地址。'
      + '汇总区四个汇总行的 A:B 合并成一格、只展示 A 列标签（不再有 B 列电流格）；Total 行 C/F/I… 放本列时长 SUM、D/G/J… 放本列耗电 SUM、E/H/K… 放备注（备注列为「备注/说明」，与数据行同一种纯文本输入框，不解析公式）；'
      + '平均电流与 Typ/Min 三行按模块把「时长 + 耗电」两列合并（C:D / F:G / I:J），A1 地址取合并区左上角（如 C15、C16、C17），右侧 D 地址兼容指向同一结果；平均电流留空时界面显示自动值，公式引用该地址也取同一个自动值。</p>'
      + '<h5>二、电池参数（随电池卡片实时联动，不硬编码容量）</h5>'
      + '<p><b>Typ</b> 电池容量 Typ；<b>Min</b> 电池容量 Min；<b>比例</b> 可放电容量比例(0.2C)（空值按 1）；<b>门限</b> 低电量门限（空值按 0）；'
      + '<b>可用容量</b> = Typ × 比例（即云表 239 的来源）；<b>折减容量</b> = Typ × 比例 × (1 − 门限)；<b>关机电压 / Wh / 电压</b>。</p>'
      + '<h5>三、结构化名称（按所在行列自动绑定）</h5>'
      + '<p><b>时长</b> 本行本列时长；<b>电流</b> 本行场景当前电流（自定义电流优先）；<b>本列时长 / 本列耗电</b> 本列全部场景行（用于 SUM）；<b>总时长 / 总耗电</b> 本列 Total。</p>'
      + '<h5>四、函数与运算</h5>'
      + '<p>四则 <b>+ − * /</b>、幂 <b>^</b>、括号、比较 <b>&gt; &lt; &gt;= &lt;= = &lt;&gt;</b>；'
      + '函数 <b>SUM / AVG / MIN / MAX / ABS / ROUND / IF / POWER / SQRT</b>（支持中文别名 求和 / 平均值 / 如果 / 四舍五入 等）；区间如 <b>B2:B13</b> 只能用于聚合函数。</p>'
      + '<h5>五、默认公式（可改）</h5>'
      + '<p>场景耗电：<b>=IF(时长=0,0,电流*时长)</b>（时长为 0 时直接得 0，缺电流也不会污染汇总）；Total：<b>=SUM(本列时长)</b>、<b>=SUM(本列耗电)</b>；'
      + '普通列可使用时长：<b>=可用容量/总耗电*总时长</b>；满电列：<b>=Typ/总耗电*总时长</b>。</p>'
      + '<h5>六、错误码</h5>'
      + '<p>#REF! 引用越界 / 缺电流；#DIV/0! 除零；#CYCLE! 循环引用；#NAME? 未知名称；#VALUE! 数值类型错误；#SYNTAX! 语法错误。</p>'
      + '</div></div>';
    document.body.appendChild(overlay);
    overlay.querySelector('[data-es-help-close]').addEventListener('click', () => overlay.remove());
    overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.remove(); });
  }

  window.ES_ENDURANCE_MODEL = { renderModelCard, bindModels, activeModel };
})();

/* ================================================================
 * 第五部分：跨项目场景电流与 DOU 对比（公式模型求值结果驱动）
 * ================================================================ */
(function () {
  'use strict';
  const E = window.ES_ENDURANCE;
  const { escapeHtml, fmt, uid, debounce } = E;
  const root = document.getElementById('enduranceStatusRoot');

  // 仅会话状态：刷新后首次默认全量，之后保留当前筛选（含用户主动清空后的空选择）
  const compareState = {
    projects: [],
    scenarioNames: [],
    knownProjectIds: [],
    knownScenarioNames: [],
    seenProjects: false,
    seenScenarios: false,
    douMetric: 'typ',
  };

  function allScenarioNames(data) {
    const set = new Set();
    data.projects.forEach((p) => p.scenarios.forEach((s) => set.add(s.name)));
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'zh-Hans-CN'));
  }

  function currentOf(project, scenarioName) {
    const scenario = project.scenarios.find((s) => s.name === scenarioName);
    return scenario ? E.scenarioCurrent(project, scenario.id) : null;
  }

  // 找项目激活模型中与指定列名匹配的对比列（DOU 对比口径）
  function findColumn(project, columnName) {
    const model = project.models.find((m) => m.id === project.activeModelId) || project.models[0];
    if (!model) return null;
    return model.columns.find((c) => c.name === columnName) || null;
  }

  function activeColumnNames(data) {
    const set = new Set();
    data.projects.forEach((p) => {
      const model = p.models.find((m) => m.id === p.activeModelId) || p.models[0];
      if (model) model.columns.forEach((c) => set.add(c.name));
    });
    return Array.from(set);
  }

  function renderCompareCard(project) {
    const data = E.store.load();
    const currentProjectIds = data.projects.map((p) => p.id);
    if (!compareState.seenProjects) {
      // 首次渲染默认全选
      compareState.projects = currentProjectIds.slice();
      compareState.knownProjectIds = currentProjectIds.slice();
      compareState.seenProjects = true;
    } else {
      const known = new Set(compareState.knownProjectIds);
      const current = new Set(currentProjectIds);
      // 已删除项目移出选择；只有真正新增的项目自动纳入，用户主动清空 / 取消的存量项目不恢复
      compareState.projects = compareState.projects.filter((id) => current.has(id));
      currentProjectIds.forEach((id) => { if (!known.has(id) && !compareState.projects.includes(id)) compareState.projects.push(id); });
      compareState.knownProjectIds = currentProjectIds.slice();
    }
    const latestNames = allScenarioNames(data);
    if (!compareState.seenScenarios) {
      compareState.scenarioNames = latestNames.slice();
      compareState.knownScenarioNames = latestNames.slice();
      compareState.seenScenarios = true;
    } else {
      const latest = new Set(latestNames);
      const known = new Set(compareState.knownScenarioNames);
      // 已删除场景移出选择；只有真正新增的场景自动纳入，用户取消勾选的存量场景不恢复
      compareState.scenarioNames = compareState.scenarioNames.filter((n) => latest.has(n));
      latestNames.forEach((name) => { if (!known.has(name) && !compareState.scenarioNames.includes(name)) compareState.scenarioNames.push(name); });
      compareState.knownScenarioNames = latestNames.slice();
    }

    const selectedProjects = data.projects.filter((p) => compareState.projects.includes(p.id));
    const selectedNames = compareState.scenarioNames;
    const columnNames = activeColumnNames(data);
    const defaultColumn = columnNames[0] || '';

    const projectChecks = data.projects.map((p) => {
      const checked = compareState.projects.includes(p.id);
      return '<label class="es-pick-chip"><input type="checkbox" data-es-compare-project="' + escapeHtml(p.id) + '"'
        + (checked ? ' checked' : '') + '><span>' + escapeHtml(p.name) + '</span></label>';
    }).join('');

    // 始终渲染全部可选场景（按当前选择决定勾选态），避免取消勾选后 chip 消失、无法重新勾选
    const scenarioChecks = latestNames.map((name) =>
      '<label class="es-pick-chip"><input type="checkbox" data-es-compare-scenario="' + escapeHtml(name) + '"'
      + (selectedNames.includes(name) ? ' checked' : '') + '><span>' + escapeHtml(name) + '</span></label>').join('');

    return '<section class="panel rounded-3xl es-card es-compare-card">'
      + '<header class="es-card-head"><h3><span class="material-symbols-outlined">compare</span>跨项目对比</h3>'
      + '<div class="es-head-actions">'
      + '<button type="button" class="template-btn template-btn-sm" data-es-compare-all>全选项目</button>'
      + '<button type="button" class="template-btn template-btn-sm" data-es-compare-none>清空选择</button>'
      + '</div></header>'
      + '<div class="es-pick-block"><span class="es-pick-label">对比项目</span><div class="es-pick-chips">' + projectChecks + '</div></div>'
      + '<div class="es-pick-block"><span class="es-pick-label">对比场景</span><div class="es-pick-chips">'
      + (scenarioChecks || '<em class="es-pick-empty">暂无可对比场景</em>') + '</div></div>'
      + '<div class="es-pick-block"><span class="es-pick-label">DOU 口径</span>'
      + '<label class="es-pick-chip"><input type="radio" name="es-dou-metric" value="typ"' + (compareState.douMetric === 'typ' ? ' checked' : '') + '><span>Typ</span></label>'
      + '<label class="es-pick-chip"><input type="radio" name="es-dou-metric" value="min"' + (compareState.douMetric === 'min' ? ' checked' : '') + '><span>Min</span></label></div>'
      + renderModelSummary(selectedProjects, defaultColumn)
      + '<div class="es-chart-wrap"><h4>场景电流对比（mA）</h4><div class="es-chart-canvas-box"><canvas id="esCurrentChart"></canvas></div></div>'
      + '<div class="es-chart-wrap"><h4>可使用时长 DOU 对比（h，按各项目激活模型的同名对比列）</h4><div class="es-chart-canvas-box"><canvas id="esDouChart"></canvas></div></div>'
      + '</section>';
  }

  // 选中项目激活模型的逐列汇总
  function renderModelSummary(projects, defaultColumn) {
    const rows = projects.map((p) => {
      const model = p.models.find((m) => m.id === p.activeModelId) || p.models[0];
      if (!model) return '<tr><td>' + escapeHtml(p.name) + '</td><td colspan="5" class="es-table-empty">暂无模型</td></tr>';
      return model.columns.map((col) => {
        const stats = E.computeModeStats(p, model, col);
        const dou = col.kind === 'full'
          ? fmt(stats.fullHours, 2)
          : (compareState.douMetric === 'min' ? fmt(stats.douMin, 2) : fmt(stats.douTyp, 2));
        return '<tr><td>' + escapeHtml(p.name) + '</td><td>' + escapeHtml(col.name) + (col.kind === 'full' ? ' <em class="es-mode-tag">满电</em>' : '')
          + '</td><td>' + fmt(stats.totalHours, 2) + '</td><td>' + fmt(stats.totalMah, 2)
          + '</td><td>' + fmt(stats.avgCurrent, 2) + '</td><td>' + dou + '</td></tr>';
      }).join('');
    }).join('');
    return '<div class="es-table-wrap"><table class="es-table es-summary-table"><thead><tr>'
      + '<th>项目</th><th>激活模型对比列</th><th>总时长 h</th><th>总耗电量 mAh</th><th>平均电流 mA</th><th>可使用时长 h</th>'
      + '</tr></thead><tbody>' + (rows || '<tr><td colspan="6" class="es-table-empty">暂无数据</td></tr>') + '</tbody></table></div>';
  }

  function bindCompare(project) {
    const data = E.store.load();
    root.querySelectorAll('[data-es-compare-project]').forEach((cb) => {
      cb.addEventListener('change', () => {
        const id = cb.dataset.esCompareProject;
        compareState.projects = compareState.projects.filter((x) => x !== id);
        if (cb.checked) compareState.projects.push(id);
        redrawCharts();
      });
    });
    root.querySelectorAll('[data-es-compare-scenario]').forEach((cb) => {
      cb.addEventListener('change', () => {
        cb.closest('.es-pick-chips').querySelectorAll('[data-es-compare-scenario]').forEach((box) => {
          const name = box.dataset.esCompareScenario;
          compareState.scenarioNames = compareState.scenarioNames.filter((n) => n !== name);
          if (box.checked) compareState.scenarioNames.push(name);
        });
        drawCurrentChart();
      });
    });
    root.querySelector('[data-es-compare-all]').addEventListener('click', () => {
      compareState.projects = data.projects.map((p) => p.id);
      window.ES_ENDURANCE_UI.render();
    });
    root.querySelector('[data-es-compare-none]').addEventListener('click', () => {
      compareState.projects = [];
      window.ES_ENDURANCE_UI.render();
    });
    root.querySelectorAll('input[name="es-dou-metric"]').forEach((radio) => {
      radio.addEventListener('change', () => {
        compareState.douMetric = radio.value;
        window.ES_ENDURANCE_UI.render();
      });
    });
    setTimeout(drawCharts, 60);
  }

  function redrawCharts() {
    drawCurrentChart();
    drawDouChart();
  }

  function drawCharts() {
    drawCurrentChart();
    drawDouChart();
  }

  // 在同一 canvas 上重绘前，先销毁并摘除登记在册的旧 Chart 实例，避免 "Canvas is already in use" 与实例泄漏。
  // 同时清除可能已被响应式反馈循环撑爆的内联宽高（Chrome canvas 上限 2^25px），让新图在确定高度的父容器内重新测量。
  function reuseCanvas(canvas) {
    for (let i = E.chartInstances.length - 1; i >= 0; i -= 1) {
      const inst = E.chartInstances[i];
      if (inst.canvas === canvas || inst.ctx && inst.ctx.canvas === canvas) {
        try { inst.destroy(); } catch (e) { /* ignore */ }
        E.chartInstances.splice(i, 1);
      }
    }
    const box = canvas.parentElement;
    const boxH = box ? box.clientHeight : 0;
    // 父容器高度异常（0 或超过视口数倍）时不建图，等下一次 render 在正常布局下重建，杜绝暴涨
    if (!box || boxH <= 0 || boxH > window.innerHeight * 4) return false;
    canvas.style.width = '';
    canvas.style.height = '';
    return true;
  }

  function selectedProjects() {
    const data = E.store.load();
    return data.projects.filter((p) => compareState.projects.includes(p.id));
  }

  function drawCurrentChart() {
    const canvas = root.querySelector('#esCurrentChart');
    if (!canvas || typeof Chart === 'undefined') return;
    if (!reuseCanvas(canvas)) return;
    const projects = selectedProjects();
    const names = compareState.scenarioNames.slice();
    const datasets = projects.map((p, i) => ({
      label: p.name,
      data: names.map((name) => currentOf(p, name)),
      backgroundColor: E.CHART_COLORS[i % E.CHART_COLORS.length] + '33',
      borderColor: E.CHART_COLORS[i % E.CHART_COLORS.length],
      borderWidth: 1.5,
      borderRadius: 6,
    }));
    E.chartInstances.push(new Chart(canvas, {
      type: 'bar',
      data: { labels: names, datasets: datasets },
      options: {
        responsive: true, maintainAspectRatio: false, resizeDelay: 120,
        plugins: { legend: { labels: { boxWidth: 12, font: { size: 11 } } } },
        scales: {
          x: { ticks: { autoSkip: false, maxRotation: 40, minRotation: 30, font: { size: 10 } } },
          y: { beginAtZero: true, title: { display: true, text: 'mA' } },
        },
      },
    }));
  }

  function drawDouChart() {
    const canvas = root.querySelector('#esDouChart');
    if (!canvas || typeof Chart === 'undefined') return;
    if (!reuseCanvas(canvas)) return;
    const projects = selectedProjects();
    // 以各项目激活模型中出现过的对比列名为分组；同名对比列才可横向比较
    const columnNameSet = new Set();
    projects.forEach((p) => {
      const model = p.models.find((m) => m.id === p.activeModelId) || p.models[0];
      if (model) model.columns.forEach((c) => columnNameSet.add(c.name));
    });
    const columnNames = Array.from(columnNameSet);
    const datasets = projects.map((p, i) => ({
      label: p.name,
      data: columnNames.map((columnName) => {
        const model = p.models.find((m) => m.id === p.activeModelId) || p.models[0];
        const col = model && model.columns.find((c) => c.name === columnName);
        if (!col) return null;
        const stats = E.computeModeStats(p, model, col);
        if (col.kind === 'full') return stats.fullHours;
        return compareState.douMetric === 'min' ? stats.douMin : stats.douTyp;
      }),
      backgroundColor: E.CHART_COLORS[i % E.CHART_COLORS.length] + '33',
      borderColor: E.CHART_COLORS[i % E.CHART_COLORS.length],
      borderWidth: 1.5,
      borderRadius: 6,
    }));
    E.chartInstances.push(new Chart(canvas, {
      type: 'bar',
      data: { labels: columnNames, datasets: datasets },
      options: {
        responsive: true, maintainAspectRatio: false, resizeDelay: 120,
        plugins: { legend: { labels: { boxWidth: 12, font: { size: 11 } } } },
        scales: {
          x: { ticks: { autoSkip: false, maxRotation: 30, minRotation: 0, font: { size: 10 } } },
          y: { beginAtZero: true, title: { display: true, text: '小时 h' } },
        },
      },
    }));
  }

  // 数据变化后的静默重绘：保留筛选，仅重建对比卡与图表
  const refreshCompareQuiet = debounce(function () {
    const project = E.store.activeProject();
    if (!project) return;
    const card = root.querySelector('.es-compare-card');
    if (!card) return;
    // 旧卡片即将被替换，先销毁其 canvas 上的图表实例，避免脱离 DOM 的实例泄漏
    E.destroyCharts();
    const next = document.createElement('div');
    next.innerHTML = renderCompareCard(project);
    card.replaceWith(next.firstElementChild);
    bindCompare(project);
  }, 240);

  Object.assign(window.ES_ENDURANCE_UI, {
    bindCompare,
    renderCompareCard,
    refreshCompareQuiet,
  });
})();

/* ================================================================
 * 第六部分：启动引导。点击「续航/功耗现状」入口、页面可见后再首次渲染，
 * 避免在 display:none 容器里初始化 Chart.js 导致尺寸为 0。
 * ================================================================ */
(function () {
  'use strict';

  function boot() {
    const root = document.getElementById('enduranceStatusRoot');
    if (!root || root.dataset.esReady === '1') return;
    root.dataset.esReady = '1';
    // 预热本地数据（首次访问写入示例项目）
    window.ES_ENDURANCE.store.load();
    document.querySelectorAll('[data-feature-target="enduranceStatusPage"]').forEach((button) => {
      button.addEventListener('click', () => {
        // openFeaturePage 在同一点击事件里执行，下一帧容器已可见
        setTimeout(() => {
          try { window.ES_ENDURANCE_UI.render(); } catch (error) { console.error('续航现状模块渲染失败', error); }
        }, 60);
      });
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
})();
