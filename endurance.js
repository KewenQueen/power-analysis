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
  function defaultBattery() {
    return {
      capTyp: '', capMin: '', wh: '', voltage: '',
      shutdownVoltage: '',
      dischargeRatio: '', lowThreshold: '',
      // 电池可放电容量查询表：[{ id, voltage, capacityMah }]，按电压降序；
      // capacityMah = 电池从满电放电至该电压时累计可放出的容量（mAh）
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
      douTyp: kind === 'full' ? '=Typ/总耗电*总时长' : '=可用容量/总耗电*总时长',
      douMin: kind === 'full' ? '=Min/总耗电*总时长' : '=Min*比例/总耗电*总时长',
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
      return { id: uid('mr'), scenarioId: scenario.id, cells: cells };
    });
    return { v: 2, id: uid('m'), name: name, columns: columns, rows: rows, foot: foot };
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
      return { id: uid('mr'), scenarioId: oldRow.scenarioId, cells: cells };
    });
    return { v: 2, id: model.id, name: model.name, columns: columns, rows: rows, foot: foot };
  }

  // 首次使用的示例项目：对齐云表「G2 标准版续航模型 V1.0（UT 阶段）」
  function seedData() {
    // [场景名, 场景电流 mA, UT原始模型(3H) 时长, 环境记忆常开 时长, 重度模型 时长, 备注]
    // 时长取自云表 HrV2LC「续航模型」E/G/I 列，保留 6 位小数以贴近原表汇总
    const scenarioDefs = [
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
    const scenarios = scenarioDefs.map((def) => ({
      id: uid('sc'),
      name: def[0],
      remark: def[5],
      standard: '',
      customCurrent: '',
      fields: {},
      records: [{ id: uid('rec'), date: todayStr(), version: 'UT', current: def[1], note: '示例数据' }],
    }));
    const byName = {};
    scenarios.forEach((s) => { byName[s.name] = s.id; });

    const project = {
      id: uid('p'),
      name: '示例项目（可删除）',
      createdAt: Date.now(),
      scenarioFields: [],
      battery: Object.assign(defaultBattery(), {
        capTyp: 246, capMin: 240, wh: 0.95, voltage: 3.87,
        shutdownVoltage: 3.4,
        dischargeRatio: 0.95, lowThreshold: 0.02,
        // 示例查询表：关机电压 3.4V 对应可放电 233.7mAh ≈ 246 × 0.95
        dischargeTable: [
          { id: uid('dt'), voltage: 4.2, capacityMah: 0 },
          { id: uid('dt'), voltage: 4.0, capacityMah: 30 },
          { id: uid('dt'), voltage: 3.8, capacityMah: 90 },
          { id: uid('dt'), voltage: 3.6, capacityMah: 170 },
          { id: uid('dt'), voltage: 3.4, capacityMah: 233.7 },
          { id: uid('dt'), voltage: 3.2, capacityMah: 243 },
          { id: uid('dt'), voltage: 3.0, capacityMah: 246 },
        ],
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
      if (!('shutdownVoltage' in b)) { b.shutdownVoltage = ''; changed = true; }
      if (!Array.isArray(b.dischargeTable)) { b.dischargeTable = []; changed = true; }
      if (!Array.isArray(project.scenarioFields)) { project.scenarioFields = []; changed = true; }
      (project.scenarios || []).forEach((scenario) => {
        if (!('standard' in scenario)) { scenario.standard = ''; changed = true; }
        if (!('customCurrent' in scenario)) { scenario.customCurrent = ''; changed = true; }
        if (!scenario.fields || typeof scenario.fields !== 'object') { scenario.fields = {}; changed = true; }
        if (!Array.isArray(scenario.records)) scenario.records = [];
      });
      if (Array.isArray(project.models)) {
        project.models = project.models.map((model) => {
          const migrated = model.v === 2 ? model : migrateModelV1(model);
          // 一次性迁移：旧默认耗电公式升级为带 0 时长保护的 IF 公式（用户自定义公式保持不动）
          (migrated.rows || []).forEach((row) => {
            Object.keys(row.cells || {}).forEach((colId) => {
              if (row.cells[colId] && row.cells[colId].m === LEGACY_MAH_FORMULA) {
                row.cells[colId].m = DEFAULT_MAH_FORMULA;
                changed = true;
              }
            });
          });
          if (migrated.v !== 2) changed = true;
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
    const table = sortDischargeTable(b.dischargeTable);
    if (!table.length) return { ok: false, reason: 'empty-table' };
    const shutdown = num(b.shutdownVoltage);
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
        alert('保存失败：浏览器存储空间可能已满，请删除部分项目或历史记录后重试。');
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
        battery: defaultBattery(),
        scenarios: [],
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
  };

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

  // ---------- A1 网格映射 ----------
  // A 列=场景名；第 i 个对比列（0-based）占 时长列(2+2i) / 耗电列(3+2i)；
  // 第 1 行表头，场景行从第 2 行开始，其后依次为 Total 行、DOU 行。
  function colLetterFor(colIndex, cellType) {
    return window.ES_FORMULA.colToLetter(2 + colIndex * 2 + (cellType === 'm' ? 1 : 0));
  }
  function rowNumberFor(rowIndex) { return 2 + rowIndex; }
  function totalRowNumber(model) { return 2 + model.rows.length; }
  function douRowNumber(model) { return 3 + model.rows.length; }
  function addressOf(model, desc) {
    const colIndex = model.columns.findIndex((c) => c.id === desc.colId);
    if (colIndex < 0) return '';
    if (desc.kind === 'cell') {
      const rowIndex = model.rows.findIndex((r) => r.id === desc.rowId);
      return colLetterFor(colIndex, desc.t) + rowNumberFor(rowIndex);
    }
    const row = desc.t === 'td' || desc.t === 'tm' ? totalRowNumber(model) : douRowNumber(model);
    return colLetterFor(colIndex, desc.t === 'td' || desc.t === 'douTyp' ? 'd' : 'm') + row;
  }

  // ---------- 公式网格求值 ----------
  // 返回 { result(desc), stats, errorOf(desc) }；一次求值内带 memo 与循环检测
  function evaluateModel(project, model) {
    const F = window.ES_FORMULA;
    const ET = F.ERROR_TYPES;
    const cache = new Map();
    const n = model.rows.length;
    const totalRow = totalRowNumber(model);
    const douRow = douRowNumber(model);

    function keyOf(desc) {
      return desc.kind + '|' + (desc.rowId || '') + '|' + desc.colId + '|' + desc.t;
    }

    // A1 地址 -> 单元格描述
    function refToDesc(ref) {
      if (ref.col < 2) return null;
      const ci = Math.floor((ref.col - 2) / 2);
      const isM = (ref.col - 2) % 2 === 1;
      if (ci >= model.columns.length) return null;
      const colId = model.columns[ci].id;
      if (ref.row >= 2 && ref.row <= 1 + n) {
        const row = model.rows[ref.row - 2];
        return { kind: 'cell', rowId: row.id, colId: colId, t: isM ? 'm' : 'd' };
      }
      if (ref.row === totalRow) return { kind: 'foot', colId: colId, t: isM ? 'tm' : 'td' };
      if (ref.row === douRow) return { kind: 'foot', colId: colId, t: isM ? 'douMin' : 'douTyp' };
      return null;
    }

    function batteryNames() {
      const b = project.battery || {};
      const capTyp = num(b.capTyp);
      const capMin = num(b.capMin);
      const ratioRaw = num(b.dischargeRatio);
      const thresholdRaw = num(b.lowThreshold);
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
        '关机电压': need(num(b.shutdownVoltage)),
        'Wh': need(num(b.wh)), '电压': need(num(b.voltage)),
      };
    }

    function evalCell(desc, visiting) {
      const key = keyOf(desc);
      if (cache.has(key)) return cache.get(key);
      if (visiting[key]) return { ok: false, error: ET.CYCLE };
      visiting[key] = true;

      let raw = '';
      let scopeRow = null;
      if (desc.kind === 'cell') {
        const row = model.rows.find((r) => r.id === desc.rowId);
        const cell = row && row.cells[desc.colId];
        raw = cell ? (cell[desc.t] === undefined || cell[desc.t] === null ? '' : cell[desc.t]) : '';
        scopeRow = row;
      } else {
        raw = (model.foot[desc.colId] || {})[desc.t] || '';
      }

      const ctx = makeCtx(desc, scopeRow, visiting);
      const result = evalRaw(raw, ctx);
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
      const colIndex = model.columns.findIndex((c) => c.id === desc.colId);
      const names = batteryNames();

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
            if (desc.kind !== 'cell' || !scopeRow) return { type: ET.REF };
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
    model.rows.forEach((row) => {
      model.columns.forEach((col) => {
        evalCell({ kind: 'cell', rowId: row.id, colId: col.id, t: 'd' }, {});
        evalCell({ kind: 'cell', rowId: row.id, colId: col.id, t: 'm' }, {});
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
    const ratio = num(b.dischargeRatio);
    const threshold = num(b.lowThreshold);
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
    computeModeStats, computeModelStats, evaluateModel,
    sortDischargeTable, lookupDischargeCapacity, deriveDischargeRatio,
    DEFAULT_MAH_FORMULA, defaultFoot, createColumn, createModelV2, migrateModelV1,
    colLetterFor, rowNumberFor, totalRowNumber, douRowNumber, addressOf,
  };
})();

/* ================================================================
 * 第二部分：页面骨架、项目管理、电池参数、可放电容量查询表
 * ================================================================ */
(function () {
  'use strict';
  const E = window.ES_ENDURANCE;
  const { escapeHtml, uid, num, fmt, debounce } = E;

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
      window.ES_ENDURANCE_SCENARIO.bindScenarios(project);
      window.ES_ENDURANCE_MODEL.bindModels(project);
      window.ES_ENDURANCE_UI.bindCompare(project);
    }
  }

  function renderEmpty() {
    return '<div class="panel rounded-3xl es-empty">'
      + '<span class="material-symbols-outlined es-empty-icon">battery_saver</span>'
      + '<h3>还没有项目</h3>'
      + '<p>新增一个项目后即可维护电池参数、场景电流与自定义续航模型，并进行多项目横向对比。</p>'
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
      + (project ? '<button type="button" class="template-btn es-rename-project-btn"><span class="material-symbols-outlined text-sm">drive_file_rename_outline</span>重命名</button>'
        + '<button type="button" class="template-btn template-btn-danger es-delete-project-btn"><span class="material-symbols-outlined text-sm">delete</span>删除项目</button>' : '')
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
      btn.addEventListener('click', () => {
        const name = window.prompt('请输入新项目名称：', '新项目 ' + (data.projects.length + 1));
        if (name === null) return;
        const trimmed = name.trim();
        if (!trimmed) { alert('项目名称不能为空'); return; }
        E.store.addProject(trimmed);
        render();
      });
    });
    const renameBtn = root.querySelector('.es-rename-project-btn');
    if (renameBtn && project) {
      renameBtn.addEventListener('click', () => {
        const name = window.prompt('修改项目名称：', project.name);
        if (name === null) return;
        const trimmed = name.trim();
        if (!trimmed) { alert('项目名称不能为空'); return; }
        project.name = trimmed;
        E.store.save();
        render();
      });
    }
    const deleteBtn = root.querySelector('.es-delete-project-btn');
    if (deleteBtn && project) {
      deleteBtn.addEventListener('click', () => {
        if (!window.confirm('确认删除项目「' + project.name + '」？\n该项目下的电池参数、场景电流与全部续航模型都会被删除，且不可恢复。')) return;
        E.store.removeProject(project.id);
        render();
      });
    }
  }

  // ---------- 项目看板骨架 ----------
  function renderBoard(project) {
    return '<div class="es-board">'
      + renderBatteryCard(project)
      + renderDischargeCard(project)
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
    { key: 'shutdownVoltage', label: '关机电压 V', placeholder: '示例：3.4' },
    { key: 'dischargeRatio', label: '可放电容量比例(0.2C)', placeholder: '0~1，示例：0.95（可由查询表自动回填）' },
    { key: 'lowThreshold', label: '低电量门限', placeholder: '0~1，示例：0.02' },
  ];

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
    const ratio = num(b.dischargeRatio);
    const threshold = num(b.lowThreshold);
    const derate = (ratio === null ? 1 : ratio) * (1 - (threshold === null ? 0 : threshold));
    const capTyp = num(b.capTyp);
    const available = capTyp === null ? null : capTyp * (ratio === null ? 1 : ratio);
    return '<section class="panel rounded-3xl es-card">'
      + '<header class="es-card-head"><h3><span class="material-symbols-outlined">battery_full</span>电池参数</h3>'
      + '<small>续航公式基于这些参数自动计算，修改后实时保存</small></header>'
      + '<div class="es-battery-grid">' + fields + '</div>'
      + '<p class="es-battery-derate">可用容量（公式名「可用容量」= Typ × 可放电比例）= <b>' + fmt(available, 2)
      + ' mAh</b>；综合可用电量系数 = 可放电比例 × (1 − 低电量门限) = <b>' + fmt(derate, 4) + '</b></p>'
      + '</section>';
  }

  const saveBattery = debounce(() => E.store.save(), 150);

  function bindBattery(project) {
    root.querySelectorAll('[data-es-battery]').forEach((input) => {
      input.addEventListener('input', () => {
        const key = input.dataset.esBattery;
        const raw = input.value.trim();
        const parsed = raw === '' ? '' : num(raw);
        // 非数字、负数不允许写入；比例 / 门限还必须落在 0~1
        let invalid = raw !== '' && (parsed === null || parsed < 0);
        if (!invalid && (key === 'dischargeRatio' || key === 'lowThreshold') && raw !== '' && parsed > 1) {
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
        if (derateEl) {
          const ratio = num(project.battery.dischargeRatio);
          const threshold = num(project.battery.lowThreshold);
          const typNow = num(project.battery.capTyp);
          const derate = (ratio === null ? 1 : ratio) * (1 - (threshold === null ? 0 : threshold));
          const avail = typNow === null ? null : typNow * (ratio === null ? 1 : ratio);
          derateEl.innerHTML = '可用容量（公式名「可用容量」= Typ × 可放电比例）= <b>' + fmt(avail, 2)
            + ' mAh</b>；综合可用电量系数 = 可放电比例 × (1 − 低电量门限) = <b>' + fmt(derate, 4) + '</b>';
        }
        // 关机电压 / Typ 容量变化时，查询表联动重算可放电容量比例
        if (key === 'shutdownVoltage' || key === 'capTyp') {
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

  // ---------- 电池可放电容量查询表 ----------
  // 依据关机电压 + 查询表自动回填可放电容量比例(0.2C)；只更新输入框值，不整卡重绘以避免丢焦点。
  // 防抖回调内重新取当前激活项目，避免快速切换项目后 pending 闭包把旧项目数据写进新 DOM。
  const syncRatioFromTable = debounce(function () {
    const project = E.store.activeProject();
    if (!project) return;
    const derived = E.deriveDischargeRatio(project);
    const syncEl = root.querySelector('.es-discharge-sync');
    const b = project.battery || {};
    // Typ 容量变化时，就地刷新查询表「占 Typ 比例」列
    const capTypNow = num(b.capTyp);
    root.querySelectorAll('[data-es-dt-row]').forEach((tr) => {
      const row = (b.dischargeTable || []).find((r) => r.id === tr.dataset.esDtRow);
      const cell = tr.querySelector('.es-dt-ratio');
      if (!row || !cell) return;
      const c = num(row.capacityMah);
      cell.textContent = (capTypNow && capTypNow > 0 && c !== null) ? fmt(c / capTypNow * 100, 2) + '%' : '—';
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
      const threshold = num(b.lowThreshold);
      const derateEl = root.querySelector('.es-battery-derate');
      if (derateEl) {
        const capTyp = num(b.capTyp);
        const avail = capTyp === null ? null : capTyp * rounded;
        derateEl.innerHTML = '可用容量（公式名「可用容量」= Typ × 可放电比例）= <b>' + fmt(avail, 2)
          + ' mAh</b>；综合可用电量系数 = 可放电比例 × (1 − 低电量门限) = <b>'
          + fmt(rounded * (1 - (threshold === null ? 0 : threshold)), 4) + '</b>';
      }
      refreshModelAndCompare();
    }
    if (syncEl) syncEl.innerHTML = dischargeSyncHint(project, derived.ratio === null ? 'no-cap-typ' : '', derived);
  }, 180);

  function dischargeSyncHint(project, reason, derived) {
    const b = project.battery || {};
    const table = E.sortDischargeTable(b.dischargeTable);
    const base = '<span class="material-symbols-outlined">';
    if (!table.length) {
      return base + 'info</span>尚未维护查询表：可导入电池规格书文档自动识别，或点「新增一行」手动维护电压 - 容量点。';
    }
    if (reason === 'no-shutdown-voltage') {
      return base + 'hourglass_top</span>已维护 ' + table.length + ' 个电压点；填写「关机电压」后将自动插值计算可放电容量比例(0.2C)。';
    }
    const d = derived || E.deriveDischargeRatio(project);
    if (reason === 'no-cap-typ' || d.reason === 'no-cap-typ') {
      return base + 'edit_note</span>关机电压 <b>' + escapeHtml(b.shutdownVoltage) + ' V</b> → 可放电容量 <b>' + fmt(d.capacityMah, 2)
        + ' mAh</b>；填写电池容量 Typ 后将自动换算并回填比例。';
    }
    if (reason === 'ratio-overflow') {
      return base + 'warning</span>关机电压 <b>' + escapeHtml(b.shutdownVoltage) + ' V</b> → 可放电容量 <b>' + fmt(d.capacityMah, 2)
        + ' mAh</b>，占 Typ 容量 <b>' + fmt(d.ratio * 100, 2) + '%</b>，超过 100%，未自动回填；请检查查询表容量或 Typ 容量是否填写正确。';
    }
    return base + 'check_circle</span>关机电压 <b>' + escapeHtml(b.shutdownVoltage) + ' V</b> → 可放电容量 <b>' + fmt(d.capacityMah, 2)
      + ' mAh</b>（线性插值）→ 占 Typ 容量 <b>' + fmt(d.ratio * 100, 2) + '%</b>，已自动回填「可放电容量比例(0.2C)」。';
  }

  function renderDischargeCard(project) {
    const b = project.battery || {};
    const capTyp = num(b.capTyp);
    // 有效点按电压降序展示；尚未填完的空行追加在末尾，保证「新增一行」立即可见
    const validRows = E.sortDischargeTable(b.dischargeTable);
    const blankRows = (b.dischargeTable || []).filter((r) => num(r.voltage) === null || num(r.capacityMah) === null);
    const rows = validRows.concat(blankRows);
    const body = rows.length ? rows.map((r) => {
      const v = num(r.voltage);
      const c = num(r.capacityMah);
      const ratio = capTyp && capTyp > 0 && c !== null ? c / capTyp : null;
      return '<tr data-es-dt-row="' + escapeHtml(r.id) + '">'
        + '<td><input type="text" inputmode="decimal" class="es-dt-input" data-es-dt-col="voltage"'
        + ' value="' + escapeHtml(r.voltage === undefined || r.voltage === null ? '' : r.voltage) + '" placeholder="示例：3.4" autocomplete="off" /></td>'
        + '<td><input type="text" inputmode="decimal" class="es-dt-input" data-es-dt-col="capacityMah"'
        + ' value="' + escapeHtml(r.capacityMah === undefined || r.capacityMah === null ? '' : r.capacityMah) + '" placeholder="示例：233.7" autocomplete="off" /></td>'
        + '<td class="es-dt-ratio">' + (ratio === null ? '—' : fmt(ratio * 100, 2) + '%') + '</td>'
        + '<td><button type="button" class="es-icon-btn es-icon-danger" data-es-dt-delete="' + escapeHtml(r.id)
        + '" title="删除该行"><span class="material-symbols-outlined">close</span></button></td>'
        + '</tr>';
    }).join('') : '<tr><td colspan="4" class="es-table-empty">暂无数据，导入文档或手动新增电压 - 容量点</td></tr>';
    const derived = E.deriveDischargeRatio(project);
    return '<section class="panel rounded-3xl es-card es-discharge-card">'
      + '<header class="es-card-head"><h3><span class="material-symbols-outlined">table_chart</span>电池可放电容量查询表</h3>'
      + '<div class="es-head-actions">'
      + '<button type="button" class="template-btn template-btn-sm" data-es-dt-import><span class="material-symbols-outlined text-sm">upload_file</span>导入文档自动识别</button>'
      + '<input type="file" accept=".xlsx,.xls,.csv" hidden data-es-dt-file />'
      + '<button type="button" class="template-btn template-btn-sm" data-es-dt-add><span class="material-symbols-outlined text-sm">add</span>新增一行</button>'
      + '<button type="button" class="template-btn template-btn-sm template-btn-danger" data-es-dt-clear><span class="material-symbols-outlined text-sm">delete_sweep</span>清空</button>'
      + '</div></header>'
      + '<p class="es-discharge-tip">容量列指电池从满电放电至该电压时<b>累计可放出的容量（mAh）</b>；导入「剩余容量 / SOC」类表格时会自动按 Typ 容量换算，支持 .xlsx / .xls / .csv。</p>'
      + '<div class="es-table-wrap"><table class="es-table es-discharge-table"><thead><tr>'
      + '<th>电压 V</th><th>可放电容量 mAh</th><th>占 Typ 比例</th><th style="width:64px">操作</th>'
      + '</tr></thead><tbody>' + body + '</tbody></table></div>'
      + '<p class="es-discharge-sync">' + dischargeSyncHint(project, derived.reason, derived.ok ? derived : null) + '</p>'
      + '</section>';
  }

  function bindDischargeCard(project) {
    const b = project.battery || (project.battery = E.defaultBattery());
    if (!Array.isArray(b.dischargeTable)) b.dischargeTable = [];

    const fileInput = root.querySelector('[data-es-dt-file]');
    root.querySelector('[data-es-dt-import]').addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', async (event) => {
      const file = event.target.files && event.target.files[0];
      event.target.value = '';
      if (!file) return;
      try {
        const parsed = await parseDischargeFile(file, num(b.capTyp));
        b.dischargeTable = parsed.rows.map((r) => ({ id: uid('dt'), voltage: r.voltage, capacityMah: r.capacityMah }));
        E.store.save();
        render();
        window.alert('识别成功：从「' + parsed.sheetName + '」导入 ' + parsed.rows.length + ' 个电压点。\n' + parsed.note);
        syncRatioFromTable();
      } catch (error) {
        window.alert('导入失败：' + (error && error.message ? error.message : '无法识别文件内容'));
      }
    });

    root.querySelector('[data-es-dt-add]').addEventListener('click', () => {
      b.dischargeTable.push({ id: uid('dt'), voltage: '', capacityMah: '' });
      E.store.save();
      render();
    });
    root.querySelector('[data-es-dt-clear]').addEventListener('click', () => {
      if (!b.dischargeTable.length) return;
      if (!window.confirm('确认清空电池可放电容量查询表？')) return;
      b.dischargeTable = [];
      E.store.save();
      render();
    });

    root.querySelectorAll('[data-es-dt-delete]').forEach((btn) => {
      btn.addEventListener('click', () => {
        b.dischargeTable = b.dischargeTable.filter((r) => r.id !== btn.dataset.esDtDelete);
        E.store.save();
        render();
        syncRatioFromTable();
      });
    });

    // 行内编辑：非法值标红不入库；失焦后重排（按电压降序）并重算
    root.querySelectorAll('.es-dt-input').forEach((input) => {
      input.addEventListener('input', () => {
        const tr = input.closest('[data-es-dt-row]');
        const row = b.dischargeTable.find((r) => r.id === tr.dataset.esDtRow);
        if (!row) return;
        const raw = input.value.trim();
        const parsed = raw === '' ? '' : num(raw);
        if (raw !== '' && (parsed === null || parsed < 0)) {
          input.classList.add('is-invalid');
          return;
        }
        input.classList.remove('is-invalid');
        row[input.dataset.esDtCol] = raw;
        E.store.save();
        // 就地更新该行占比
        const capTyp = num(b.capTyp);
        const ratioCell = tr.querySelector('.es-dt-ratio');
        const v = num(row.voltage);
        const c = num(row.capacityMah);
        if (ratioCell) ratioCell.textContent = (capTyp && capTyp > 0 && v !== null && c !== null) ? fmt(c / capTyp * 100, 2) + '%' : '—';
        syncRatioFromTable();
      });
      input.addEventListener('blur', () => {
        if (input.classList.contains('is-invalid')) return;
        // 失焦后按电压降序重排有效点（空行保留在末尾），顺序未变则不重绘以免打断操作
        const valid = b.dischargeTable.filter((r) => num(r.voltage) !== null && num(r.capacityMah) !== null);
        const blanks = b.dischargeTable.filter((r) => num(r.voltage) === null || num(r.capacityMah) === null);
        const sortedValid = E.sortDischargeTable(valid);
        const orderChanged = sortedValid.some((r, i) => r.id !== valid[i].id);
        if (orderChanged) {
          b.dischargeTable = sortedValid.concat(blanks);
          E.store.save();
          render();
        }
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
  const { escapeHtml, uid, num, fmt, todayStr, debounce } = E;
  const root = document.getElementById('enduranceStatusRoot');

  const refreshModelsQuiet = debounce(function () {
    window.ES_ENDURANCE_UI.refreshModelAndCompare && window.ES_ENDURANCE_UI.refreshModelAndCompare();
  }, 200);

  function renderScenarioCard(project) {
    const customFields = project.scenarioFields || [];
    const head = '<th>场景</th><th>功耗标准</th><th>当前电流</th><th>测试版本</th><th>更新日期</th><th>备注/说明</th>'
      + customFields.map((f) => '<th class="es-custom-th">' + escapeHtml(f.name)
        + '<button type="button" class="es-field-del" data-es-field-delete="' + escapeHtml(f.id) + '" title="删除该列"><span class="material-symbols-outlined">close</span></button></th>').join('')
      + '<th class="es-col-ops">管理</th>';

    const rows = project.scenarios.map((scenario) => {
      const rec = E.latestRecord(scenario);
      const custom = num(scenario.customCurrent);
      const isCustom = custom !== null;
      const current = isCustom ? custom : (rec ? num(rec.current) : null);
      const currentCell = '<td class="es-sc-current">'
        + (current === null ? '—' : fmt(current, 3) + ' mA')
        + (isCustom ? '<em class="es-custom-chip" title="自定义电流，不来自测试记录">自定义</em>' : '')
        + '<button type="button" class="es-icon-btn es-inline-edit" data-es-action="custom-current" title="设置 / 清除自定义电流"><span class="material-symbols-outlined">edit</span></button>'
        + (isCustom ? '<button type="button" class="es-icon-btn es-inline-clear" data-es-action="clear-custom" title="清除自定义，恢复测试记录电流"><span class="material-symbols-outlined">restart_alt</span></button>' : '')
        + '</td>';
      const customCells = customFields.map((f) => '<td><input type="text" class="es-field-input" data-es-field-row="' + escapeHtml(scenario.id)
        + '" data-es-field-col="' + escapeHtml(f.id) + '" value="' + escapeHtml(scenario.fields && scenario.fields[f.id] || '') + '" autocomplete="off" /></td>').join('');
      return '<tr data-es-scenario="' + escapeHtml(scenario.id) + '">'
        + '<td class="es-sc-name">' + escapeHtml(scenario.name)
        + '<button type="button" class="es-icon-btn es-inline-edit" data-es-action="rename-scenario" title="重命名场景"><span class="material-symbols-outlined">drive_file_rename_outline</span></button></td>'
        + '<td><input type="text" class="es-field-input es-standard-input" data-es-standard="' + escapeHtml(scenario.id)
        + '" value="' + escapeHtml(scenario.standard || '') + '" placeholder="功耗标准 / 门限" autocomplete="off" /></td>'
        + currentCell
        + '<td class="es-sc-version">' + (rec ? escapeHtml(rec.version || '—') : '—') + '</td>'
        + '<td class="es-sc-date">' + (rec ? escapeHtml(rec.date || '—') : '暂无数据') + '</td>'
        + '<td><input type="text" class="es-field-input es-remark-input" data-es-remark="' + escapeHtml(scenario.id)
        + '" value="' + escapeHtml(scenario.remark || '') + '" placeholder="备注 / 说明" autocomplete="off" /></td>'
        + customCells
        + '<td class="es-sc-ops">'
        + '<button type="button" class="es-mini-btn" data-es-action="update-current" title="按日期/版本录入测试电流（留痕）"><span class="material-symbols-outlined">edit_note</span>更新</button>'
        + '<button type="button" class="es-mini-btn" data-es-action="history" title="查看历史记录"><span class="material-symbols-outlined">history</span>历史</button>'
        + '<button type="button" class="es-mini-btn es-mini-danger" data-es-action="delete-scenario" title="删除场景"><span class="material-symbols-outlined">close</span></button>'
        + '</td></tr>';
    }).join('');

    const colCount = 7 + customFields.length;
    return '<section class="panel rounded-3xl es-card">'
      + '<header class="es-card-head"><h3><span class="material-symbols-outlined">bolt</span>场景电流</h3>'
      + '<div class="es-head-actions">'
      + '<button type="button" class="template-btn template-btn-sm" data-es-add-field><span class="material-symbols-outlined text-sm">view_column</span>新增列</button>'
      + '<button type="button" class="template-btn template-btn-sm es-add-scenario-btn"><span class="material-symbols-outlined text-sm">add</span>新增场景</button>'
      + '</div></header>'
      + '<div class="es-table-wrap"><table class="es-table es-scenario-table"><thead><tr>' + head + '</tr></thead><tbody>'
      + (rows || '<tr><td colspan="' + colCount + '" class="es-table-empty">还没有场景，点右上角「新增场景」开始维护；测试电流按日期/版本留痕，也可直接填写自定义电流</td></tr>')
      + '</tbody></table></div></section>';
  }

  function bindScenarios(project) {
    const addBtn = root.querySelector('.es-add-scenario-btn');
    if (addBtn) addBtn.addEventListener('click', () => openScenarioEditor(project, null));

    root.querySelector('[data-es-add-field]').addEventListener('click', () => {
      const name = window.prompt('新列名称（如：功耗来源、测试人）：', '');
      if (name === null) return;
      const trimmed = name.trim();
      if (!trimmed) { alert('列名不能为空'); return; }
      if (!Array.isArray(project.scenarioFields)) project.scenarioFields = [];
      if (project.scenarioFields.some((f) => f.name === trimmed)) { alert('已存在同名列'); return; }
      project.scenarioFields.push({ id: uid('sf'), name: trimmed });
      E.store.save();
      window.ES_ENDURANCE_UI.render();
    });

    root.querySelectorAll('[data-es-field-delete]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const fid = btn.dataset.esFieldDelete;
        const field = (project.scenarioFields || []).find((f) => f.id === fid);
        if (!field) return;
        if (!window.confirm('确认删除列「' + field.name + '」？各场景在该列填写的内容会一并删除。')) return;
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
        if (!scenario || scenario.standard === input.value) return;
        scenario.standard = input.value.trim();
        E.store.save();
      });
    });
    root.querySelectorAll('.es-remark-input').forEach((input) => {
      input.addEventListener('blur', () => {
        const scenario = project.scenarios.find((s) => s.id === input.dataset.esRemark);
        if (!scenario || scenario.remark === input.value) return;
        scenario.remark = input.value.trim();
        E.store.save();
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
          else if (action === 'custom-current') openCustomCurrent(project, id);
          else if (action === 'clear-custom') clearCustomCurrent(project, id);
          else if (action === 'delete-scenario') deleteScenario(project, id);
        });
      });
    });
  }

  function deleteScenario(project, id) {
    const scenario = project.scenarios.find((s) => s.id === id);
    if (!scenario) return;
    if (!window.confirm('确认删除场景「' + scenario.name + '」及其全部电流记录？续航模型中引用该场景的行会一并删除。')) return;
    project.scenarios = project.scenarios.filter((s) => s.id !== id);
    project.models.forEach((model) => {
      model.rows = model.rows.filter((r) => r.scenarioId !== id);
    });
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  // 新增 / 重命名场景；新增时同步给该项目每个模型加一行
  function openScenarioEditor(project, scenario) {
    const name = window.prompt(scenario ? '修改场景名称：' : '请输入新场景名称：', scenario ? scenario.name : '');
    if (name === null) return;
    const trimmed = name.trim();
    if (!trimmed) { alert('场景名称不能为空'); return; }
    if (scenario) {
      scenario.name = trimmed;
    } else {
      const created = {
        id: uid('sc'), name: trimmed, remark: '', standard: '',
        customCurrent: '', fields: {}, records: [],
      };
      project.scenarios.push(created);
      project.models.forEach((model) => {
        const cells = {};
        model.columns.forEach((col) => { cells[col.id] = { d: '0', m: E.DEFAULT_MAH_FORMULA }; });
        model.rows.push({ id: uid('mr'), scenarioId: created.id, cells: cells });
      });
    }
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  // 自定义电流：不进入测试历史，直接覆盖生效值；留空则清除自定义
  function openCustomCurrent(project, scenarioId) {
    const scenario = project.scenarios.find((s) => s.id === scenarioId);
    if (!scenario) return;
    const text = window.prompt(
      '为「' + scenario.name + '」设置自定义电流（mA）：\n'
      + '自定义电流会直接覆盖测试记录用于公式计算，但不会进入历史记录。\n'
      + '清空输入并确定，可恢复为最新测试记录电流。',
      scenario.customCurrent === undefined || scenario.customCurrent === null ? '' : String(scenario.customCurrent),
    );
    if (text === null) return;
    const trimmed = text.trim();
    if (trimmed === '') { scenario.customCurrent = ''; }
    else {
      const v = num(trimmed);
      if (v === null || v < 0) { alert('电流需为不小于 0 的数字'); return; }
      scenario.customCurrent = String(v);
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

  // 电流更新弹窗：日期 / 版本 / 电流 / 备注，确认后追加一条记录（历史不覆盖）
  function openCurrentEditor(project, scenarioId) {
    const scenario = project.scenarios.find((s) => s.id === scenarioId);
    if (!scenario) return;
    const last = E.latestRecord(scenario);
    const date = window.prompt('更新日期（YYYY-MM-DD）：', last ? last.date : todayStr());
    if (date === null) return;
    const dateTrim = date.trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateTrim)) { alert('日期格式应为 YYYY-MM-DD'); return; }
    const version = window.prompt('测试版本号：', last ? (last.version || '') : '');
    if (version === null) return;
    const currentText = window.prompt('场景电流（mA）：', last ? String(last.current) : '');
    if (currentText === null) return;
    const current = num(currentText);
    if (current === null || current < 0) { alert('电流需为不小于 0 的数字'); return; }
    const note = window.prompt('备注 / 说明（可留空）：', last ? (last.note || '') : '');
    if (note === null) return;
    scenario.records.push({
      id: uid('rec'), date: dateTrim, version: version.trim(),
      current: current, note: note.trim(),
    });
    scenario.records.sort((a, b) => String(a.date).localeCompare(String(b.date)));
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  // 历史记录查看
  function openHistory(project, scenarioId) {
    const scenario = project.scenarios.find((s) => s.id === scenarioId);
    if (!scenario || !scenario.records.length) { alert('该场景还没有测试电流记录'); return; }
    const lines = scenario.records.slice().reverse().map((r, i) => {
      const tag = i === 0 && num(scenario.customCurrent) === null ? '（当前生效）' : '';
      return (r.date || '—') + ' | ' + (r.version || '未填版本') + ' | ' + fmt(num(r.current), 3) + ' mA'
        + (r.note ? ' | ' + r.note : '') + tag;
    });
    const customLine = num(scenario.customCurrent) !== null
      ? '当前生效为自定义电流 ' + fmt(num(scenario.customCurrent), 3) + ' mA（测试记录不参与计算）\n\n'
      : '';
    window.alert('「' + scenario.name + '」电流更新历史（共 ' + scenario.records.length + ' 条）：\n' + customLine + '\n' + lines.join('\n'));
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
    if (desc.kind === 'cell') {
      const row = model.rows.find((r) => r.id === desc.rowId);
      const cell = row && row.cells[desc.colId];
      return cell && cell[desc.t] !== undefined && cell[desc.t] !== null ? String(cell[desc.t]) : '';
    }
    return (model.foot[desc.colId] && model.foot[desc.colId][desc.t]) || '';
  }
  function setRaw(model, desc, raw) {
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
      + '<header class="es-card-head"><h3><span class="material-symbols-outlined">modeling</span>自定义续航模型</h3>'
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

  function formulaCellHtml(ev, model, desc, digits) {
    const addr = E.addressOf(model, desc);
    const raw = getRaw(model, desc);
    const r = ev.result(desc);
    const isFormula = String(raw).trim().charAt(0) === '=';
    const cls = 'es-formula-cell' + (isFormula ? ' is-formula' : '') + (!r.ok ? ' is-error' : '');
    const text = !r.ok ? r.error : (r.empty ? '—' : fmt(r.value, digits === undefined ? 2 : digits));
    return '<td class="' + cls + '" title="地址 ' + escapeHtml(addr) + '，点击 fx 编辑公式 / 数值">'
      + '<span class="es-cell-text' + (!r.ok ? ' es-cell-text-error' : '') + '">' + escapeHtml(text) + '</span>'
      + '<button type="button" class="es-fx-btn" data-es-fx="' + escapeHtml(descKey(desc)) + '" title="编辑 ' + escapeHtml(addr) + '">fx</button></td>';
  }

  function renderModelBody(project, model) {
    const ev = E.evaluateModel(project, model);

    // 两行表头
    const headTop = '<th class="es-corner" rowspan="2">场景</th><th rowspan="2">电流 mA</th>'
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
        return '<th class="es-mode-head" colspan="2"><div class="es-mode-title">' + escapeHtml(col.name) + tag + ops + '</div>'
          + '<small class="es-col-addr">时长列 ' + dLetter + ' / 耗电列 ' + mLetter + '</small>' + win + need + note + '</th>';
      }).join('');
    const headSub = model.columns.map(() => '<th class="es-sub">时长 h</th><th class="es-sub">耗电量 mAh</th>').join('');

    const body = model.rows.map((row, rowIndex) => {
      const scenario = project.scenarios.find((s) => s.id === row.scenarioId);
      const scenarioName = scenario ? scenario.name : '（场景已删除）';
      const current = scenario ? E.scenarioCurrent(project, scenario.id) : null;
      const cells = model.columns.map((col, ci) => {
        const cell = row.cells[col.id] || { d: '', m: E.DEFAULT_MAH_FORMULA };
        const dDesc = { kind: 'cell', rowId: row.id, colId: col.id, t: 'd' };
        const mDesc = { kind: 'cell', rowId: row.id, colId: col.id, t: 'm' };
        const dAddr = E.addressOf(model, dDesc);
        const dResult = ev.result(dDesc);
        const dInvalid = !dResult.ok;
        const isFormulaD = String(cell.d || '').trim().charAt(0) === '=';
        const dInput = '<td class="es-duration-wrap' + (dInvalid ? ' is-error' : '') + (isFormulaD ? ' is-formula' : '') + '">'
          + '<input type="text" class="es-duration-input" data-es-row="' + escapeHtml(row.id) + '" data-es-col="' + escapeHtml(col.id) + '"'
          + ' value="' + escapeHtml(cell.d === undefined || cell.d === null ? '' : cell.d) + '"'
          + ' title="地址 ' + escapeHtml(dAddr) + '；可填数字，或以 = 开头写公式" autocomplete="off" />'
          + '<small class="es-formula-result' + (dInvalid ? ' is-error' : '') + '"' + (isFormulaD ? '' : ' hidden') + '>'
          + (isFormulaD ? (dInvalid ? escapeHtml(dResult.error) : fmt(dResult.value, 2)) : '') + '</small>'
          + '</td>';
        return dInput + formulaCellHtml(ev, model, mDesc);
      }).join('');
      return '<tr data-es-model-row="' + escapeHtml(row.id) + '"><td class="es-sc-name">' + escapeHtml(scenarioName) + '</td>'
        + '<td class="es-current-cell">' + (current === null ? '—' : fmt(current, 2)) + '</td>' + cells + '</tr>';
    }).join('');

    // tfoot：总时长/总耗电（可编辑公式）、平均电流（自动）、可使用时长 Typ/Min（可编辑公式）
    // 每个对比列含 时长 / 耗电 两个子格，严格与表体对齐
    const totalRowHtml = '<tr class="es-stat-row"><td>Total（可编辑公式）</td><td></td>'
      + model.columns.map((col) => formulaCellHtml(ev, model, { kind: 'foot', colId: col.id, t: 'td' }, 2)
        + formulaCellHtml(ev, model, { kind: 'foot', colId: col.id, t: 'tm' }, 2)).join('') + '</tr>';

    const avgRow = '<tr class="es-stat-row es-stat-auto"><td>平均电流 mA（自动）</td><td></td>'
      + model.columns.map((col) => {
        const s = ev.stats[col.id] || {};
        return '<td colspan="2" class="es-stat-mah">' + (s.avgCurrent === null ? '—' : fmt(s.avgCurrent, 2)) + '</td>';
      }).join('') + '</tr>';

    const douRow = (label, cellType) => '<tr class="es-stat-row"><td>' + label + '</td><td></td>'
      + model.columns.map((col) => formulaCellHtml(ev, model, { kind: 'foot', colId: col.id, t: cellType }, 2)
        + '<td class="es-stat-blank"></td>').join('') + '</tr>';

    const totals = totalRowHtml + avgRow
      + douRow('可使用时长（Typ）h', 'douTyp')
      + douRow('可使用时长（Min）h', 'douMin');

    return '<div class="es-mode-toolbar">'
      + '<button type="button" class="es-mini-btn" data-es-model-action="add-row"><span class="material-symbols-outlined">playlist_add</span>新增场景行</button>'
      + '<button type="button" class="es-mini-btn" data-es-model-action="add-col"><span class="material-symbols-outlined">view_column</span>新增对比列</button>'
      + '<button type="button" class="es-mini-btn" data-es-model-action="rename"><span class="material-symbols-outlined">drive_file_rename_outline</span>重命名模型</button>'
      + '<button type="button" class="es-mini-btn" data-es-model-action="help"><span class="material-symbols-outlined">function</span>公式说明</button>'
      + '<button type="button" class="es-mini-btn es-mini-danger" data-es-model-action="delete-model"><span class="material-symbols-outlined">delete</span>删除模型</button>'
      + '</div>'
      + '<div class="es-model-table-wrap"><table class="es-model-table"><thead><tr>' + headTop + '</tr><tr>' + headSub + '</tr></thead>'
      + '<tbody>' + (body || '<tr><td colspan="' + (2 + model.columns.length * 2) + '" class="es-table-empty">请先在上方「场景电流」中新增场景，或点工具栏「新增场景行」</td></tr>') + '</tbody>'
      + '<tfoot>' + totals + '</tfoot></table></div>'
      + '<p class="es-model-tip"><span class="material-symbols-outlined">info</span>'
      + '默认口径对齐云表：场景耗电 = 时长 × 场景电流（0 时长自动得 0）；可使用时长 = 可用容量（Typ × 可放电比例，不硬编码）÷ 总耗电 × 总时长。'
      + '点单元格右侧 <b>fx</b> 可编辑任意公式；时长格也支持直接输入 = 公式。</p>';
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
          const dDesc = { kind: 'cell', rowId: rowId, colId: col.id, t: 'd' };
          const mDesc = { kind: 'cell', rowId: rowId, colId: col.id, t: 'm' };
          const dResult = ev.result(dDesc);
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
          if (mCell) paintFormulaCell(mCell, ev.result(mDesc));
        });
      });
    }
    const tfoot = root.querySelector('.es-model-table tfoot');
    if (tfoot) {
      ['td', 'tm', 'douTyp', 'douMin'].forEach((t) => {
        model.columns.forEach((col) => {
          const desc = { kind: 'foot', colId: col.id, t: t };
          const cell = findFxCell(tfoot, desc);
          if (cell) paintFormulaCell(cell, ev.result(desc));
        });
      });
      // 平均电流行自动值（每列一个 colspan=2 单元格）
      const autoRow = tfoot.querySelector('.es-stat-auto');
      if (autoRow) {
        const avgValues = model.columns.map((col) => (ev.stats[col.id] || {}).avgCurrent);
        autoRow.querySelectorAll('.es-stat-mah').forEach((td, i) => {
          td.textContent = avgValues[i] === null || avgValues[i] === undefined ? '—' : fmt(avgValues[i], 2);
        });
      }
    }
  }

  function paintFormulaCell(cell, result) {
    cell.classList.toggle('is-error', !result.ok);
    const text = cell.querySelector('.es-cell-text');
    if (!text) return;
    text.classList.toggle('es-cell-text-error', !result.ok);
    text.textContent = !result.ok ? result.error : (result.empty ? '—' : fmt(result.value, 2));
  }

  // ---------- 事件绑定 ----------
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

    // fx 公式编辑
    root.querySelectorAll('[data-es-fx]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const model = activeModel(project);
        const desc = parseDesc(btn.dataset.esFx);
        if (!model || !desc) return;
        openFormulaEditor(project, model, desc);
      });
    });

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
        // 公式交给引擎求值并展示语义错误；纯数字必须为不小于 0 的有限数，非法值标红且不入库
        const parsed = num(trimmed);
        const invalid = !isFormula && trimmed !== '' && (parsed === null || parsed < 0);
        wrap.classList.toggle('is-error', invalid);
        if (invalid) return;
        if (!row.cells[input.dataset.esCol]) row.cells[input.dataset.esCol] = { d: '', m: E.DEFAULT_MAH_FORMULA };
        row.cells[input.dataset.esCol].d = rawText;
        // 公式型时长格补充样式标记
        wrap.classList.toggle('is-formula', isFormula);
        E.store.save();
        updateModelCells(project, model);
        quietRefreshCompare();
      });
    }
  }

  // ---------- 模型 / 列 / 行管理 ----------
  function handleModelAction(project, action) {
    const model = activeModel(project);
    if (action === 'add') {
      const name = window.prompt('新模型名称：', '续航模型 ' + (project.models.length + 1));
      if (name === null) return;
      const trimmed = name.trim();
      if (!trimmed) { alert('模型名称不能为空'); return; }
      const created = E.createModelV2(project, trimmed, [{ name: '典型使用', kind: 'normal' }]);
      project.models.push(created);
      project.activeModelId = created.id;
      E.store.save();
      window.ES_ENDURANCE_UI.render();
      return;
    }
    if (action === 'copy') {
      if (!model) return;
      const name = window.prompt('复制出的模型名称：', model.name + '（副本）');
      if (name === null) return;
      const trimmed = name.trim();
      if (!trimmed) { alert('模型名称不能为空'); return; }
      const copy = cloneModel(model, trimmed);
      project.models.push(copy);
      project.activeModelId = copy.id;
      E.store.save();
      window.ES_ENDURANCE_UI.render();
      return;
    }
    if (action === 'help') { openFormulaHelp(activeModel(project)); return; }
    if (!model) return;
    if (action === 'rename') {
      const name = window.prompt('修改模型名称：', model.name);
      if (name === null) return;
      const trimmed = name.trim();
      if (!trimmed) { alert('模型名称不能为空'); return; }
      model.name = trimmed;
      E.store.save();
      window.ES_ENDURANCE_UI.render();
      return;
    }
    if (action === 'delete-model') {
      if (!window.confirm('确认删除模型「' + model.name + '」？')) return;
      project.models = project.models.filter((m) => m.id !== model.id);
      project.activeModelId = project.models[0] ? project.models[0].id : null;
      E.store.save();
      window.ES_ENDURANCE_UI.render();
      return;
    }
    if (action === 'add-col') { addColumn(project, model); return; }
    if (action === 'add-row') { addScenarioRow(project, model); return; }
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
      return { id: uid('mr'), scenarioId: row.scenarioId, cells: cells };
    });
    return copy;
  }

  function addColumn(project, model) {
    const name = window.prompt('新对比列（使用模式）名称，如：中度使用：', '');
    if (name === null) return;
    const trimmed = name.trim();
    if (!trimmed) { alert('列名称不能为空'); return; }
    const full = window.confirm('该列是否为「满电使用」口径？\n\n确定 = 满电（可使用时长按 Typ / Min 满容量计算）\n取消 = 普通口径（按可用容量 = Typ × 可放电比例 计算）');
    const col = E.createColumn(trimmed, full ? 'full' : 'normal');
    model.columns.push(col);
    model.foot[col.id] = E.defaultFoot(col.kind);
    project.scenarios.forEach((scenario) => {
      let row = model.rows.find((r) => r.scenarioId === scenario.id);
      if (!row) {
        row = { id: uid('mr'), scenarioId: scenario.id, cells: {} };
        model.rows.push(row);
      }
      row.cells[col.id] = { d: '0', m: E.DEFAULT_MAH_FORMULA };
    });
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  function addScenarioRow(project, model) {
    const missing = project.scenarios.filter((s) => !model.rows.some((r) => r.scenarioId === s.id));
    const options = missing.map((s, i) => (i + 1) + '. ' + s.name).join('\n');
    const tip = '输入要加入本模型的场景序号；输入 0 新建场景。\n\n'
      + (options || '（项目场景已全部加入本模型，只能新建）') + '\n0. 新建场景';
    const answer = window.prompt(tip, '1');
    if (answer === null) return;
    const pick = num(answer);
    // pick=0 表示新建场景；其余序号必须落在可加入场景范围内
    if (pick === 0) {
      const created = createScenarioInteractively(project);
      if (!created) return;
      appendRowForScenario(model, created.id);
    } else {
      if (pick === null || pick < 1 || pick > missing.length) { alert('序号无效'); return; }
      appendRowForScenario(model, missing[pick - 1].id);
    }
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  function appendRowForScenario(model, scenarioId) {
    if (model.rows.some((r) => r.scenarioId === scenarioId)) return;
    const cells = {};
    model.columns.forEach((col) => { cells[col.id] = { d: '0', m: E.DEFAULT_MAH_FORMULA }; });
    model.rows.push({ id: uid('mr'), scenarioId: scenarioId, cells: cells });
  }

  function createScenarioInteractively(project) {
    const name = window.prompt('请输入新场景名称：', '');
    if (name === null) return null;
    const trimmed = name.trim();
    if (!trimmed) { alert('场景名称不能为空'); return null; }
    const created = {
      id: uid('sc'), name: trimmed, remark: '', standard: '',
      customCurrent: '', fields: {}, records: [],
    };
    project.scenarios.push(created);
    return created;
  }

  function openColumnEditor(project, model, colId) {
    if (!model) return;
    const col = model.columns.find((c) => c.id === colId);
    if (!col) return;
    const name = window.prompt('对比列名称：', col.name);
    if (name === null) return;
    const trimmed = name.trim();
    if (!trimmed) { alert('列名称不能为空'); return; }
    const win = window.prompt('使用时间窗口（可留空，如 8am~8pm）：', col.window || '');
    if (win === null) return;
    const needText = window.prompt('目标耗电量 mAh（可留空）：', col.needMah === null || col.needMah === undefined ? '' : String(col.needMah));
    if (needText === null) return;
    const need = num(needText);
    if (needText.trim() !== '' && (need === null || need < 0)) { alert('目标耗电量需为不小于 0 的数字'); return; }
    const note = window.prompt('备注（可留空）：', col.note || '');
    if (note === null) return;
    const full = window.confirm('该列是否为「满电使用」口径？\n\n确定 = 满电（可使用时长按 Typ / Min 满容量计算）\n取消 = 普通口径（按可用容量计算）');
    col.name = trimmed;
    col.window = win.trim();
    col.needMah = needText.trim() === '' ? null : need;
    col.note = note.trim();
    const newKind = full ? 'full' : 'normal';
    if (newKind !== col.kind) {
      col.kind = newKind;
      model.foot[col.id] = E.defaultFoot(newKind);
    }
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  function deleteColumn(project, model, colId) {
    if (!model) return;
    const col = model.columns.find((c) => c.id === colId);
    if (!col) return;
    if (model.columns.length <= 1) { alert('每个续航模型至少保留一个对比列'); return; }
    if (!window.confirm('确认删除对比列「' + col.name + '」？该列下各场景的时长 / 耗电公式会一并删除。')) return;
    model.columns = model.columns.filter((c) => c.id !== colId);
    delete model.foot[colId];
    model.rows.forEach((row) => { delete row.cells[colId]; });
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
      + '<p class="es-formula-sub">直接填数字，或以 <b>=</b> 开头编写公式；留空表示空单元格。</p>'
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
      if (desc.kind === 'cell') {
        textarea.value = desc.t === 'm' ? E.DEFAULT_MAH_FORMULA : '0';
      } else {
        const col = model.columns.find((c) => c.id === desc.colId);
        textarea.value = E.defaultFoot(col ? col.kind : 'normal')[desc.t];
      }
      trial();
    });
    overlay.querySelector('[data-es-fx-save]').addEventListener('click', () => {
      const text = textarea.value.trim();
      if (text.charAt(0) === '=') {
        const check = window.ES_FORMULA.checkSyntax(text.slice(1));
        if (!check.ok) { alert('公式语法错误：' + check.error); return; }
      } else if (text !== '' && num(text) === null) {
        alert('请输入数字，或以 = 开头的公式');
        return;
      }
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
      + '<p>A 列为场景名；第 1 行表头，场景行从第 2 行开始'
      + (model ? '（本表共 ' + n + ' 个场景行，Total 在第 ' + E.totalRowNumber(model) + ' 行，可使用时长在第 ' + E.douRowNumber(model) + ' 行）' : '') + '；'
      + '第 1 个对比列占 B（时长）/ C（耗电），第 2 个占 D / E，依次类推。鼠标悬停单元格可见地址。'
      + '表中「平均电流」为自动展示行，不占用公式地址；Typ / Min 两个可使用时长同处虚拟 DOU 行的时长 / 耗电子列。</p>'
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
      + '<div class="es-chart-wrap"><h4>场景电流对比（mA）</h4><canvas id="esCurrentChart"></canvas></div>'
      + '<div class="es-chart-wrap"><h4>可使用时长 DOU 对比（h，按各项目激活模型的同名对比列）</h4><canvas id="esDouChart"></canvas></div>'
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

  // 在同一 canvas 上重绘前，先销毁并摘除登记在册的旧 Chart 实例，避免 "Canvas is already in use" 与实例泄漏
  function reuseCanvas(canvas) {
    for (let i = E.chartInstances.length - 1; i >= 0; i -= 1) {
      const inst = E.chartInstances[i];
      if (inst.canvas === canvas || inst.ctx && inst.ctx.canvas === canvas) {
        try { inst.destroy(); } catch (e) { /* ignore */ }
        E.chartInstances.splice(i, 1);
      }
    }
  }

  function selectedProjects() {
    const data = E.store.load();
    return data.projects.filter((p) => compareState.projects.includes(p.id));
  }

  function drawCurrentChart() {
    const canvas = root.querySelector('#esCurrentChart');
    if (!canvas || typeof Chart === 'undefined') return;
    reuseCanvas(canvas);
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
        responsive: true, maintainAspectRatio: false,
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
    reuseCanvas(canvas);
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
        responsive: true, maintainAspectRatio: false,
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
