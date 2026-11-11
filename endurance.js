/* ================================================================
 * 续航 / 功耗现状模块（endurance.js）
 * 纯前端、localStorage 持久化，不依赖 Supabase，游客模式可用。
 * 数据结构见 ENDURANCE_DEFAULT_DATA；计算公式见 computeModelStats。
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
      dischargeRatio: '', lowThreshold: '',
      envStandbyCurrent: '', wakeCurrent: '',
    };
  }

  // 首次使用的示例项目（对齐参考截图，可直接删除）
  function seedData() {
    const scenarioDefs = [
      ['蓝牙音频+环境', 27.44, '中档：慢档=1：2'],
      ['蓝牙音频', 22.82, ''],
      ['蓝牙通话', 21.8, ''],
      ['视频通话（WIFI）', 105.76, ''],
      ['拍照', 91.75, ''],
      ['录像1080P', 198.41, ''],
      ['录音纪要', 12.55, ''],
      ['环境记忆待机', 17.52, '中档：慢档=1：2'],
      ['待机', 10.31, ''],
      ['视频通话（BT OCR）', 98.295, ''],
    ];
    const scenarios = scenarioDefs.map(([name, current, remark]) => ({
      id: uid('sc'),
      name: name,
      remark: remark,
      records: [{ id: uid('rec'), date: todayStr(), version: 'J3501', current: current, note: '示例数据' }],
    }));
    const byName = {};
    scenarios.forEach((s) => { byName[s.name] = s.id; });

    const modes = [
      { id: uid('md'), name: '重度使用', window: '8am~8pm', needMah: 485, note: '', fullOnly: false },
      { id: uid('md'), name: '满电使用', window: '', needMah: null, note: '', fullOnly: true },
      { id: uid('md'), name: '轻度使用', window: '8am~6pm', needMah: 240, note: '在意隐私的商务人士，偶尔视觉对话', fullOnly: false },
      { id: uid('md'), name: '功能尝鲜', window: '8am~8pm', needMah: 459, note: '会议少，对全天记录比较好奇', fullOnly: false },
    ];
    const m = {};
    modes.forEach((mode, i) => { m[i] = mode.id; });
    // [场景, 重度h, 满电h, 轻度h, 尝鲜h]
    const durationDefs = [
      ['蓝牙音频+环境', 2.0, 1.53, 0, 2.0],
      ['蓝牙音频', 0, 0, 1.5, 0],
      ['蓝牙通话', 4.0, 3.06, 2.5, 1.0],
      ['视频通话（WIFI）', 0.5, 0.38, 0.2, 0.5],
      ['拍照', 0, 0, 0, 0],
      ['录像1080P', 0.1, 0.0765, 0, 0.1],
      ['录音纪要', 2.0, 1.53, 0, 0.5],
      ['环境记忆待机', 3.4, 2.6, 0, 7.9],
      ['待机', 0, 0, 5.8, 0],
      ['视频通话（BT OCR）', 0.5, 0.38, 0.2, 0.5],
    ];
    const rows = durationDefs.map(([name, h0, h1, h2, h3]) => ({
      scenarioId: byName[name],
      durations: {
        [m[0]]: h0, [m[1]]: h1, [m[2]]: h2, [m[3]]: h3,
      },
    }));

    const project = {
      id: uid('p'),
      name: '示例项目（可删除）',
      createdAt: Date.now(),
      battery: Object.assign(defaultBattery(), {
        capTyp: 246, capMin: 240, wh: 0.95, voltage: 3.87,
        dischargeRatio: 0.95, lowThreshold: 0.02,
        envStandbyCurrent: 14, wakeCurrent: 67,
      }),
      scenarios: scenarios,
      models: [{
        id: uid('m'),
        name: '默认续航模型',
        modes: modes,
        rows: rows,
      }],
      activeModelId: null,
    };
    project.activeModelId = project.models[0].id;
    return { activeProjectId: project.id, projects: [project] };
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

  // ---------- 业务查询与公式 ----------
  // 场景当前生效电流 = 最新一条记录（按日期，同日取最后录入）
  function latestRecord(scenario) {
    if (!scenario || !Array.isArray(scenario.records) || !scenario.records.length) return null;
    return scenario.records.slice().sort((a, b) => {
      const cmp = String(a.date).localeCompare(String(b.date));
      return cmp !== 0 ? cmp : 0;
    }).slice(-1)[0];
  }

  function scenarioCurrent(project, scenarioId) {
    const scenario = project.scenarios.find((s) => s.id === scenarioId);
    const rec = latestRecord(scenario);
    return rec ? num(rec.current) : null;
  }

  // 单个使用模式的统计：场景耗电量、总时长、平均电流、满电时长、DOU
  function computeModeStats(project, model, mode) {
    const details = [];
    let totalHours = 0;
    let totalMah = 0;
    model.rows.forEach((row) => {
      const hours = num(row.durations[mode.id]);
      const current = scenarioCurrent(project, row.scenarioId);
      const h = hours || 0;
      const mah = (hours !== null && current !== null) ? hours * current : null;
      if (h) totalHours += h;
      if (mah !== null) totalMah += mah;
      details.push({ scenarioId: row.scenarioId, hours: hours, current: current, mah: mah });
    });
    const avgCurrent = totalHours > 0 ? totalMah / totalHours : null;
    const b = project.battery || {};
    const capTyp = num(b.capTyp);
    const capMin = num(b.capMin);
    const ratio = num(b.dischargeRatio);
    const threshold = num(b.lowThreshold);
    const derate = (ratio === null ? 1 : ratio) * (1 - (threshold === null ? 0 : threshold));
    const fullHours = (avgCurrent && capTyp) ? capTyp / avgCurrent : null;
    const douTyp = (avgCurrent && capTyp) ? capTyp * derate / avgCurrent : null;
    const douMin = (avgCurrent && capMin) ? capMin * derate / avgCurrent : null;
    return { details, totalHours, totalMah, avgCurrent, fullHours, douTyp, douMin, derate };
  }

  // 整个模型的统计（每个模式一份）
  function computeModelStats(project, model) {
    const modes = {};
    model.modes.forEach((mode) => { modes[mode.id] = computeModeStats(project, model, mode); });
    return modes;
  }

  window.ES_ENDURANCE = {
    STORAGE_KEY, CHART_COLORS, chartInstances, destroyCharts,
    uid, escapeHtml, num, fmt, todayStr, debounce,
    store, defaultBattery, latestRecord, scenarioCurrent,
    computeModeStats, computeModelStats,
  };
})();

/* ================================================================
 * 第二部分：页面骨架、项目管理、电池参数
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
      window.ES_ENDURANCE_SCENARIO.bindScenarios(project);
      window.ES_ENDURANCE_MODEL.bindModels(project);
      window.ES_ENDURANCE_UI.bindCompare(project);
    }
  }

  function renderEmpty() {
    return '<div class="es-empty panel rounded-3xl">'
      + '<span class="material-symbols-outlined es-empty-icon">battery_saver</span>'
      + '<h3>还没有项目</h3>'
      + '<p>新增一个项目后即可维护电池参数、场景电流台账与自定义续航模型，并进行多项目横向对比。</p>'
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
        if (!window.confirm('确认删除项目「' + project.name + '」？\n该项目下的电池参数、场景电流台账与全部续航模型都会被删除，且不可恢复。')) return;
        E.store.removeProject(project.id);
        render();
      });
    }
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
    { key: 'dischargeRatio', label: '可放电容量比例', placeholder: '0~1，示例：0.95' },
    { key: 'lowThreshold', label: '低电量门限', placeholder: '0~1，示例：0.02' },
    { key: 'envStandbyCurrent', label: '环境记忆准入功耗 mA', placeholder: '示例：14' },
    { key: 'wakeCurrent', label: '自然式唤醒准入功耗 mA', placeholder: '示例：67' },
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
    return '<section class="panel rounded-3xl es-card">'
      + '<header class="es-card-head"><h3><span class="material-symbols-outlined">battery_full</span>电池参数</h3>'
      + '<small>续航时长基于这些参数自动计算，修改后实时保存</small></header>'
      + '<div class="es-battery-grid">' + fields + '</div>'
      + '<p class="es-battery-derate">综合可用电量系数 = 可放电比例 × (1 − 低电量门限) = <b>' + fmt(derate, 4) + '</b></p>'
      + '</section>';
  }

  const saveBattery = debounce((project) => E.store.save(), 150);

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
        saveBattery(project);
        const derateEl = root.querySelector('.es-battery-derate b');
        if (derateEl) {
          const ratio = num(project.battery.dischargeRatio);
          const threshold = num(project.battery.lowThreshold);
          const derate = (ratio === null ? 1 : ratio) * (1 - (threshold === null ? 0 : threshold));
          derateEl.textContent = fmt(derate, 4);
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
      // 交给对比模块自己的静默刷新（内部会先销毁旧图表，避免 Canvas 泄漏）
      window.ES_ENDURANCE_UI.refreshCompareQuiet && window.ES_ENDURANCE_UI.refreshCompareQuiet();
    }
  }, 220);

  window.ES_ENDURANCE_UI = { render, refreshModelAndCompare };
})();

/* ================================================================
 * 第三部分：场景电流台账（日期 / 版本 / 电流，可更新与追溯历史）
 * ================================================================ */
(function () {
  'use strict';
  const E = window.ES_ENDURANCE;
  const { escapeHtml, uid, num, fmt, todayStr } = E;
  const root = document.getElementById('enduranceStatusRoot');

  function renderScenarioCard(project) {
    const rows = project.scenarios.map((scenario) => {
      const rec = E.latestRecord(scenario);
      const current = rec ? fmt(num(rec.current), 3) : '—';
      const meta = rec ? escapeHtml(rec.date + ' · ' + (rec.version || '未填版本')) : '暂无数据';
      return '<tr data-es-scenario="' + escapeHtml(scenario.id) + '">'
        + '<td class="es-sc-name">' + escapeHtml(scenario.name)
        + (scenario.remark ? '<small>' + escapeHtml(scenario.remark) + '</small>' : '') + '</td>'
        + '<td class="es-sc-current">' + current + ' mA</td>'
        + '<td class="es-sc-meta">' + meta + '</td>'
        + '<td class="es-sc-ops">'
        + '<button type="button" class="es-mini-btn" data-es-action="update-current" title="更新电流"><span class="material-symbols-outlined">edit</span>更新</button>'
        + '<button type="button" class="es-mini-btn" data-es-action="history" title="历史记录"><span class="material-symbols-outlined">history</span>历史</button>'
        + '<button type="button" class="es-mini-btn es-mini-danger" data-es-action="delete-scenario" title="删除场景"><span class="material-symbols-outlined">close</span></button>'
        + '</td></tr>';
    }).join('');

    return '<section class="panel rounded-3xl es-card">'
      + '<header class="es-card-head"><h3><span class="material-symbols-outlined">bolt</span>场景电流台账</h3>'
      + '<button type="button" class="template-btn template-btn-sm es-add-scenario-btn"><span class="material-symbols-outlined text-sm">add</span>新增场景</button></header>'
      + '<div class="es-table-wrap"><table class="es-table"><thead><tr>'
      + '<th>场景</th><th>当前电流</th><th>最近更新（日期 · 版本）</th><th class="es-col-ops">操作</th>'
      + '</tr></thead><tbody>'
      + (rows || '<tr><td colspan="4" class="es-table-empty">还没有场景，点右上角「新增场景」开始维护；场景电流按日期/版本留痕，每次更新都会保留历史</td></tr>')
      + '</tbody></table></div></section>';
  }

  function bindScenarios(project) {
    const addBtn = root.querySelector('.es-add-scenario-btn');
    if (addBtn) addBtn.addEventListener('click', () => openScenarioEditor(project, null));

    root.querySelectorAll('[data-es-scenario]').forEach((tr) => {
      const id = tr.dataset.esScenario;
      tr.querySelector('[data-es-action="update-current"]').addEventListener('click', () => openCurrentEditor(project, id));
      tr.querySelector('[data-es-action="history"]').addEventListener('click', () => openHistory(project, id));
      tr.querySelector('[data-es-action="delete-scenario"]').addEventListener('click', () => {
        const scenario = project.scenarios.find((s) => s.id === id);
        if (!scenario) return;
        if (!window.confirm('确认删除场景「' + scenario.name + '」及其全部电流记录？续航模型中引用该场景的行需要手动调整。')) return;
        project.scenarios = project.scenarios.filter((s) => s.id !== id);
        project.models.forEach((model) => {
          model.rows = model.rows.filter((r) => r.scenarioId !== id);
        });
        E.store.save();
        window.ES_ENDURANCE_UI.render();
      });
    });
  }

  // 新增 / 重命名场景
  function openScenarioEditor(project, scenario) {
    const name = window.prompt(scenario ? '修改场景名称：' : '请输入新场景名称：', scenario ? scenario.name : '');
    if (name === null) return;
    const trimmed = name.trim();
    if (!trimmed) { alert('场景名称不能为空'); return; }
    if (scenario) {
      scenario.name = trimmed;
    } else {
      const created = { id: uid('sc'), name: trimmed, remark: '', records: [] };
      project.scenarios.push(created);
      // 同步给该项目的每个模型加一行（默认时长 0）
      project.models.forEach((model) => {
        const durations = {};
        model.modes.forEach((mode) => { durations[mode.id] = 0; });
        model.rows.push({ scenarioId: created.id, durations: durations });
      });
    }
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  // 电流更新弹窗：日期 / 版本 / 电流，确认后追加一条记录（历史不覆盖）
  function openCurrentEditor(project, scenarioId) {
    const scenario = project.scenarios.find((s) => s.id === scenarioId);
    if (!scenario) return;
    const last = E.latestRecord(scenario);
    const date = window.prompt('更新日期（YYYY-MM-DD）：', last ? last.date : todayStr());
    if (date === null) return;
    const dateTrim = date.trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateTrim)) { alert('日期格式应为 YYYY-MM-DD'); return; }
    const version = window.prompt('版本号：', last ? (last.version || '') : '');
    if (version === null) return;
    const currentText = window.prompt('场景电流（mA）：', last ? String(last.current) : '');
    if (currentText === null) return;
    const current = num(currentText);
    if (current === null || current < 0) { alert('电流需为不小于 0 的数字'); return; }
    scenario.records.push({
      id: uid('rec'), date: dateTrim, version: version.trim(),
      current: current, note: '',
    });
    scenario.records.sort((a, b) => String(a.date).localeCompare(String(b.date)));
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  // 历史记录查看
  function openHistory(project, scenarioId) {
    const scenario = project.scenarios.find((s) => s.id === scenarioId);
    if (!scenario || !scenario.records.length) { alert('该场景还没有电流记录'); return; }
    const lines = scenario.records.slice().reverse().map((r, i) => {
      const tag = i === 0 ? '（当前生效）' : '';
      return (r.date || '—') + ' | ' + (r.version || '未填版本') + ' | ' + fmt(num(r.current), 3) + ' mA' + tag;
    });
    window.alert('「' + scenario.name + '」电流更新历史（共 ' + scenario.records.length + ' 条）：\n\n' + lines.join('\n'));
  }

  window.ES_ENDURANCE_SCENARIO = { renderScenarioCard, bindScenarios };
})();

/* ================================================================
 * 第四部分：自定义续航模型（模式管理 + 场景×时长矩阵 + 自动计算）
 * ================================================================ */
(function () {
  'use strict';
  const E = window.ES_ENDURANCE;
  const { escapeHtml, num, fmt, uid } = E;
  const root = document.getElementById('enduranceStatusRoot');

  function activeModel(project) {
    return project.models.find((m) => m.id === project.activeModelId) || project.models[0] || null;
  }

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

  function renderModelBody(project, model) {
    const stats = E.computeModelStats(project, model);
    // 两行表头：模式名（跨两列）+ 时长/耗电量
    const headTop = '<th class="es-corner" rowspan="2">场景</th><th rowspan="2">电流 mA</th>'
      + model.modes.map((mode) => {
        const tag = mode.fullOnly ? '<em class="es-mode-tag">满电</em>' : '';
        const win = mode.window ? '<small>' + escapeHtml(mode.window) + '</small>' : '';
        const need = (mode.needMah !== null && mode.needMah !== undefined && mode.needMah !== '')
          ? '<small>目标 ' + fmt(num(mode.needMah), 0) + ' mAh</small>' : '';
        const note = mode.note ? '<small class="es-mode-note">' + escapeHtml(mode.note) + '</small>' : '';
        const ops = '<span class="es-mode-ops">'
          + '<button type="button" class="es-icon-btn" data-es-mode-edit="' + escapeHtml(mode.id) + '" title="编辑使用模式"><span class="material-symbols-outlined">edit</span></button>'
          + '<button type="button" class="es-icon-btn es-icon-danger" data-es-mode-delete="' + escapeHtml(mode.id) + '" title="删除使用模式"><span class="material-symbols-outlined">delete</span></button>'
          + '</span>';
        return '<th class="es-mode-head" colspan="2"><div class="es-mode-title">' + escapeHtml(mode.name) + tag + ops + '</div>' + win + need + note + '</th>';
      }).join('');
    const headSub = model.modes.map(() => '<th class="es-sub">时长 h</th><th class="es-sub">耗电量 mAh</th>').join('');

    const body = project.scenarios.map((scenario) => {
      const row = model.rows.find((r) => r.scenarioId === scenario.id);
      const current = E.scenarioCurrent(project, scenario.id);
      const cells = model.modes.map((mode) => {
        const hours = row ? num(row.durations[mode.id]) : null;
        const mah = (hours !== null && current !== null && hours !== null) ? hours * current : null;
        return '<td class="es-duration-cell"><input type="text" inputmode="decimal" class="es-duration-input"'
          + ' data-es-scenario-id="' + escapeHtml(scenario.id) + '" data-es-mode-id="' + escapeHtml(mode.id) + '"'
          + ' value="' + (hours === null ? '' : hours) + '" placeholder="0" autocomplete="off" /></td>'
          + '<td class="es-mah-cell">' + (mah === null ? '—' : fmt(mah, 2)) + '</td>';
      }).join('');
      return '<tr><td class="es-sc-name">' + escapeHtml(scenario.name) + '</td>'
        + '<td class="es-current-cell">' + (current === null ? '—' : fmt(current, 2)) + '</td>' + cells + '</tr>';
    }).join('');

    // 汇总行（每个模式两列：时长合计、耗电量合计放时长列下方对齐）
    // 满电模式只展示满电使用时长；普通模式展示带可用电量折减的 DOU（Typ/Min）
    const statRow = (label, pick, digits) => '<tr class="es-stat-row"><td>' + label + '</td><td></td>'
      + model.modes.map((mode) => {
        const s = stats[mode.id] || {};
        const v = pick(mode, s);
        const dText = (v.hours === undefined || v.hours === '' || v.hours === null) ? '' : fmt(v.hours, digits);
        const mText = (v.mah === undefined || v.mah === '' || v.mah === null) ? '' : fmt(v.mah, digits);
        return '<td class="es-stat-duration">' + dText + '</td>'
          + '<td class="es-stat-mah">' + mText + '</td>';
      }).join('') + '</tr>';

    const totals = statRow('总时长 / 总耗电量', (mode, s) => ({ hours: s.totalHours, mah: s.totalMah }), 2)
      + statRow('平均电流 mA', (mode, s) => ({ hours: '', mah: s.avgCurrent }), 2)
      + statRow('满电使用时长 h（满电模式）', (mode, s) => ({ hours: mode.fullOnly ? s.fullHours : null, mah: '' }), 2)
      + statRow('DOU（Typ）h', (mode, s) => ({ hours: mode.fullOnly ? null : s.douTyp, mah: '' }), 2)
      + statRow('DOU（Min）h', (mode, s) => ({ hours: mode.fullOnly ? null : s.douMin, mah: '' }), 2);

    return '<div class="es-mode-toolbar">'
      + '<button type="button" class="es-mini-btn" data-es-model-action="rename"><span class="material-symbols-outlined">drive_file_rename_outline</span>重命名模型</button>'
      + '<button type="button" class="es-mini-btn" data-es-model-action="add-mode"><span class="material-symbols-outlined">add</span>新增使用模式</button>'
      + '<button type="button" class="es-mini-btn es-mini-danger" data-es-model-action="delete-model"><span class="material-symbols-outlined">delete</span>删除模型</button>'
      + '</div>'
      + '<div class="es-model-table-wrap"><table class="es-model-table"><thead><tr>' + headTop + '</tr><tr>' + headSub + '</tr></thead>'
      + '<tbody>' + (body || '<tr><td colspan="' + (2 + model.modes.length * 2) + '" class="es-table-empty">请先在上方「场景电流台账」新增场景</td></tr>') + '</tbody>'
      + '<tfoot>' + totals + '</tfoot></table></div>'
      + '<p class="es-model-tip"><span class="material-symbols-outlined">info</span>时长可直接编辑并自动保存；耗电量 = 时长 × 场景当前电流，DOU = 电池容量 Typ/Min × 可放电比例 × (1 − 低电量门限) ÷ 平均电流。</p>';
  }

  function bindModels(project) {
    const quietRefreshCompare = E.debounce(() => {
      window.ES_ENDURANCE_UI.refreshCompareQuiet && window.ES_ENDURANCE_UI.refreshCompareQuiet();
    }, 220);
    // 模型切换
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
    // 模式表头内的编辑 / 删除使用模式
    root.querySelectorAll('[data-es-mode-edit]').forEach((btn) => {
      btn.addEventListener('click', () => openModeEditor(project, activeModel(project), btn.dataset.esModeEdit));
    });
    root.querySelectorAll('[data-es-mode-delete]').forEach((btn) => {
      btn.addEventListener('click', () => deleteMode(project, activeModel(project), btn.dataset.esModeDelete));
    });
    // 时长矩阵编辑（事件委托，就地更新 mAh 与汇总，不整卡重绘以保留输入焦点）
    const tableWrap = root.querySelector('.es-model-table-wrap');
    if (tableWrap) {
      tableWrap.addEventListener('input', (event) => {
        const input = event.target.closest ? event.target.closest('.es-duration-input') : null;
        if (!input) return;
        const model = activeModel(project);
        if (!model) return;
        const row = model.rows.find((r) => r.scenarioId === input.dataset.esScenarioId);
        if (!row) return;
        const raw = input.value.trim();
        const parsed = raw === '' ? '' : num(raw);
        // 非法或负数输入只标红，不写入数据，避免汇总出现负值 / NaN
        if (raw !== '' && (parsed === null || parsed < 0)) {
          input.classList.add('is-invalid');
          return;
        }
        input.classList.remove('is-invalid');
        row.durations[input.dataset.esModeId] = raw;
        E.store.save();
        updateModelCells(project, model);
        quietRefreshCompare();
      });
    }
  }

  // 就地刷新 mAh 与 tfoot 汇总
  function updateModelCells(project, model) {
    const stats = E.computeModelStats(project, model);
    const tbody = root.querySelector('.es-model-table tbody');
    if (tbody) {
      tbody.querySelectorAll('tr').forEach((tr) => {
        const scenarioId = tr.querySelector('.es-duration-input')?.dataset.esScenarioId;
        if (!scenarioId) return;
        const current = E.scenarioCurrent(project, scenarioId);
        model.modes.forEach((mode, i) => {
          const row = model.rows.find((r) => r.scenarioId === scenarioId);
          const hours = row ? num(row.durations[mode.id]) : null;
          const mah = (hours !== null && current !== null) ? hours * current : null;
          const mahCell = tr.querySelectorAll('.es-mah-cell')[i];
          if (mahCell) mahCell.textContent = mah === null ? '—' : fmt(mah, 2);
        });
      });
    }
    const tfoot = root.querySelector('.es-model-table tfoot');
    if (!tfoot) return;
    const setFoot = (rowIndex, pick, digits) => {
      const tr = tfoot.querySelectorAll('tr')[rowIndex];
      if (!tr) return;
      model.modes.forEach((mode, i) => {
        const s = stats[mode.id] || {};
        const v = pick(mode, s);
        const cells = tr.querySelectorAll('td');
        // tfoot 每行前两格是行名与占位空列，模式数据从第 3 格开始，每模式占两格
        const dCell = cells[2 + i * 2];
        const mCell = cells[3 + i * 2];
        if (dCell) dCell.textContent = (v.hours === undefined || v.hours === '' || v.hours === null) ? '' : fmt(v.hours, digits);
        if (mCell) mCell.textContent = (v.mah === undefined || v.mah === '' || v.mah === null) ? '' : fmt(v.mah, digits);
      });
    };
    setFoot(0, (mode, s) => ({ hours: s.totalHours, mah: s.totalMah }), 2);
    setFoot(1, (mode, s) => ({ hours: '', mah: s.avgCurrent }), 2);
    setFoot(2, (mode, s) => ({ hours: mode.fullOnly ? s.fullHours : null, mah: '' }), 2);
    setFoot(3, (mode, s) => ({ hours: mode.fullOnly ? null : s.douTyp, mah: '' }), 2);
    setFoot(4, (mode, s) => ({ hours: mode.fullOnly ? null : s.douMin, mah: '' }), 2);
  }

  function handleModelAction(project, action) {
    const model = activeModel(project);
    if (action === 'add') {
      const name = window.prompt('新模型名称：', '续航模型 ' + (project.models.length + 1));
      if (name === null) return;
      const trimmed = name.trim();
      if (!trimmed) { alert('模型名称不能为空'); return; }
      const created = createModel(project, trimmed, false);
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
      const copy = JSON.parse(JSON.stringify(model));
      copy.id = uid('m');
      copy.name = trimmed;
      copy.modes.forEach((mode) => { mode.id = uid('md'); });
      const idMap = {};
      model.modes.forEach((old, i) => { idMap[old.id] = copy.modes[i].id; });
      copy.rows = copy.rows.map((r) => {
        const durations = {};
        Object.keys(r.durations).forEach((oldId) => { durations[idMap[oldId]] = r.durations[oldId]; });
        return { scenarioId: r.scenarioId, durations: durations };
      });
      project.models.push(copy);
      project.activeModelId = copy.id;
      E.store.save();
      window.ES_ENDURANCE_UI.render();
      return;
    }
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
    if (action === 'add-mode') {
      const name = window.prompt('新使用模式名称（如：中度使用）：', '');
      if (name === null) return;
      const trimmed = name.trim();
      if (!trimmed) { alert('模式名称不能为空'); return; }
      const win = window.prompt('使用时间窗口（可留空，如 8am~8pm）：', '');
      if (win === null) return;
      const needText = window.prompt('目标耗电量 mAh（可留空，满电模式通常不设）：', '');
      if (needText === null) return;
      const need = num(needText);
      if (needText.trim() !== '' && (need === null || need < 0)) { alert('目标耗电量需为不小于 0 的数字'); return; }
      const note = window.prompt('模式备注（可留空）：', '');
      if (note === null) return;
      const fullOnly = window.confirm('该模式是否为「满电使用」模式？\n\n确定 = 满电使用（不设目标耗电量，以满电时长为准）\n取消 = 普通使用模式');
      const mode = {
        id: uid('md'), name: trimmed, window: win.trim(),
        needMah: needText.trim() === '' ? null : need, note: note.trim(), fullOnly: !!fullOnly,
      };
      model.modes.push(mode);
      project.scenarios.forEach((scenario) => {
        let row = model.rows.find((r) => r.scenarioId === scenario.id);
        if (!row) { row = { scenarioId: scenario.id, durations: {} }; model.rows.push(row); }
        row.durations[mode.id] = 0;
      });
      E.store.save();
      window.ES_ENDURANCE_UI.render();
    }
  }

  // 编辑使用模式：名称 / 时间窗口 / 目标耗电量 / 备注 / 是否满电模式
  function openModeEditor(project, model, modeId) {
    if (!model) return;
    const mode = model.modes.find((md) => md.id === modeId);
    if (!mode) return;
    const name = window.prompt('使用模式名称：', mode.name);
    if (name === null) return;
    const trimmed = name.trim();
    if (!trimmed) { alert('模式名称不能为空'); return; }
    const win = window.prompt('使用时间窗口（可留空，如 8am~8pm）：', mode.window || '');
    if (win === null) return;
    const needText = window.prompt('目标耗电量 mAh（可留空，满电模式通常不设）：',
      mode.needMah === null || mode.needMah === undefined ? '' : String(mode.needMah));
    if (needText === null) return;
    const need = num(needText);
    if (needText.trim() !== '' && (need === null || need < 0)) { alert('目标耗电量需为不小于 0 的数字'); return; }
    const note = window.prompt('模式备注（可留空）：', mode.note || '');
    if (note === null) return;
    const fullOnly = window.confirm('该模式是否为「满电使用」模式？\n\n确定 = 满电使用（不设目标耗电量，以满电时长为准）\n取消 = 普通使用模式');
    mode.name = trimmed;
    mode.window = win.trim();
    mode.needMah = needText.trim() === '' ? null : need;
    mode.note = note.trim();
    mode.fullOnly = !!fullOnly;
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  // 删除使用模式：同步清理各场景行里的时长字段，每个模型至少保留一个模式
  function deleteMode(project, model, modeId) {
    if (!model) return;
    const mode = model.modes.find((md) => md.id === modeId);
    if (!mode) return;
    if (model.modes.length <= 1) { alert('每个续航模型至少保留一个使用模式'); return; }
    if (!window.confirm('确认删除使用模式「' + mode.name + '」？\n该模式下各场景的时长配置会一并删除。')) return;
    model.modes = model.modes.filter((md) => md.id !== modeId);
    model.rows.forEach((row) => { delete row.durations[modeId]; });
    E.store.save();
    window.ES_ENDURANCE_UI.render();
  }

  // 新建空模型：自动带入项目已有场景，模式给一个默认「典型使用」
  function createModel(project, name) {
    const mode = { id: uid('md'), name: '典型使用', window: '8am~8pm', needMah: null, note: '', fullOnly: false };
    const model = { id: uid('m'), name: name, modes: [mode], rows: [] };
    project.scenarios.forEach((scenario) => {
      model.rows.push({ scenarioId: scenario.id, durations: { [mode.id]: 0 } });
    });
    project.models.push(model);
    return model;
  }

  window.ES_ENDURANCE_MODEL = { renderModelCard, bindModels, activeModel };
})();

/* ================================================================
 * 第五部分：项目横向对比（场景电流 & DOU 柱状图，默认全量对比）
 * ================================================================ */
(function () {
  'use strict';
  const E = window.ES_ENDURANCE;
  const { escapeHtml, num, fmt } = E;
  const root = document.getElementById('enduranceStatusRoot');

  // 对比选择只保存在本次会话；首次进入默认全量，之后尊重用户的清空 / 勾选结果
  const compareState = {
    projects: [], scenarioNames: [], seenProjects: [],
    projectsInit: false, scenariosInit: false, douMetric: 'Typ',
  };

  // 跨项目场景池：按场景名取并集（跨项目同名场景视为同一对比维度）
  function allScenarioNames(data) {
    const names = [];
    data.projects.forEach((p) => {
      p.scenarios.forEach((s) => { if (names.indexOf(s.name) === -1) names.push(s.name); });
    });
    return names;
  }

  function renderCompareCard() {
    const data = E.store.load();
    if (data.projects.length < 1) return '';
    if (!compareState.projectsInit) {
      compareState.projects = data.projects.map((p) => p.id);
      compareState.seenProjects = data.projects.map((p) => p.id);
      compareState.projectsInit = true;
    }
    // 新建的项目默认纳入对比（seen 列表区分“新出现”与“用户手动取消勾选”）
    data.projects.forEach((p) => {
      if (compareState.seenProjects.indexOf(p.id) === -1) {
        compareState.seenProjects.push(p.id);
        if (compareState.projects.indexOf(p.id) === -1) compareState.projects.push(p.id);
      }
    });
    const names = allScenarioNames(data);
    if (!compareState.scenariosInit) {
      compareState.scenarioNames = names.slice();
      compareState.scenariosInit = true;
    }

    const projectChips = data.projects.map((p) => {
      const on = compareState.projects.indexOf(p.id) !== -1;
      return '<button type="button" class="es-pick-chip' + (on ? ' is-on' : '') + '" data-es-pick-project="' + escapeHtml(p.id) + '">'
        + '<span class="material-symbols-outlined">' + (on ? 'check_box' : 'check_box_outline_blank') + '</span>'
        + escapeHtml(p.name) + '</button>';
    }).join('');
    const scenarioChips = names.length ? names.map((name) => {
      const on = compareState.scenarioNames.indexOf(name) !== -1;
      return '<button type="button" class="es-pick-chip es-pick-chip-sm' + (on ? ' is-on' : '') + '" data-es-pick-scenario="' + escapeHtml(name) + '">'
        + '<span class="material-symbols-outlined">' + (on ? 'check_box' : 'check_box_outline_blank') + '</span>'
        + escapeHtml(name) + '</button>';
    }).join('') : '<span class="es-pick-empty">所有项目都还没有场景，先在「场景电流台账」中新增</span>';

    return '<section class="panel rounded-3xl es-card es-compare-card">'
      + '<header class="es-card-head"><h3><span class="material-symbols-outlined">bar_chart</span>项目横向对比</h3>'
      + '<div class="es-dou-toggle" role="group" aria-label="DOU 口径">'
      + '<button type="button" class="es-mini-btn' + (compareState.douMetric === 'Typ' ? ' is-on' : '') + '" data-es-dou="Typ">DOU（Typ）</button>'
      + '<button type="button" class="es-mini-btn' + (compareState.douMetric === 'Min' ? ' is-on' : '') + '" data-es-dou="Min">DOU（Min）</button>'
      + '</div></header>'
      + '<div class="es-pick-group"><span class="es-pick-label">对比项目'
      + '<button type="button" class="es-link-btn" data-es-pick-all="project">全选</button>'
      + '<button type="button" class="es-link-btn" data-es-pick-clear="project">清空</button></span>'
      + '<div class="es-pick-chips">' + projectChips + '</div></div>'
      + '<div class="es-pick-group"><span class="es-pick-label">对比场景（默认全量）'
      + '<button type="button" class="es-link-btn" data-es-pick-all="scenario">全选</button>'
      + '<button type="button" class="es-link-btn" data-es-pick-clear="scenario">清空</button></span>'
      + '<div class="es-pick-chips">' + scenarioChips + '</div></div>'
      + '<div class="es-chart-grid">'
      + '<div class="es-chart-box"><h4>场景电流对比（mA）</h4><div class="es-canvas-wrap"><canvas id="esCurrentChart"></canvas></div></div>'
      + '<div class="es-chart-box"><h4>DOU 对比（h，取各项目当前激活模型）</h4><div class="es-canvas-wrap"><canvas id="esDouChart"></canvas></div></div>'
      + '</div></section>';
  }

  function bindCompare() {
    root.querySelectorAll('[data-es-pick-project]').forEach((btn) => {
      btn.addEventListener('click', () => togglePick(compareState.projects, btn.dataset.esPickProject));
    });
    root.querySelectorAll('[data-es-pick-scenario]').forEach((btn) => {
      btn.addEventListener('click', () => togglePick(compareState.scenarioNames, btn.dataset.esPickScenario));
    });
    root.querySelectorAll('[data-es-pick-all]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const data = E.store.load();
        if (btn.dataset.esPickAll === 'project') compareState.projects = data.projects.map((p) => p.id);
        else compareState.scenarioNames = allScenarioNames(data);
        softRerender();
      });
    });
    root.querySelectorAll('[data-es-pick-clear]').forEach((btn) => {
      btn.addEventListener('click', () => {
        if (btn.dataset.esPickClear === 'project') compareState.projects = [];
        else compareState.scenarioNames = [];
        softRerender();
      });
    });
    root.querySelectorAll('[data-es-dou]').forEach((btn) => {
      btn.addEventListener('click', () => { compareState.douMetric = btn.dataset.esDou; softRerender(); });
    });
    drawCharts();
  }

  function togglePick(storeList, id) {
    const i = storeList.indexOf(id);
    if (i === -1) storeList.push(id); else storeList.splice(i, 1);
    softRerender();
  }

  // 只重绘对比卡片，保留其他卡片的输入焦点
  function softRerender() {
    const card = root.querySelector('.es-compare-card');
    if (!card) return;
    E.destroyCharts();
    const next = document.createElement('div');
    next.innerHTML = renderCompareCard();
    card.replaceWith(next.firstElementChild);
    bindCompare();
  }

  function baseOptions(yTitle) {
    return {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } },
        tooltip: { callbacks: { label: (ctx) => ctx.dataset.label + '：' + (ctx.parsed.y === null ? '无数据' : fmt(ctx.parsed.y, 2)) } },
      },
      scales: {
        x: { ticks: { font: { size: 10 }, maxRotation: 35, minRotation: 0 }, grid: { display: false } },
        y: { beginAtZero: true, title: { display: true, text: yTitle, font: { size: 11 } }, grid: { color: 'rgba(148,163,184,.18)' } },
      },
    };
  }

  function drawCharts() {
    const data = E.store.load();
    const projects = data.projects.filter((p) => compareState.projects.indexOf(p.id) !== -1);
    // 场景筛选按名称生效，跨项目同名场景在同一柱组对比
    const scenarios = allScenarioNames(data)
      .filter((name) => compareState.scenarioNames.indexOf(name) !== -1)
      .map((name) => ({ name: name }));

    drawCurrentChart(projects, scenarios);
    drawDouChart(projects);
  }

  function drawCurrentChart(projects, scenarios) {
    const canvas = document.getElementById('esCurrentChart');
    if (!canvas || typeof Chart === 'undefined') return;
    const labels = scenarios.map((s) => s.name);
    const datasets = projects.map((p, i) => ({
      label: p.name,
      backgroundColor: E.CHART_COLORS[i % E.CHART_COLORS.length],
      borderRadius: 4,
      data: scenarios.map((s) => {
        const target = p.scenarios.find((x) => x.name === s.name);
        if (!target) return null;
        const rec = E.latestRecord(target);
        return rec ? E.num(rec.current) : null;
      }),
    }));
    const chart = new Chart(canvas, {
      type: 'bar',
      data: { labels: labels, datasets: datasets },
      options: baseOptions('mA'),
    });
    E.chartInstances && E.chartInstances.push(chart);
  }

  function drawDouChart(projects) {
    const canvas = document.getElementById('esDouChart');
    if (!canvas || typeof Chart === 'undefined') return;
    // 以各项目激活模型的模式名并集作为数据集
    const modeNames = [];
    projects.forEach((p) => {
      const model = p.models.find((m) => m.id === p.activeModelId) || p.models[0];
      if (!model) return;
      model.modes.forEach((mode) => { if (modeNames.indexOf(mode.name) === -1) modeNames.push(mode.name); });
    });
    const labels = projects.map((p) => p.name);
    const datasets = modeNames.map((modeName, i) => ({
      label: modeName,
      backgroundColor: E.CHART_COLORS[i % E.CHART_COLORS.length],
      borderRadius: 4,
      data: projects.map((p) => {
        const model = p.models.find((m) => m.id === p.activeModelId) || p.models[0];
        if (!model) return null;
        const mode = model.modes.find((m) => m.name === modeName);
        if (!mode) return null;
        const stats = E.computeModeStats(p, model, mode);
        return compareState.douMetric === 'Min' ? stats.douMin : stats.douTyp;
      }),
    }));
    const chart = new Chart(canvas, {
      type: 'bar',
      data: { labels: labels, datasets: datasets },
      options: baseOptions('小时'),
    });
    E.chartInstances && E.chartInstances.push(chart);
  }

  // 暴露给主控制器与模型编辑后的静默刷新
  window.ES_ENDURANCE_UI.renderCompareCard = renderCompareCard;
  window.ES_ENDURANCE_UI.bindCompare = bindCompare;
  window.ES_ENDURANCE_UI.refreshCompareQuiet = softRerender;
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
