const LETTER_CODES = Array.from({ length: 26 }, (_, index) => String.fromCharCode(65 + index));
const ACTION_CODES = Array.from({ length: 100 }, (_, index) => String(index).padStart(2, "0"));
const STORAGE_KEY = "auto-instrument-rhythm-designer";
const MODULE_COLORS = [
  "#8ec5ff",
  "#f6c879",
  "#90d6a3",
  "#f4a3a3",
  "#c8b6ff",
  "#8fe3d8",
  "#f4b77f",
  "#c1d96f"
];

const initialState = {
  metadata: {
    modes: [
      { code: "A", name: "放卡" },
      { code: "B", name: "一步法" },
      { code: "C", name: "两步一稀释（单次稀释）" },
      { code: "D", name: "两步一稀释（多次稀释）" },
      { code: "E", name: "两步两稀释" }
    ],
    modules: [
      { code: "A", name: "移液器" },
      { code: "B", name: "摆渡车" },
      { code: "C", name: "孵育仓" },
      { code: "D", name: "检测" }
    ],
    actions: [
      { moduleCode: "A", code: "00", name: "取TIP" },
      { moduleCode: "A", code: "01", name: "弃TIP" },
      { moduleCode: "A", code: "02", name: "取稀释液1" },
      { moduleCode: "A", code: "03", name: "取稀释液2" },
      { moduleCode: "B", code: "00", name: "送样本架" },
      { moduleCode: "C", code: "00", name: "恒温孵育" },
      { moduleCode: "D", code: "00", name: "光学检测" }
    ]
  },
  editor: {
    modeCode: "B",
    beatDuration: 30,
    rows: [
      { id: crypto.randomUUID(), moduleCode: "A", actionCode: "00", startTime: 0, previousActionId: "", duration: 2.5, expanded: true },
      { id: crypto.randomUUID(), moduleCode: "A", actionCode: "02", startTime: 2.5, previousActionId: "", duration: 3 },
      { id: crypto.randomUUID(), moduleCode: "D", actionCode: "00", startTime: 8, previousActionId: "", duration: 4 }
    ]
  }
};

const state = loadState();
let activeEdit = {
  modeCode: null,
  moduleCode: null,
  actionKey: null
};

const elements = {
  modeCount: document.getElementById("modeCount"),
  moduleCount: document.getElementById("moduleCount"),
  actionCount: document.getElementById("actionCount"),
  modeCodeInput: document.getElementById("modeCodeInput"),
  modeNameInput: document.getElementById("modeNameInput"),
  moduleCodeInput: document.getElementById("moduleCodeInput"),
  moduleNameInput: document.getElementById("moduleNameInput"),
  actionModuleInput: document.getElementById("actionModuleInput"),
  actionCodeInput: document.getElementById("actionCodeInput"),
  actionNameInput: document.getElementById("actionNameInput"),
  modeList: document.getElementById("modeList"),
  moduleList: document.getElementById("moduleList"),
  actionList: document.getElementById("actionList"),
  storageStatus: document.getElementById("storageStatus"),
  fileLoader: document.getElementById("fileLoader"),
  editorModeSelect: document.getElementById("editorModeSelect"),
  beatDurationInput: document.getElementById("beatDurationInput"),
  timelineWrap: document.getElementById("timelineWrap"),
  timelineTitle: document.getElementById("timelineTitle"),
  timelineSubtitle: document.getElementById("timelineSubtitle")
};

bindEvents();
renderAll();

function bindEvents() {
  document.querySelectorAll(".tab-button").forEach((button) => {
    button.addEventListener("click", () => switchTab(button.dataset.tabTarget));
  });

  document.getElementById("saveModeBtn").addEventListener("click", saveMode);
  document.getElementById("saveModuleBtn").addEventListener("click", saveModule);
  document.getElementById("saveActionBtn").addEventListener("click", saveAction);
  document.getElementById("addModeBtn").addEventListener("click", resetModeForm);
  document.getElementById("addModuleBtn").addEventListener("click", resetModuleForm);
  document.getElementById("addActionBtn").addEventListener("click", resetActionForm);
  document.getElementById("loadBtn").addEventListener("click", () => elements.fileLoader.click());
  document.getElementById("editorLoadBtn").addEventListener("click", () => elements.fileLoader.click());
  document.getElementById("saveBtn").addEventListener("click", handleSoftSave);
  document.getElementById("saveAsBtn").addEventListener("click", handleSaveAs);
  document.getElementById("clearBtn").addEventListener("click", clearAllData);
  document.getElementById("generateTimelineBtn").addEventListener("click", generateTimeline);
  document.getElementById("addTimelineRowBtn").addEventListener("click", addTimelineRow);
  elements.fileLoader.addEventListener("change", handleFileImport);
  elements.editorModeSelect.addEventListener("change", (event) => {
    state.editor.modeCode = event.target.value;
    persist("模式已切换。");
    renderEditorHeader();
  });
  elements.beatDurationInput.addEventListener("change", (event) => {
    state.editor.beatDuration = normalizeNumber(event.target.value, 1, 999, 30, false);
    persist("节拍时间已更新。");
    renderEditorHeader();
  });
}

function renderAll() {
  populateSelect(elements.modeCodeInput, LETTER_CODES, activeEdit.modeCode);
  populateSelect(elements.moduleCodeInput, LETTER_CODES, activeEdit.moduleCode);
  populateSelect(elements.actionCodeInput, ACTION_CODES);
  renderActionModuleSelect();
  renderCounts();
  renderLists();
  renderEditorHeader();
  renderTimeline();
}

function renderCounts() {
  elements.modeCount.textContent = state.metadata.modes.length;
  elements.moduleCount.textContent = state.metadata.modules.length;
  elements.actionCount.textContent = state.metadata.actions.length;
}

function renderLists() {
  renderList(elements.modeList, state.metadata.modes, ({ code, name }) => ({
    code,
    name,
    onEdit: () => {
      activeEdit.modeCode = code;
      elements.modeCodeInput.value = code;
      elements.modeNameInput.value = name;
    },
    onDelete: () => {
      if (!window.confirm(`确认删除模式 ${code} ${name} 吗？`)) {
        return;
      }
      state.metadata.modes = state.metadata.modes.filter((item) => item.code !== code);
      if (state.editor.modeCode === code) {
        state.editor.modeCode = state.metadata.modes[0]?.code || "";
      }
      persist(`已删除模式 ${code}。`);
      resetModeForm();
      renderAll();
    }
  }));

  renderList(elements.moduleList, state.metadata.modules, ({ code, name }) => ({
    code,
    name,
    onEdit: () => {
      activeEdit.moduleCode = code;
      elements.moduleCodeInput.value = code;
      elements.moduleNameInput.value = name;
    },
    onDelete: () => {
      const usedByAction = state.metadata.actions.some((action) => action.moduleCode === code);
      if (usedByAction) {
        window.alert("该模块下仍存在动作，请先删除对应动作。");
        return;
      }
      if (!window.confirm(`确认删除模块 ${code} ${name} 吗？`)) {
        return;
      }
      state.metadata.modules = state.metadata.modules.filter((item) => item.code !== code);
      state.editor.rows = state.editor.rows.filter((row) => row.moduleCode !== code);
      persist(`已删除模块 ${code}。`);
      resetModuleForm();
      renderAll();
    }
  }));

  renderList(elements.actionList, state.metadata.actions, ({ moduleCode, code, name }) => ({
    code: `${moduleCode}${code}`,
    name,
    onEdit: () => {
      activeEdit.actionKey = `${moduleCode}-${code}`;
      elements.actionModuleInput.value = moduleCode;
      elements.actionCodeInput.value = code;
      elements.actionNameInput.value = name;
    },
    onDelete: () => {
      if (!window.confirm(`确认删除动作 ${moduleCode}${code} ${name} 吗？`)) {
        return;
      }
      state.metadata.actions = state.metadata.actions.filter((item) => !(item.moduleCode === moduleCode && item.code === code));
      state.editor.rows = state.editor.rows.filter((row) => !(row.moduleCode === moduleCode && row.actionCode === code));
      persist(`已删除动作 ${moduleCode}${code}。`);
      resetActionForm();
      renderAll();
    }
  }));
}

function renderList(container, items, mapItem) {
  container.innerHTML = "";
  if (!items.length) {
    container.innerHTML = '<div class="list-item"><span class="item-name">暂无数据</span></div>';
    return;
  }

  items.forEach((item) => {
    const meta = mapItem(item);
    const fragment = document.getElementById("listItemTemplate").content.cloneNode(true);
    fragment.querySelector(".item-code").textContent = meta.code;
    fragment.querySelector(".item-name").textContent = meta.name;
    fragment.querySelector(".text-btn.edit").addEventListener("click", meta.onEdit);
    fragment.querySelector(".text-btn.delete").addEventListener("click", meta.onDelete);
    container.appendChild(fragment);
  });
}

function renderActionModuleSelect() {
  const moduleOptions = state.metadata.modules.map((module) => ({ value: module.code, label: `${module.code} ${module.name}` }));
  elements.actionModuleInput.innerHTML = moduleOptions.map((option) => `<option value="${option.value}">${option.label}</option>`).join("");
  if (!elements.actionModuleInput.value && moduleOptions[0]) {
    elements.actionModuleInput.value = moduleOptions[0].value;
  }
}

function renderEditorHeader() {
  const modeOptions = state.metadata.modes.map((mode) => ({ value: mode.code, label: `${mode.code} ${mode.name}` }));
  elements.editorModeSelect.innerHTML = modeOptions.map((option) => `<option value="${option.value}">${option.label}</option>`).join("");
  if (!modeOptions.some((option) => option.value === state.editor.modeCode)) {
    state.editor.modeCode = modeOptions[0]?.value || "";
  }
  elements.editorModeSelect.value = state.editor.modeCode;
  elements.beatDurationInput.value = state.editor.beatDuration || "";
  const currentMode = state.metadata.modes.find((mode) => mode.code === state.editor.modeCode);
  if (currentMode && state.editor.beatDuration) {
    elements.timelineTitle.textContent = `${currentMode.code} ${currentMode.name} 节拍方案`;
    elements.timelineSubtitle.textContent = `当前节拍 ${state.editor.beatDuration} 秒，可直接编辑动作起始时间、前级动作与动作时间。`;
  } else {
    elements.timelineTitle.textContent = "未生成节拍方案";
    elements.timelineSubtitle.textContent = "选择模式并输入节拍时间后，将自动生成可编辑时序表。";
  }
}

function renderTimeline() {
  const beatDuration = state.editor.beatDuration;
  if (!state.editor.modeCode || !beatDuration || beatDuration < 1) {
    elements.timelineWrap.innerHTML = `
      <div class="timeline-empty">
        <strong>等待生成时序表</strong>
        <span>先选择模式并填写节拍时间，再点击 OK。</span>
      </div>
    `;
    return;
  }

  const rows = recalculateRows(state.editor.rows, beatDuration);
  state.editor.rows = rows;
  const timeMarks = Array.from({ length: beatDuration * 2 }, (_, index) => ((index + 1) * 0.5).toFixed(1).replace(".0", ""));
  const moduleColorMap = buildModuleColorMap();

  const legend = state.metadata.modules
    .map((module) => `<span class="summary-chip" style="--chip-color:${moduleColorMap[module.code] || "#999"}">${module.code} ${module.name}</span>`)
    .join("");

  const headerCells = timeMarks.map((mark) => `<th>${mark}</th>`).join("");
  const bodyRows = rows.map((row, index) => renderTimelineRow(row, index, moduleColorMap, beatDuration, rows.length)).join("");

  elements.timelineWrap.innerHTML = `
    <div style="padding: 16px 16px 0;">${legend}</div>
    <table class="timeline-table">
      <thead>
        <tr>
          <th class="left-sticky">模块</th>
          <th class="left-sticky-2">动作编号</th>
          <th class="left-sticky-3">动作名称</th>
          <th class="left-sticky-4">时间编辑</th>
          ${headerCells}
        </tr>
      </thead>
      <tbody>
        ${bodyRows}
      </tbody>
    </table>
  `;

  bindTimelineEvents();
}

function renderTimelineRow(row, index, moduleColorMap, beatDuration) {
  const action = getActionByCode(row.moduleCode, row.actionCode);
  const color = moduleColorMap[row.moduleCode] || "#999";
  const bar = buildTimelineBar(row, beatDuration, color, action?.name || "未命名动作");
  const moduleOptions = state.metadata.modules.map((module) => `
    <option value="${module.code}" ${module.code === row.moduleCode ? "selected" : ""}>
      ${module.code} ${module.name}
    </option>
  `).join("");
  const actionOptions = getActionsForModule(row.moduleCode).map((actionItem) => `
    <option value="${actionItem.code}" ${actionItem.code === row.actionCode ? "selected" : ""}>
      ${row.moduleCode}${actionItem.code}
    </option>
  `).join("");
  const previousOptions = [
    '<option value="">--</option>',
    ...state.editor.rows
      .filter((candidate) => candidate.id !== row.id)
      .map((candidate) => {
        const candidateAction = getActionByCode(candidate.moduleCode, candidate.actionCode);
        const text = `${candidate.moduleCode}${candidate.actionCode} ${candidateAction?.name || "未命名动作"}`;
        return `<option value="${candidate.id}" ${candidate.id === row.previousActionId ? "selected" : ""}>${text}</option>`;
      })
  ].join("");
  const activeClass = row.expanded ? "active" : "";

  return `
    <tr data-row-id="${row.id}">
      <td class="left-sticky">
        <select class="cell-select module-select" data-row-id="${row.id}">${moduleOptions}</select>
      </td>
      <td class="left-sticky-2">
        <select class="cell-select action-select" data-row-id="${row.id}">${actionOptions}</select>
      </td>
      <td class="left-sticky-3">
        <button class="action-name-button ${activeClass}" data-row-id="${row.id}">${action?.name || "请选择动作"}</button>
      </td>
      <td class="left-sticky-4">
        ${row.expanded ? `
          <div class="details-panel">
            <label>起始时间
              <input class="cell-input small start-input" data-row-id="${row.id}" type="number" min="0" max="${beatDuration}" step="0.1" value="${formatNumber(row.manualStartTime ?? row.startTime)}">
            </label>
            <label>前级动作
              <select class="cell-select prev-select" data-row-id="${row.id}">${previousOptions}</select>
            </label>
            <label>动作时间
              <input class="cell-input small duration-input" data-row-id="${row.id}" type="number" min="0.1" max="${beatDuration}" step="0.1" value="${formatNumber(row.duration || 1)}">
            </label>
            <button class="text-btn delete-row-btn" data-row-id="${row.id}">删除当前动作</button>
          </div>
        ` : `
          <div class="details-panel">
            <span>起始: ${formatNumber(row.startTime)}s</span>
            <span>时长: ${formatNumber(row.duration)}s</span>
            <span>${row.previousActionId ? "前级驱动" : "手动起始"}</span>
          </div>
        `}
      </td>
      ${bar}
    </tr>
  `;
}

function bindTimelineEvents() {
  document.querySelectorAll(".module-select").forEach((select) => {
    select.addEventListener("change", (event) => {
      const row = findRow(event.target.dataset.rowId);
      row.moduleCode = event.target.value;
      row.actionCode = getActionsForModule(row.moduleCode)[0]?.code || "00";
      persist("动作所属模块已更新。");
      renderTimeline();
    });
  });

  document.querySelectorAll(".action-select").forEach((select) => {
    select.addEventListener("change", (event) => {
      const row = findRow(event.target.dataset.rowId);
      row.actionCode = event.target.value;
      persist("动作编号已更新。");
      renderTimeline();
    });
  });

  document.querySelectorAll(".action-name-button").forEach((button) => {
    button.addEventListener("click", (event) => {
      const row = findRow(event.target.dataset.rowId);
      row.expanded = !row.expanded;
      renderTimeline();
    });
  });

  document.querySelectorAll(".start-input").forEach((input) => {
    input.addEventListener("change", (event) => {
      const row = findRow(event.target.dataset.rowId);
      row.manualStartTime = normalizeNumber(event.target.value, 0, state.editor.beatDuration, row.startTime, true);
      persist("动作起始时间已更新。");
      renderTimeline();
    });
  });

  document.querySelectorAll(".prev-select").forEach((select) => {
    select.addEventListener("change", (event) => {
      const row = findRow(event.target.dataset.rowId);
      row.previousActionId = event.target.value;
      persist("前级动作已更新。");
      renderTimeline();
    });
  });

  document.querySelectorAll(".duration-input").forEach((input) => {
    input.addEventListener("change", (event) => {
      const row = findRow(event.target.dataset.rowId);
      row.duration = normalizeNumber(event.target.value, 0.1, state.editor.beatDuration, 1, true);
      persist("动作时间已更新。");
      renderTimeline();
    });
  });

  document.querySelectorAll(".delete-row-btn").forEach((button) => {
    button.addEventListener("click", (event) => {
      const rowId = event.target.dataset.rowId;
      state.editor.rows = state.editor.rows
        .filter((row) => row.id !== rowId)
        .map((row) => ({
          ...row,
          previousActionId: row.previousActionId === rowId ? "" : row.previousActionId
        }));
      persist("动作行已删除。");
      renderTimeline();
    });
  });
}

function saveMode() {
  const code = elements.modeCodeInput.value;
  const name = elements.modeNameInput.value.trim();
  if (!code || !name) {
    window.alert("请填写完整的模式编号和名称。");
    return;
  }
  upsertByCode(state.metadata.modes, { code, name });
  activeEdit.modeCode = null;
  state.editor.modeCode ||= code;
  persist(`模式 ${code} 已保存。`);
  resetModeForm();
  renderAll();
}

function saveModule() {
  const code = elements.moduleCodeInput.value;
  const name = elements.moduleNameInput.value.trim();
  if (!code || !name) {
    window.alert("请填写完整的模块编号和名称。");
    return;
  }
  upsertByCode(state.metadata.modules, { code, name });
  activeEdit.moduleCode = null;
  persist(`模块 ${code} 已保存。`);
  resetModuleForm();
  renderAll();
}

function saveAction() {
  const moduleCode = elements.actionModuleInput.value;
  const code = elements.actionCodeInput.value;
  const name = elements.actionNameInput.value.trim();
  if (!moduleCode || !code || !name) {
    window.alert("请填写完整的动作信息。");
    return;
  }

  const existingIndex = state.metadata.actions.findIndex((item) => item.moduleCode === moduleCode && item.code === code);
  if (existingIndex >= 0) {
    state.metadata.actions[existingIndex] = { moduleCode, code, name };
  } else {
    state.metadata.actions.push({ moduleCode, code, name });
    state.metadata.actions.sort((a, b) => `${a.moduleCode}${a.code}`.localeCompare(`${b.moduleCode}${b.code}`));
  }

  activeEdit.actionKey = null;
  persist(`动作 ${moduleCode}${code} 已保存。`);
  resetActionForm();
  renderAll();
}

function resetModeForm() {
  activeEdit.modeCode = null;
  elements.modeCodeInput.value = LETTER_CODES.find((code) => !state.metadata.modes.some((mode) => mode.code === code)) || state.metadata.modes[0]?.code || "A";
  elements.modeNameInput.value = "";
}

function resetModuleForm() {
  activeEdit.moduleCode = null;
  elements.moduleCodeInput.value = LETTER_CODES.find((code) => !state.metadata.modules.some((module) => module.code === code)) || state.metadata.modules[0]?.code || "A";
  elements.moduleNameInput.value = "";
}

function resetActionForm() {
  activeEdit.actionKey = null;
  elements.actionModuleInput.value = state.metadata.modules[0]?.code || "";
  elements.actionCodeInput.value = "00";
  elements.actionNameInput.value = "";
}

function handleSoftSave() {
  persist("数据已保存在当前浏览器中。");
}

function handleSaveAs() {
  const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
  anchor.href = url;
  anchor.download = `instrument-rhythm-${stamp}.json`;
  anchor.click();
  URL.revokeObjectURL(url);
  elements.storageStatus.textContent = "已导出 JSON 文件。";
}

function handleFileImport(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const imported = JSON.parse(reader.result);
      if (!imported.metadata || !imported.editor) {
        throw new Error("invalid-shape");
      }
      imported.editor.rows = (imported.editor.rows || []).map((row) => ({
        expanded: false,
        ...row,
        id: row.id || crypto.randomUUID()
      }));
      Object.assign(state, imported);
      persist(`已从 ${file.name} 导入数据。`);
      renderAll();
    } catch (error) {
      window.alert("导入失败，文件格式不是有效的设计数据。");
    } finally {
      event.target.value = "";
    }
  };
  reader.readAsText(file, "utf-8");
}

function clearAllData() {
  if (!window.confirm("确认清空全部模式、模块、动作与当前节拍设计吗？")) {
    return;
  }
  const cleanState = {
    metadata: {
      modes: [],
      modules: [],
      actions: []
    },
    editor: {
      modeCode: "",
      beatDuration: 30,
      rows: []
    }
  };
  Object.assign(state, cleanState);
  persist("数据已清空，并恢复到空白编辑状态。");
  resetModeForm();
  resetModuleForm();
  resetActionForm();
  renderAll();
}

function generateTimeline() {
  state.editor.modeCode = elements.editorModeSelect.value;
  state.editor.beatDuration = normalizeNumber(elements.beatDurationInput.value, 1, 999, 30, false);
  if (!state.editor.rows.length) {
    addTimelineRow(false);
  }
  persist("时序表已生成。");
  renderEditorHeader();
  renderTimeline();
}

function addTimelineRow(shouldRender = true) {
  if (!state.metadata.modules.length) {
    window.alert("请先在一级动作列表中创建模块。");
    return;
  }
  const firstModule = state.metadata.modules[0]?.code || "A";
  const firstAction = getActionsForModule(firstModule)[0]?.code || "00";
  state.editor.rows.push({
    id: crypto.randomUUID(),
    moduleCode: firstModule,
    actionCode: firstAction,
    startTime: 0,
    manualStartTime: 0,
    previousActionId: "",
    duration: 1,
    expanded: true
  });
  persist("已新增动作行。");
  if (shouldRender) {
    renderTimeline();
  }
}

function buildTimelineBar(row, beatDuration, color, label) {
  const totalColumns = beatDuration * 2;
  const startColumn = Math.round((row.startTime || 0) * 2);
  const spanColumns = Math.max(1, Math.round((row.duration || 0.5) * 2));
  let cursor = 0;
  let html = "";
  while (cursor < totalColumns) {
    if (cursor === startColumn) {
      const colspan = Math.min(spanColumns, totalColumns - startColumn);
      html += `
        <td class="timeline-cell" colspan="${colspan}">
          <div class="bar-fragment" style="left: 0; right: 0; background:${color};">
            <span class="bar-label">${label}</span>
          </div>
        </td>
      `;
      cursor += colspan;
      continue;
    }
    html += '<td class="timeline-cell"></td>';
    cursor += 1;
  }
  return html;
}

function recalculateRows(rows, beatDuration) {
  const rowMap = new Map(rows.map((row) => [row.id, { ...row }]));
  const visiting = new Set();
  const visited = new Set();

  function compute(rowId) {
    if (visited.has(rowId)) {
      return rowMap.get(rowId);
    }
    if (visiting.has(rowId)) {
      const row = rowMap.get(rowId);
      row.previousActionId = "";
      row.startTime = normalizeNumber(row.manualStartTime ?? row.startTime ?? 0, 0, beatDuration, 0, true);
      return row;
    }
    visiting.add(rowId);
    const row = rowMap.get(rowId);
    row.duration = normalizeNumber(row.duration, 0.1, beatDuration, 1, true);
    const manualStart = normalizeNumber(row.manualStartTime ?? row.startTime ?? 0, 0, beatDuration, 0, true);
    row.manualStartTime = manualStart;
    if (row.previousActionId && rowMap.has(row.previousActionId)) {
      const previous = compute(row.previousActionId);
      row.startTime = clamp(previous.startTime + previous.duration, 0, beatDuration);
    } else {
      row.previousActionId = "";
      row.startTime = manualStart;
    }
    row.startTime = clamp(row.startTime, 0, Math.max(0, beatDuration - row.duration));
    visiting.delete(rowId);
    visited.add(rowId);
    return row;
  }

  rows.forEach((row) => compute(row.id));
  return rows.map((row) => ({ ...rowMap.get(row.id) }));
}

function buildModuleColorMap() {
  return Object.fromEntries(
    state.metadata.modules.map((module, index) => [module.code, MODULE_COLORS[index % MODULE_COLORS.length]])
  );
}

function findRow(rowId) {
  return state.editor.rows.find((row) => row.id === rowId);
}

function getActionsForModule(moduleCode) {
  return state.metadata.actions.filter((action) => action.moduleCode === moduleCode);
}

function getActionByCode(moduleCode, actionCode) {
  return state.metadata.actions.find((action) => action.moduleCode === moduleCode && action.code === actionCode);
}

function switchTab(targetId) {
  document.querySelectorAll(".tab-button").forEach((button) => {
    button.classList.toggle("active", button.dataset.tabTarget === targetId);
  });
  document.querySelectorAll(".tab-panel").forEach((panel) => {
    panel.classList.toggle("active", panel.id === targetId);
  });
}

function persist(message) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  elements.storageStatus.textContent = message;
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return structuredClone(initialState);
    }
    const parsed = JSON.parse(raw);
    parsed.editor.rows = (parsed.editor.rows || []).map((row) => ({
      expanded: Boolean(row.expanded),
      ...row,
      id: row.id || crypto.randomUUID()
    }));
    return {
      metadata: {
        modes: parsed.metadata?.modes || structuredClone(initialState.metadata.modes),
        modules: parsed.metadata?.modules || structuredClone(initialState.metadata.modules),
        actions: parsed.metadata?.actions || structuredClone(initialState.metadata.actions)
      },
      editor: {
        modeCode: parsed.editor?.modeCode || initialState.editor.modeCode,
        beatDuration: parsed.editor?.beatDuration || initialState.editor.beatDuration,
        rows: parsed.editor?.rows?.length ? parsed.editor.rows : structuredClone(initialState.editor.rows)
      }
    };
  } catch {
    return structuredClone(initialState);
  }
}

function populateSelect(select, options, preferredValue) {
  select.innerHTML = options.map((value) => `<option value="${value}">${value}</option>`).join("");
  if (preferredValue && options.includes(preferredValue)) {
    select.value = preferredValue;
  } else if (options.length) {
    select.value = options[0];
  }
}

function upsertByCode(collection, item) {
  const index = collection.findIndex((entry) => entry.code === item.code);
  if (index >= 0) {
    collection[index] = item;
  } else {
    collection.push(item);
    collection.sort((a, b) => a.code.localeCompare(b.code));
  }
}

function normalizeNumber(value, min, max, fallback, allowDecimal) {
  const numericValue = allowDecimal ? Number.parseFloat(value) : Number.parseInt(value, 10);
  if (Number.isNaN(numericValue)) {
    return fallback;
  }
  return clamp(numericValue, min, max);
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(value, max));
}

function formatNumber(value) {
  return Number(value).toFixed(1).replace(".0", "");
}

resetModeForm();
resetModuleForm();
resetActionForm();
