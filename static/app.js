const ROW_HEIGHT = 50;
const HEADER_HEIGHT = 84;
const MONTH_HEADER_Y = 10;
const MONTH_HEADER_HEIGHT = 16;
const DAY_HEADER_Y = 32;
const DAY_HEADER_HEIGHT = 34;

const ONBOARDING_STORAGE_KEY = "task-gantt:onboarding:v1";

const ONBOARDING_STEPS = [
  {
    target: ".sidebar-panel",
    title: "先从左侧选项目",
    copy: "这里是你能访问的项目列表。点击一个项目后，右侧会同步显示项目概览、甘特图和任务清单。",
    tip: "如果项目很多，历史项目会折叠在底部，不会打断当前工作。",
    placement: "right",
  },
  {
    target: ".top-actions",
    title: "顶部是主要工作入口",
    copy: "新建项目、导入文件或文本、会议更新、历史回退、导出，都集中在这里。日常使用通常从这里开始。",
    tip: "第一次创建项目可以点“新建项目”，粘贴路线图后让模型自动拆任务。",
    placement: "bottom",
  },
  {
    target: ".summary-panel",
    title: "先看项目状态，再决定动作",
    copy: "项目概览会显示当前健康度、进度、时间范围和关键统计。进入页面后先看这里，能快速判断项目是否推进正常。",
    tip: "如果发现延期或阻塞，再去任务清单里调整具体任务。",
    placement: "bottom",
  },
  {
    target: ".gantt-panel",
    title: "甘特图负责看排期关系",
    copy: "这里用时间轴展示任务跨度和依赖。你可以横向滚动看更远日期，点击甘特图中的任务条选择任务。",
    tip: "修改任务日期或依赖后，点“重算排期”可以重新整理时间线。",
    placement: "top",
  },
  {
    target: ".table-panel",
    title: "任务清单负责精确更新",
    copy: "要更新进展时，先在任务清单点选一行，再到下方“编辑当前任务”里改状态、进度、负责人、工时或备注。",
    tip: "单个任务进度建议在这里手动改，改完点“保存”或“保存并重算”。",
    placement: "top",
  },
  {
    target: "#open-meeting-dialog",
    title: "周会后用会议更新批量推进",
    copy: "如果要根据会议纪要更新项目进展，点击“会议更新”，粘贴周会记录，系统会对照当前任务自动调整进度、状态和备注。",
    tip: "如果是新增一批任务，用“导入文件/文本”；如果只是更新已有任务，用“会议更新”。",
    placement: "bottom",
  },
];
const PROJECT_ANALYSIS_STAGES = [
  { key: "submit", threshold: 0, label: "\u6b63\u5728\u63d0\u4ea4\u9879\u76ee\u63cf\u8ff0\u2026" },
  { key: "llm", threshold: 22, label: "\u6b63\u5728\u8c03\u7528 DeepSeek \u5206\u6790\u9879\u76ee\u8303\u56f4\u2026" },
  { key: "shape", threshold: 48, label: "\u6b63\u5728\u6574\u7406\u4efb\u52a1\u3001\u4f9d\u8d56\u548c\u4f30\u65f6\u2026" },
  { key: "persist", threshold: 74, label: "\u6b63\u5728\u5199\u5165\u672c\u5730\u9879\u76ee\u5e76\u8ba1\u7b97\u6392\u671f\u2026" },
];

const state = {
  projects: [],
  deletedProjects: [],
  historyExpanded: false,
  snapshots: [],
  detail: null,
  selectedProjectId: null,
  selectedTaskId: null,
  toastTimer: null,
  llmProgressTimer: null,
  llmProgressValue: 0,
  llmProgressActive: false,
  importBusy: false,
  snapshotBusy: false,
  meetingBusy: false,
  importProgressTimer: null,
  importProgressValue: 0,
  importProgressActive: false,
  onboardingActive: false,
  onboardingIndex: 0,
};

const els = {};

window.addEventListener("DOMContentLoaded", init);
window.addEventListener("resize", debounce(renderGantt, 80));

async function init() {
  cacheElements();
  setDefaultDates();
  resetLlmProgress();
  bindEvents();
  await loadBootstrap();
  scheduleFirstRunOnboarding();
}

function cacheElements() {
  els.projectDialog = document.getElementById("project-dialog");
  els.importDialog = document.getElementById("import-dialog");
  els.meetingDialog = document.getElementById("meeting-dialog");
  els.snapshotDialog = document.getElementById("snapshot-dialog");
  els.projectForm = document.getElementById("project-form");
  els.importForm = document.getElementById("import-form");
  els.meetingForm = document.getElementById("meeting-form");
  els.snapshotForm = document.getElementById("snapshot-form");
  els.projectList = document.getElementById("project-list");
  els.deletedProjectWrap = document.getElementById("deleted-project-wrap");
  els.deletedProjectList = document.getElementById("deleted-project-list");
  els.historyToggleBtn = document.getElementById("history-toggle-btn");
  els.historyToggleCount = document.getElementById("history-toggle-count");
  els.projectSummary = document.getElementById("project-summary");
  els.ganttMeta = document.getElementById("gantt-meta");
  els.tableMeta = document.getElementById("table-meta");
  els.taskTableBody = document.getElementById("task-table-body");
  els.taskForm = document.getElementById("task-form");
  els.taskEditorDetails = document.getElementById("task-editor-details");
  els.taskEditorSummary = els.taskEditorDetails.querySelector("summary");
  els.editorState = document.getElementById("editor-state");
  els.ganttViewport = document.getElementById("gantt-viewport");
  els.ganttSvg = document.getElementById("gantt-svg");
  els.toast = document.getElementById("toast");
  els.openProjectDialog = document.getElementById("open-project-dialog");
  els.openImportDialog = document.getElementById("open-import-dialog");
  els.openMeetingDialog = document.getElementById("open-meeting-dialog");
  els.openSnapshotDialog = document.getElementById("open-snapshot-dialog");
  els.addTaskBtn = document.getElementById("add-task-btn");
  els.rescheduleBtn = document.getElementById("reschedule-btn");
  els.saveRescheduleBtn = document.getElementById("save-reschedule-btn");
  els.deleteTaskBtn = document.getElementById("delete-task-btn");
  els.exportAllXlsxBtn = document.getElementById("export-all-xlsx-btn");
  els.projectModeButtons = Array.from(document.querySelectorAll("[data-project-mode]"));
  els.smartImportBtn = document.getElementById("smart-import-btn");
  els.projectDialogCloseButton = els.projectDialog.querySelector('[data-close-dialog="project-dialog"]');
  els.llmProgressPanel = document.getElementById("llm-progress-panel");
  els.llmProgressTitle = document.getElementById("llm-progress-title");
  els.llmProgressPercent = document.getElementById("llm-progress-percent");
  els.llmProgressFill = document.getElementById("llm-progress-fill");
  els.llmProgressText = document.getElementById("llm-progress-text");
  els.llmProgressBar = els.llmProgressPanel.querySelector(".llm-progress-bar");
  els.llmProgressSteps = Array.from(els.llmProgressPanel.querySelectorAll("[data-progress-step]"));
  els.importProgressPanel = document.getElementById("import-progress-panel");
  els.importProgressTitle = document.getElementById("import-progress-title");
  els.importProgressPercent = document.getElementById("import-progress-percent");
  els.importProgressFill = document.getElementById("import-progress-fill");
  els.importProgressText = document.getElementById("import-progress-text");
  els.importProgressBar = document.getElementById("import-progress-bar");
  els.importProgressSteps = Array.from(document.querySelectorAll("[data-import-progress-step]"));
  els.snapshotList = document.getElementById("snapshot-list");
  els.openOnboardingBtn = document.getElementById("open-onboarding-btn");
  els.onboardingLayer = document.getElementById("onboarding-layer");
  els.onboardingCard = document.getElementById("onboarding-card");
  els.onboardingHighlight = document.getElementById("onboarding-highlight");
  els.onboardingCount = document.getElementById("onboarding-count");
  els.onboardingTitle = document.getElementById("onboarding-title");
  els.onboardingCopy = document.getElementById("onboarding-copy");
  els.onboardingTip = document.getElementById("onboarding-tip");
  els.onboardingFill = document.getElementById("onboarding-progress-fill");
  els.onboardingClose = document.getElementById("onboarding-close");
  els.onboardingSkip = document.getElementById("onboarding-skip");
  els.onboardingPrev = document.getElementById("onboarding-prev");
  els.onboardingNext = document.getElementById("onboarding-next");
}

function setDefaultDates() {
  const today = formatDateInput(new Date());
  const future = formatDateInput(addDays(new Date(), 21));
  els.projectForm.elements.start_date.value = today;
  els.projectForm.elements.due_date.value = future;
  els.importForm.elements.start_date.value = today;
}

function bindEvents() {
  els.openProjectDialog.addEventListener("click", () => openDialog(els.projectDialog));
  els.openImportDialog.addEventListener("click", () => openDialog(els.importDialog));
  els.openMeetingDialog.addEventListener("click", () => openDialog(els.meetingDialog));
  els.openSnapshotDialog.addEventListener("click", handleOpenSnapshotDialog);
  if (els.openOnboardingBtn) {
    els.openOnboardingBtn.addEventListener("click", () => startOnboarding({ force: true }));
  }

  document.querySelectorAll("[data-close-dialog]").forEach((button) => {
    button.addEventListener("click", () => closeDialog(document.getElementById(button.dataset.closeDialog)));
  });


  if (els.onboardingClose) {
    els.onboardingClose.addEventListener("click", completeOnboarding);
  }
  if (els.onboardingSkip) {
    els.onboardingSkip.addEventListener("click", completeOnboarding);
  }
  if (els.onboardingPrev) {
    els.onboardingPrev.addEventListener("click", () => moveOnboarding(-1));
  }
  if (els.onboardingNext) {
    els.onboardingNext.addEventListener("click", () => moveOnboarding(1));
  }
  document.addEventListener("keydown", handleOnboardingKeydown);
  window.addEventListener("resize", debounce(positionOnboardingCard, 80));
  window.addEventListener("scroll", debounce(positionOnboardingCard, 80), true);
  els.projectModeButtons.forEach((button) => {
    button.addEventListener("click", () => handleProjectCreate(button.dataset.projectMode));
  });

  els.projectDialog.addEventListener("cancel", (event) => {
    if (state.llmProgressActive) {
      event.preventDefault();
    }
  });
  els.projectDialog.addEventListener("close", () => {
    if (!state.llmProgressActive) {
      resetLlmProgress();
    }
  });

  els.importDialog.addEventListener("cancel", (event) => {
    if (state.importBusy) {
      event.preventDefault();
    }
  });
  els.importDialog.addEventListener("close", () => {
    if (!state.importProgressActive) {
      resetImportProgress();
    }
  });

    els.meetingDialog.addEventListener("cancel", (event) => {
    if (state.meetingBusy) {
      event.preventDefault();
    }
  });
  els.snapshotDialog.addEventListener("cancel", (event) => {
    if (state.snapshotBusy) {
      event.preventDefault();
    }
  });
  els.importForm.addEventListener("submit", handleImportSubmit);
  els.meetingForm.addEventListener("submit", handleMeetingUpdateSubmit);
  els.smartImportBtn.addEventListener("click", handleSmartImportSubmit);
  els.addTaskBtn.addEventListener("click", handleAddTask);
  els.rescheduleBtn.addEventListener("click", handleReschedule);
  els.taskForm.addEventListener("submit", (event) => {
    event.preventDefault();
    saveTask(false);
  });
  els.saveRescheduleBtn.addEventListener("click", () => saveTask(true));
  els.deleteTaskBtn.addEventListener("click", handleDeleteTask);

  document.querySelectorAll("[data-export]").forEach((button) => {
    button.addEventListener("click", () => exportCurrentProject(button.dataset.export));
  });
  if (els.exportAllXlsxBtn) {
    els.exportAllXlsxBtn.addEventListener("click", exportAllProjectsXlsx);
  }

  els.projectList.addEventListener("click", (event) => {
    const deleteButton = event.target.closest("[data-delete-project-id]");
    if (deleteButton) {
      event.preventDefault();
      event.stopPropagation();
      handleDeleteProject(Number(deleteButton.dataset.deleteProjectId));
      return;
    }

    const item = event.target.closest("[data-project-id]");
    if (!item) {
      return;
    }
    loadProject(Number(item.dataset.projectId));
  });

  if (els.historyToggleBtn) {
    els.historyToggleBtn.addEventListener("click", () => {
      if (!state.deletedProjects.length) {
        return;
      }
      state.historyExpanded = !state.historyExpanded;
      renderDeletedProjectList();
    });
  }

  if (els.deletedProjectList) {
    els.deletedProjectList.addEventListener("click", (event) => {
      const restoreButton = event.target.closest("[data-restore-project-id]");
      if (!restoreButton) {
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      handleRestoreProject(Number(restoreButton.dataset.restoreProjectId));
    });
  }

  els.taskTableBody.addEventListener("click", (event) => {
    const row = event.target.closest("[data-task-id]");
    if (!row) {
      return;
    }
    selectTask(Number(row.dataset.taskId));
  });
}


function hasSeenOnboarding() {
  try {
    return window.localStorage.getItem(ONBOARDING_STORAGE_KEY) === "done";
  } catch (error) {
    return true;
  }
}

function markOnboardingSeen() {
  try {
    window.localStorage.setItem(ONBOARDING_STORAGE_KEY, "done");
  } catch (error) {
    // LocalStorage can be blocked in strict browser modes. The guide still works for the current session.
  }
}

function scheduleFirstRunOnboarding() {
  if (hasSeenOnboarding()) {
    return;
  }
  window.setTimeout(() => startOnboarding(), 520);
}

function startOnboarding(options = {}) {
  if (!els.onboardingLayer) {
    return;
  }
  if (!options.force && hasSeenOnboarding()) {
    return;
  }
  closeOpenDialogsForOnboarding();
  state.onboardingActive = true;
  state.onboardingIndex = 0;
  els.onboardingLayer.hidden = false;
  els.onboardingLayer.setAttribute("aria-hidden", "false");
  document.body.classList.add("onboarding-active");
  renderOnboardingStep();
}

function closeOpenDialogsForOnboarding() {
  [els.projectDialog, els.importDialog, els.meetingDialog, els.snapshotDialog].forEach((dialog) => {
    if (!dialog || !dialog.open) {
      return;
    }
    if (typeof dialog.close === "function") {
      dialog.close();
    } else {
      dialog.removeAttribute("open");
    }
  });
}

function getOnboardingTarget(step) {
  const target = document.querySelector(step.target);
  return target || document.querySelector(".main-shell") || document.body;
}

function clearOnboardingTarget() {
  document.querySelectorAll(".onboarding-target").forEach((node) => {
    node.classList.remove("onboarding-target");
  });
}

function renderOnboardingStep() {
  if (!state.onboardingActive || !els.onboardingLayer) {
    return;
  }
  const step = ONBOARDING_STEPS[state.onboardingIndex];
  const target = getOnboardingTarget(step);
  clearOnboardingTarget();
  target.classList.add("onboarding-target");
  target.scrollIntoView({ block: "center", inline: "nearest", behavior: "smooth" });

  els.onboardingCount.textContent = `${state.onboardingIndex + 1} / ${ONBOARDING_STEPS.length}`;
  els.onboardingTitle.textContent = step.title;
  els.onboardingCopy.textContent = step.copy;
  els.onboardingTip.textContent = step.tip;
  els.onboardingPrev.disabled = state.onboardingIndex === 0;
  els.onboardingNext.textContent = state.onboardingIndex === ONBOARDING_STEPS.length - 1 ? "完成" : "下一步";
  els.onboardingFill.style.width = `${((state.onboardingIndex + 1) / ONBOARDING_STEPS.length) * 100}%`;

  window.setTimeout(positionOnboardingCard, 220);
}

function positionOnboardingCard() {
  if (!state.onboardingActive || !els.onboardingCard) {
    return;
  }
  const step = ONBOARDING_STEPS[state.onboardingIndex];
  const target = getOnboardingTarget(step);
  const rect = target.getBoundingClientRect();
  if (els.onboardingHighlight) {
    const pad = 8;
    els.onboardingHighlight.style.left = `${Math.max(8, rect.left - pad)}px`;
    els.onboardingHighlight.style.top = `${Math.max(8, rect.top - pad)}px`;
    els.onboardingHighlight.style.width = `${Math.min(window.innerWidth - 16, rect.width + pad * 2)}px`;
    els.onboardingHighlight.style.height = `${Math.min(window.innerHeight - 16, rect.height + pad * 2)}px`;
  }
  const card = els.onboardingCard;
  const margin = 18;
  const cardWidth = Math.min(390, window.innerWidth - margin * 2);
  const cardHeight = card.offsetHeight || 280;
  let left = margin;
  let top = margin;

  if (window.innerWidth <= 760) {
    left = margin;
    top = Math.min(window.innerHeight - cardHeight - margin, Math.max(margin, rect.bottom + 12));
  } else if (step.placement === "right") {
    left = rect.right + margin;
    top = rect.top;
  } else if (step.placement === "top") {
    left = rect.left;
    top = rect.top - cardHeight - margin;
  } else {
    left = rect.left;
    top = rect.bottom + margin;
  }

  left = Math.min(Math.max(margin, left), window.innerWidth - cardWidth - margin);
  top = Math.min(Math.max(margin, top), window.innerHeight - cardHeight - margin);
  card.style.width = `${cardWidth}px`;
  card.style.left = `${left}px`;
  card.style.top = `${top}px`;
}

function moveOnboarding(direction) {
  if (!state.onboardingActive) {
    return;
  }
  const nextIndex = state.onboardingIndex + direction;
  if (nextIndex >= ONBOARDING_STEPS.length) {
    completeOnboarding();
    return;
  }
  state.onboardingIndex = clamp(nextIndex, 0, ONBOARDING_STEPS.length - 1);
  renderOnboardingStep();
}

function completeOnboarding() {
  if (!state.onboardingActive) {
    markOnboardingSeen();
    return;
  }
  markOnboardingSeen();
  state.onboardingActive = false;
  clearOnboardingTarget();
  document.body.classList.remove("onboarding-active");
  if (els.onboardingLayer) {
    els.onboardingLayer.hidden = true;
    els.onboardingLayer.setAttribute("aria-hidden", "true");
  }
}

function handleOnboardingKeydown(event) {
  if (!state.onboardingActive) {
    return;
  }
  if (event.key === "Escape") {
    completeOnboarding();
  }
  if (event.key === "ArrowRight") {
    moveOnboarding(1);
  }
  if (event.key === "ArrowLeft") {
    moveOnboarding(-1);
  }
}
function openDialog(dialog) {
  if (!dialog) {
    return;
  }
  if (dialog === els.projectDialog) {
    resetLlmProgress();
  }
  if (dialog === els.importDialog) {
    resetImportProgress();
  }
  if (typeof dialog.showModal === "function") {
    dialog.showModal();
  } else {
    dialog.setAttribute("open", "open");
  }
}

function closeDialog(dialog) {
  if (!dialog) {
    return;
  }
  if (dialog === els.projectDialog && state.llmProgressActive) {
    return;
  }
  if (dialog === els.importDialog && state.importProgressActive) {
    return;
  }
  if (dialog === els.meetingDialog && state.meetingBusy) {
    return;
  }
  if (typeof dialog.close === "function") {
    dialog.close();
  } else {
    dialog.removeAttribute("open");
  }
}

function wait(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function getLlmProgressStage(value) {
  let stage = PROJECT_ANALYSIS_STAGES[0];
  for (const candidate of PROJECT_ANALYSIS_STAGES) {
    if (value >= candidate.threshold) {
      stage = candidate;
    }
  }
  return stage;
}

function setProjectCreateBusy(isBusy) {
  state.llmProgressActive = isBusy;
  els.projectForm.classList.toggle("is-processing", isBusy);


  els.projectModeButtons.forEach((button) => {
    button.disabled = isBusy;
  });
  if (els.projectDialogCloseButton) {
    els.projectDialogCloseButton.disabled = isBusy;
  }
}

function setImportBusy(isBusy) {
  state.importBusy = isBusy;
  els.importForm.classList.toggle("is-processing", isBusy);
  els.importForm.querySelectorAll("input, textarea, button").forEach((control) => {
    control.disabled = isBusy;
  });
}


function setMeetingBusy(isBusy) {
  state.meetingBusy = isBusy;
  els.meetingForm.classList.toggle("is-processing", isBusy);
  els.meetingForm.querySelectorAll("input, textarea, button").forEach((control) => {
    control.disabled = isBusy;
  });
}
function setSnapshotBusy(isBusy) {
  state.snapshotBusy = isBusy;
  els.snapshotForm.classList.toggle("is-processing", isBusy);
  const controls = els.snapshotForm.querySelectorAll("button");
  controls.forEach((control) => {
    control.disabled = isBusy;
  });
}
function stopImportProgressTimer() {
  if (state.importProgressTimer) {
    clearInterval(state.importProgressTimer);
    state.importProgressTimer = null;
  }
}

function renderImportProgress(textOverride = "") {
  const value = Math.round(state.importProgressValue);
  const stage = getLlmProgressStage(value);
  els.importProgressPanel.hidden = false;
  els.importProgressFill.style.width = `${value}%`;
  els.importProgressPercent.textContent = `${value}%`;
  els.importProgressText.textContent = textOverride || stage.label;
  els.importProgressBar.setAttribute("aria-valuenow", String(value));

  const currentIndex = PROJECT_ANALYSIS_STAGES.findIndex((item) => item.key === stage.key);
  els.importProgressSteps.forEach((step, index) => {
    step.classList.toggle("done", index < currentIndex || value >= 100);
    step.classList.toggle("active", index === currentIndex && value < 100);
  });
}

function startImportProgress() {
  stopImportProgressTimer();
  state.importProgressActive = true;
  state.importProgressValue = 8;
  els.importProgressPanel.classList.remove("is-error", "is-complete");
  els.importProgressTitle.textContent = "\u6b63\u5728\u667a\u80fd\u8ffd\u52a0\u4efb\u52a1";
  renderImportProgress("\u5df2\u63d0\u4ea4\u6587\u672c\uff0c\u51c6\u5907\u8c03\u7528\u6a21\u578b\u2026");

  state.importProgressTimer = window.setInterval(() => {
    let next = state.importProgressValue;
    if (next < 22) {
      next += 4;
    } else if (next < 48) {
      next += 2.8;
    } else if (next < 74) {
      next += 1.6;
    } else if (next < 88) {
      next += 0.9;
    }
    state.importProgressValue = Math.min(88, next);
    renderImportProgress();
    if (state.importProgressValue >= 88) {
      stopImportProgressTimer();
    }
  }, 280);
}

async function finishImportProgress(message) {
  stopImportProgressTimer();
  state.importProgressValue = 100;
  els.importProgressPanel.classList.remove("is-error");
  els.importProgressPanel.classList.add("is-complete");
  els.importProgressTitle.textContent = "\u4efb\u52a1\u8ffd\u52a0\u5b8c\u6210";
  renderImportProgress(message || "\u65b0\u589e\u4efb\u52a1\u5df2\u7ecf\u5199\u5165\u5f53\u524d\u9879\u76ee\uff0c\u6b63\u5728\u5237\u65b0\u7518\u7279\u56fe\u2026");
  await wait(420);
  state.importProgressActive = false;
}

function failImportProgress(message) {
  stopImportProgressTimer();
  els.importProgressPanel.classList.remove("is-complete");
  els.importProgressPanel.classList.add("is-error");
  els.importProgressTitle.textContent = "\u4efb\u52a1\u8ffd\u52a0\u5931\u8d25";
  renderImportProgress(message || "\u8c03\u7528\u5931\u8d25\uff0c\u8bf7\u7a0d\u540e\u91cd\u8bd5");
  state.importProgressActive = false;
}

function resetImportProgress() {
  stopImportProgressTimer();
  state.importProgressValue = 0;
  state.importProgressActive = false;
  els.importProgressPanel.hidden = true;
  els.importProgressPanel.classList.remove("is-error", "is-complete");
  els.importProgressTitle.textContent = "\u6b63\u5728\u51c6\u5907\u89e3\u6790\u6587\u672c";
  els.importProgressPercent.textContent = "0%";
  els.importProgressFill.style.width = "0%";
  els.importProgressText.textContent = "\u51c6\u5907\u63d0\u4ea4\u6587\u672c\u89e3\u6790\u2026";
  els.importProgressBar.setAttribute("aria-valuenow", "0");
  els.importProgressSteps.forEach((step) => {
    step.classList.remove("done", "active");
  });
}

function stopLlmProgressTimer() {
  if (state.llmProgressTimer) {
    clearInterval(state.llmProgressTimer);
    state.llmProgressTimer = null;
  }
}

function renderLlmProgress(textOverride = "") {
  const value = Math.round(state.llmProgressValue);
  const stage = getLlmProgressStage(value);
  els.llmProgressPanel.hidden = false;
  els.llmProgressFill.style.width = `${value}%`;
  els.llmProgressPercent.textContent = `${value}%`;
  els.llmProgressText.textContent = textOverride || stage.label;
  els.llmProgressBar.setAttribute("aria-valuenow", String(value));

  const currentIndex = PROJECT_ANALYSIS_STAGES.findIndex((item) => item.key === stage.key);
  els.llmProgressSteps.forEach((step, index) => {
    step.classList.toggle("done", index < currentIndex || value >= 100);
    step.classList.toggle("active", index === currentIndex && value < 100);
  });
}

function startLlmProgress() {
  stopLlmProgressTimer();
  state.llmProgressValue = 8;
  els.llmProgressPanel.classList.remove("is-error", "is-complete");
  els.llmProgressTitle.textContent = "\u6b63\u5728\u7528 LLM \u5206\u6790\u9879\u76ee";
  setProjectCreateBusy(true);
  renderLlmProgress("\u5df2\u63d0\u4ea4\u9879\u76ee\u63cf\u8ff0\uff0c\u51c6\u5907\u8c03\u7528\u6a21\u578b\u2026");

  state.llmProgressTimer = window.setInterval(() => {
    let next = state.llmProgressValue;
    if (next < 22) {
      next += 4;
    } else if (next < 48) {
      next += 2.8;
    } else if (next < 74) {
      next += 1.6;
    } else if (next < 88) {
      next += 0.9;
    }
    state.llmProgressValue = Math.min(88, next);
    renderLlmProgress();
    if (state.llmProgressValue >= 88) {
      stopLlmProgressTimer();
    }
  }, 280);
}

async function finishLlmProgress(message) {
  stopLlmProgressTimer();
  state.llmProgressValue = 100;
  els.llmProgressPanel.classList.remove("is-error");
  els.llmProgressPanel.classList.add("is-complete");
  els.llmProgressTitle.textContent = "\u9879\u76ee\u5206\u6790\u5b8c\u6210";
  renderLlmProgress(message || "\u4efb\u52a1\u5df2\u7ecf\u751f\u6210\u5b8c\u6210\uff0c\u6b63\u5728\u6253\u5f00\u9879\u76ee\u2026");
  await wait(420);
  setProjectCreateBusy(false);
}

function failLlmProgress(message) {
  stopLlmProgressTimer();
  els.llmProgressPanel.classList.remove("is-complete");
  els.llmProgressPanel.classList.add("is-error");
  els.llmProgressTitle.textContent = "\u9879\u76ee\u5206\u6790\u5931\u8d25";
  renderLlmProgress(message || "\u8c03\u7528\u5931\u8d25\uff0c\u8bf7\u68c0\u67e5 Key \u6216\u7a0d\u540e\u91cd\u8bd5");
  setProjectCreateBusy(false);
}

function resetLlmProgress() {
  stopLlmProgressTimer();
  state.llmProgressValue = 0;
  setProjectCreateBusy(false);
  els.llmProgressPanel.hidden = true;
  els.llmProgressPanel.classList.remove("is-error", "is-complete");
  els.llmProgressTitle.textContent = "\u6b63\u5728\u51c6\u5907\u9879\u76ee\u5206\u6790";
  els.llmProgressPercent.textContent = "0%";
  els.llmProgressFill.style.width = "0%";
  els.llmProgressText.textContent = "\u51c6\u5907\u63d0\u4ea4\u9879\u76ee\u63cf\u8ff0\u2026";
  els.llmProgressBar.setAttribute("aria-valuenow", "0");
  els.llmProgressSteps.forEach((step) => {
    step.classList.remove("done", "active");
  });
}

async function api(path, options = {}) {
  const response = await fetch(path, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  if (!response.ok) {
    let message = `请求失败: ${response.status}`;
    try {
      const payload = await response.json();
      if (payload.error) {
        message = payload.error;
      }
    } catch (error) {
      // ignore
    }
    throw new Error(message);
  }
  const contentType = response.headers.get("Content-Type") || "";
  return contentType.includes("application/json") ? response.json() : response;
}

async function loadBootstrap() {
  try {
    const data = await api("/api/bootstrap");
    state.projects = data.projects || [];
    state.deletedProjects = data.deleted_projects || [];
    if (data.detail) {
      acceptDetail(data.detail);
    } else {
      state.detail = null;
      state.selectedProjectId = data.selected_project_id ?? null;
      state.selectedTaskId = null;
      els.taskEditorDetails.open = false;
    }
    renderAll();
  } catch (error) {
    showToast(error.message, true);
  }
}

async function refreshProjects() {
  const data = await api("/api/projects");
  state.projects = data.projects || [];
  state.deletedProjects = data.deleted_projects || [];
}

async function loadProject(projectId) {
  try {
    const detail = await api(`/api/projects/${projectId}`);
    acceptDetail(detail);
    await refreshProjects();
    renderAll();
  } catch (error) {
    showToast(error.message, true);
  }
}

async function handleDeleteProject(projectId) {
  const project = state.projects.find((item) => item.id === projectId);
  if (!project) {
    return;
  }

  if (!window.confirm(`确认删除项目“${project.name}”吗？删除后可在历史项目中恢复。`)) {
    return;
  }

  try {
    await api(`/api/projects/${projectId}`, {
      method: "DELETE",
      body: JSON.stringify({}),
    });
    await refreshProjects();

    if (!state.projects.length) {
      state.detail = null;
      state.selectedProjectId = null;
      state.selectedTaskId = null;
      els.taskEditorDetails.open = false;
      renderAll();
      showToast("项目已删除，可在历史项目中恢复");
      return;
    }

    const nextProjectId = state.selectedProjectId === projectId ? state.projects[0].id : state.selectedProjectId;
    if (nextProjectId) {
      await loadProject(nextProjectId);
    } else {
      renderAll();
    }
    showToast("项目已删除，可在历史项目中恢复");
  } catch (error) {
    showToast(error.message, true);
  }
}

async function handleRestoreProject(projectId) {
  try {
    const detail = await api(`/api/projects/${projectId}/restore`, {
      method: "POST",
      body: JSON.stringify({}),
    });
    acceptDetail(detail);
    await refreshProjects();
    renderAll();
    showToast("项目已恢复");
  } catch (error) {
    showToast(error.message, true);
  }
}

function acceptDetail(detail) {
  state.detail = detail;
  state.selectedProjectId = detail?.project?.id ?? null;
  const taskIds = new Set((detail?.tasks || []).map((task) => task.id));
  if (!taskIds.has(state.selectedTaskId)) {
    state.selectedTaskId = detail?.tasks?.[0]?.id ?? null;
  }
}

function getCurrentProject() {
  return state.detail?.project || null;
}

function getCurrentTasks() {
  return state.detail?.tasks || [];
}

function getSelectedTask() {
  return getCurrentTasks().find((task) => task.id === state.selectedTaskId) || null;
}

function selectTask(taskId) {
  state.selectedTaskId = taskId;
  els.taskEditorDetails.open = true;
  renderTaskTable();
  renderTaskEditor();
  renderGanttMeta();
  renderTableMeta();
  renderGantt();
}

function renderAll() {
  renderProjectList();
  renderDeletedProjectList();
  renderProjectSummary();
  renderGanttMeta();
  renderTaskTable();
  renderTableMeta();
  renderTaskEditor();
  renderGantt();
}

function renderProjectList() {
  if (!state.projects.length) {
    els.projectList.innerHTML = '<div class="empty-state compact">还没有项目。</div>';
    return;
  }

  els.projectList.innerHTML = state.projects
    .map((project) => {
      const active = project.id === state.selectedProjectId ? "active" : "";
      const total = Number(project.total_tasks || 0);
      const completed = Number(project.completed_tasks || 0);
      const blocked = Number(project.blocked_tasks || 0);
      const ratio = total ? Math.round((completed / total) * 100) : 0;
      const dateText = `${project.start_date || "未定"}${project.due_date ? ` -> ${project.due_date}` : ""}`;
      return `
        <article class="project-item ${active}" data-project-id="${project.id}">
          <div class="project-item-head">
            <div class="project-title-wrap">
              <h3>${escapeHtml(project.name)}</h3>
              <p class="project-item-meta">${escapeHtml(project.description || "暂无描述")}</p>
            </div>
            <button type="button" class="project-delete-button" data-delete-project-id="${project.id}" aria-label="删除项目 ${escapeHtml(project.name)}">删除</button>
          </div>
          <div class="sidebar-chip-row">
            <span class="chip">任务 · ${completed}/${total}</span>
            ${blocked ? `<span class="chip warning">阻塞 · ${blocked}</span>` : ""}
          </div>
          <div class="project-progress"><span style="width:${ratio}%"></span></div>
          <small>${dateText}</small>
        </article>
      `;
    })
    .join("");
}

function renderDeletedProjectList() {
  if (!els.deletedProjectList || !els.deletedProjectWrap || !els.historyToggleBtn || !els.historyToggleCount) {
    return;
  }

  const count = state.deletedProjects.length;
  els.historyToggleCount.textContent = String(count);
  els.historyToggleBtn.disabled = count === 0;
  els.historyToggleBtn.setAttribute("aria-expanded", count > 0 && state.historyExpanded ? "true" : "false");
  els.historyToggleBtn.classList.toggle("active", count > 0 && state.historyExpanded);
  els.deletedProjectWrap.hidden = !(count > 0 && state.historyExpanded);

  if (!count) {
    els.deletedProjectList.innerHTML = '<div class="empty-state compact">还没有历史项目。</div>';
    return;
  }

  els.deletedProjectList.innerHTML = state.deletedProjects
    .map((project) => {
      const deletedAt = formatDateTimeLabel(project.deleted_at);
      return `
        <article class="project-item deleted">
          <div class="project-item-head">
            <div class="project-title-wrap">
              <h3>${escapeHtml(project.name)}</h3>
              <p class="project-item-meta">${escapeHtml(project.description || "暂无描述")}</p>
            </div>
            <button type="button" class="project-restore-button" data-restore-project-id="${project.id}" aria-label="恢复项目 ${escapeHtml(project.name)}">恢复</button>
          </div>
          <div class="sidebar-chip-row">
            <span class="chip">任务 · ${Number(project.total_tasks || 0)}</span>
          </div>
          <small>删除于 ${deletedAt}</small>
        </article>
      `;
    })
    .join("");
}

function renderSnapshotList() {
  const project = getCurrentProject();
  if (!els.snapshotList) {
    return;
  }
  if (!project) {
    els.snapshotList.innerHTML = '<div class="empty-state compact">先选择一个项目，再查看可回退的历史快照。</div>';
    return;
  }
  if (state.snapshotBusy && !state.snapshots.length) {
    els.snapshotList.innerHTML = '<div class="empty-state compact">正在读取历史快照…</div>';
    return;
  }
  if (!state.snapshots.length) {
    els.snapshotList.innerHTML = '<div class="empty-state compact">当前项目还没有历史快照。后续每次改动后会自动生成。</div>';
    return;
  }

  els.snapshotList.innerHTML = state.snapshots
    .map((snapshot) => `
      <article class="snapshot-item">
        <div class="snapshot-item-head">
          <div>
            <strong>${escapeHtml(formatDateTimeLabel(snapshot.created_at))}</strong>
            <p class="snapshot-meta">${escapeHtml(snapshot.action || "snapshot")} · ${Number(snapshot.task_count || 0)} 个任务</p>
            <p class="snapshot-summary">${escapeHtml(snapshot.summary || "自动快照")}</p>
          </div>
          <button type="button" class="primary" data-snapshot-restore="${snapshot.id}">恢复到此时</button>
        </div>
      </article>
    `)
    .join("");

  els.snapshotList.querySelectorAll("[data-snapshot-restore]").forEach((button) => {
    button.addEventListener("click", () => handleRestoreSnapshot(Number(button.dataset.snapshotRestore)));
  });
}
function renderProjectSummary() {
  const project = getCurrentProject();
  const stats = state.detail?.stats;
  if (!project || !stats) {
    els.projectSummary.innerHTML = '<div class="empty-state">先创建或选择一个项目。</div>';
    return;
  }

  const statCards = [
    ["任务进度", `${stats.completed_tasks}/${stats.total_tasks}`],
    ["预估工时", `${stats.estimated_hours} h`],
    ["实际工时", `${stats.actual_hours} h`],
    ["估时可信度", `${Math.round((stats.average_confidence || 0) * 100)}%`],
  ];

  els.projectSummary.innerHTML = `
    <div class="summary-layout summary-layout-simple">
      <div class="summary-copy">
        <p class="eyebrow">项目概览</p>
        <h2>${escapeHtml(project.name)}</h2>
        <p class="summary-subtext">${escapeHtml(project.description || "暂无描述")}</p>
      </div>
      <div class="stats-grid summary-stats-grid">
        ${statCards
          .map(
            ([label, value]) => `
              <div class="stat-card">
                <span class="stat-label">${label}</span>
                <span class="stat-value">${value}</span>
              </div>
            `
          )
          .join("")}
      </div>
    </div>
  `;
}

function renderTaskTable() {
  const tasks = getCurrentTasks();
  if (!tasks.length) {
    els.taskTableBody.innerHTML = '<tr><td colspan="8" class="empty-state">当前项目还没有任务，可以用顶部按钮新增或导入。</td></tr>';
    return;
  }

  els.taskTableBody.innerHTML = tasks
    .map((task) => {
      const active = task.id === state.selectedTaskId ? "active" : "";
      return `
        <tr class="task-row ${active}" data-task-id="${task.id}">
          <td>${escapeHtml(task.title)}</td>
          <td><span class="status-chip status-${task.status}">${formatStatus(task.status)}</span></td>
          <td>${escapeHtml(task.owner || "-")}</td>
          <td>${task.start_date || "-"}</td>
          <td>${task.end_date || "-"}</td>
          <td>${task.estimate_hours} h</td>
          <td>${task.progress}%</td>
          <td>${task.dependency_titles?.length ? escapeHtml(task.dependency_titles.join(", ")) : "-"}</td>
        </tr>
      `;
    })
    .join("");
}

function renderTaskEditor() {
  const task = getSelectedTask();
  els.deleteTaskBtn.disabled = !task;
  els.taskEditorSummary.textContent = task ? `编辑任务 · ${task.title}` : "编辑当前任务";

  if (!task) {
    els.editorState.innerHTML = '<div class="empty-state compact">先在上方表格或甘特图里选择一个任务。</div>';
    els.taskForm.reset();
    return;
  }

  els.editorState.innerHTML = `
    <div class="editor-state-card">
      <span class="status-chip status-${task.status}">${formatStatus(task.status)}</span>
      <span class="chip">时间 ${task.start_date || "未定"} -> ${task.end_date || "未定"}</span>
      <span class="chip">预估 ${task.estimate_hours} h</span>
      <span class="chip">依赖 ${(task.dependency_ids || []).length} 个</span>
    </div>
  `;

  const form = els.taskForm.elements;
  for (const [key, value] of Object.entries(task)) {
    if (form[key]) {
      form[key].value = value ?? "";
    }
  }
  form.task_id.value = task.id;
  form.dependency_ids.value = (task.dependency_ids || []).join(",");
}

function renderTableMeta() {
  const tasks = getCurrentTasks();
  const task = getSelectedTask();
  if (!tasks.length) {
    els.tableMeta.innerHTML = '<div class="empty-state compact">没有任务时，这里会显示任务数量、进度和当前选择。</div>';
    return;
  }

  const doneCount = tasks.filter((item) => item.status === "done").length;
  const progressAvg = Math.round(tasks.reduce((sum, item) => sum + (item.progress || 0), 0) / tasks.length);
  els.tableMeta.innerHTML = `
    <span class="meta-pill"><strong>${tasks.length}</strong> 个任务</span>
    <span class="meta-pill"><strong>${doneCount}</strong> 个已完成</span>
    <span class="meta-pill"><strong>${progressAvg}%</strong> 平均进度</span>
    <span class="meta-pill ${task ? "active" : ""}">${task ? `当前选择: ${escapeHtml(task.title)}` : "当前未选择任务"}</span>
  `;
}

function renderGanttMeta() {
  const project = getCurrentProject();
  const stats = state.detail?.stats;
  const task = getSelectedTask();
  if (!project || !stats) {
    els.ganttMeta.innerHTML = '';
    return;
  }

  const selectedText = task
    ? `当前高亮: ${escapeHtml(task.title)} (${task.start_date || "未定"} -> ${task.end_date || "未定"})`
    : "点击条形或表格行可高亮对应任务";

  els.ganttMeta.innerHTML = `
    <span class="meta-pill active">项目周期 ${project.start_date || "未定"} -> ${project.due_date || stats.projected_finish || "未定"}</span>
    <span class="meta-pill">共 ${stats.total_tasks} 个任务</span>
    <span class="meta-pill">${selectedText}</span>
  `;
}

function renderGantt() {
  const tasks = getCurrentTasks();
  if (!tasks.length) {
    els.ganttSvg.setAttribute("width", 760);
    els.ganttSvg.setAttribute("height", 180);
    els.ganttSvg.setAttribute("viewBox", "0 0 760 180");
    els.ganttSvg.innerHTML = `
      <rect x="0" y="0" width="760" height="180" rx="20" fill="rgba(255,255,255,0.94)" />
      <text x="380" y="94" text-anchor="middle" fill="#60728a" font-size="16">当前项目还没有可显示的任务，先新增或导入任务即可</text>
    `;
    return;
  }

  const starts = tasks.map((task) => parseDate(task.start_date)).filter(Boolean);
  const ends = tasks.map((task) => parseDate(task.end_date)).filter(Boolean);
  const rangeStart = new Date(Math.min(...starts));
  const rangeEnd = new Date(Math.max(...ends));
  const totalDays = Math.max(1, diffDays(rangeStart, rangeEnd) + 1);
  const longestTitle = Math.max(...tasks.map((task) => String(task.title || "").length), 8);
  const labelWidth = clamp(longestTitle * 12 + 56, 220, 320);
  const viewportWidth = Math.max(els.ganttViewport.clientWidth || 980, 980);
  const dayWidth = clamp(Math.round(760 / Math.max(totalDays, 12)), 34, 56);
  const width = Math.max(viewportWidth - 12, labelWidth + totalDays * dayWidth + 42);
  const height = HEADER_HEIGHT + tasks.length * ROW_HEIGHT + 24;
  const weekdays = ["日", "一", "二", "三", "四", "五", "六"];

  const monthHeaders = [];
  const dayHeaders = [];
  const dayBands = [];

  const pushMonthHeader = (startIndex, endIndex, monthNumber) => {
    const segmentWidth = (endIndex - startIndex) * dayWidth;
    const x = labelWidth + startIndex * dayWidth + 4;
    const usableWidth = Math.max(42, segmentWidth - 8);
    monthHeaders.push(`
      <g>
        <rect x="${x}" y="${MONTH_HEADER_Y}" width="${usableWidth}" height="${MONTH_HEADER_HEIGHT}" rx="8" fill="rgba(255,255,255,0.86)" />
        <text x="${x + usableWidth / 2}" y="${MONTH_HEADER_Y + 12}" text-anchor="middle" fill="#38516d" font-size="11" font-weight="700">${monthNumber}月</text>
      </g>
    `);
  };

  let segmentStart = 0;
  for (let index = 0; index < totalDays; index += 1) {
    const current = addDays(rangeStart, index);
    const x = labelWidth + index * dayWidth;
    const isWeekend = current.getDay() === 0 || current.getDay() === 6;
    const isBoundary = current.getDate() === 1 || index === 0;
    const bandFill = isWeekend
      ? "rgba(11,63,107,0.05)"
      : index % 2 === 0
        ? "rgba(13,45,76,0.018)"
        : "rgba(255,255,255,0)";
    const cellFill = isWeekend ? "rgba(11,63,107,0.09)" : "rgba(255,255,255,0.72)";
    const cellStroke = isBoundary ? "rgba(11,63,107,0.18)" : "rgba(13,45,76,0.05)";
    const weekdayColor = isWeekend ? "#0b3f6b" : "#73849a";

    dayBands.push(`
      <rect x="${x}" y="${HEADER_HEIGHT}" width="${dayWidth}" height="${height - HEADER_HEIGHT}" fill="${bandFill}" />
    `);

    dayHeaders.push(`
      <g>
        <rect x="${x + 2}" y="${DAY_HEADER_Y}" width="${Math.max(24, dayWidth - 4)}" height="${DAY_HEADER_HEIGHT}" rx="10" fill="${cellFill}" stroke="${cellStroke}" stroke-width="${isBoundary ? 1.1 : 0.8}" />
        <text x="${x + dayWidth / 2}" y="${DAY_HEADER_Y + 14}" text-anchor="middle" fill="#082a4a" font-size="12" font-weight="700">${current.getDate()}</text>
        <text x="${x + dayWidth / 2}" y="${DAY_HEADER_Y + 28}" text-anchor="middle" fill="${weekdayColor}" font-size="9">${weekdays[current.getDay()]}</text>
      </g>
    `);

    const next = addDays(current, 1);
    if (index === totalDays - 1 || next.getMonth() !== current.getMonth()) {
      pushMonthHeader(segmentStart, index + 1, current.getMonth() + 1);
      segmentStart = index + 1;
    }
  }

  const rows = [];
  const bars = [];
  tasks.forEach((task, index) => {
    const rowY = HEADER_HEIGHT + index * ROW_HEIGHT;
    const startOffset = diffDays(rangeStart, task.start_date);
    const duration = Math.max(1, diffDays(task.start_date, task.end_date) + 1);
    const barX = labelWidth + startOffset * dayWidth + 5;
    const barY = rowY + 11;
    const barWidth = Math.max(20, duration * dayWidth - 10);
    const isSelected = task.id === state.selectedTaskId;
    const barColor = statusColor(task.status, isSelected);

    rows.push(`
      <g>
        ${isSelected ? `<rect x="10" y="${rowY + 2}" width="${width - 20}" height="${ROW_HEIGHT - 4}" rx="14" fill="rgba(11,63,107,0.06)" />` : ""}
        <text x="18" y="${rowY + 20}" fill="#082a4a" font-size="13" font-weight="700">${escapeHtml(task.title)}</text>
        <text x="18" y="${rowY + 36}" fill="#60728a" font-size="11">${task.estimate_hours}h · ${formatStatus(task.status)}</text>
      </g>
    `);

    bars.push(`
      <g class="gantt-bar" data-bar-id="${task.id}" data-task-id="${task.id}">
        <title>${escapeHtml(task.title)} (${task.start_date} -> ${task.end_date})</title>
        <rect x="${barX}" y="${barY}" width="${barWidth}" height="24" rx="12" ry="12" fill="${barColor}" stroke="${isSelected ? "rgba(8,42,74,0.55)" : "rgba(13,45,76,0.08)"}" stroke-width="${isSelected ? 1.6 : 1}" />
      </g>
    `);
  });

  els.ganttSvg.setAttribute("width", width);
  els.ganttSvg.setAttribute("height", height);
  els.ganttSvg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  els.ganttSvg.innerHTML = `
    <defs>
      <filter id="softGlow" x="-20%" y="-40%" width="160%" height="200%">
        <feDropShadow dx="0" dy="8" stdDeviation="8" flood-color="rgba(13,45,76,0.12)" />
      </filter>
    </defs>
    <rect x="0" y="0" width="${width}" height="${height}" rx="22" fill="rgba(255,255,255,0.96)" />
    <rect x="0" y="0" width="${width}" height="${HEADER_HEIGHT}" rx="22" fill="rgba(246,249,253,0.96)" />
    <rect x="0" y="0" width="${labelWidth}" height="${height}" rx="22" fill="rgba(239,245,252,0.96)" />
    <text x="18" y="28" fill="#60728a" font-size="11" font-weight="700">任务</text>
    <text x="18" y="46" fill="#7b8da5" font-size="10">预估 · 状态</text>
    ${dayBands.join("")}
    ${monthHeaders.join("")}
    ${dayHeaders.join("")}
    ${rows.join("")}
    <g filter="url(#softGlow)">${bars.join("")}</g>
  `;

  els.ganttSvg.querySelectorAll("[data-bar-id]").forEach((node) => {
    node.addEventListener("click", () => selectTask(Number(node.dataset.taskId)));
  });
}

function statusColor(status, isSelected = false) {
  if (isSelected) {
    return "#2f74b8";
  }
  return {
    planned: "#6e8093",
    in_progress: "#2f74b8",
    blocked: "#c95f67",
    done: "#18855b",
  }[status] || "#6e8093";
}

async function handleProjectCreate(mode) {
  const name = els.projectForm.elements.name.value.trim();
  if (!name) {
    showToast("\u8bf7\u8f93\u5165\u9879\u76ee\u540d\u79f0", true);
    return;
  }

  const description = els.projectForm.elements.description.value.trim();
  const useLlmProgress = mode === "smart";
  if (useLlmProgress && !description) {
    showToast("\u8bf7\u5148\u7c98\u8d34\u9879\u76ee\u63cf\u8ff0\u6216\u8def\u7ebf\u56fe\u6587\u672c", true);
    return;
  }

  if (useLlmProgress) {
    startLlmProgress();
  } else {
    resetLlmProgress();
  }

  const payload = {
    name,
    description,
    start_date: els.projectForm.elements.start_date.value,
    due_date: els.projectForm.elements.due_date.value || null,
    use_suggestions: useLlmProgress,
  };

  try {
    const detail = await api("/api/projects", {
      method: "POST",
      body: JSON.stringify(payload),
    });

    const analysis = detail.analysis || {};
    const importedCount = detail.imported_task_count || detail.tasks?.length || 0;
    if (useLlmProgress) {
      const progressMessage = importedCount
        ? `${analysis.source === "llm" ? `\u5df2\u5b8c\u6210 ${analysis.model || "LLM"} \u5206\u6790` : "\u5df2\u5b8c\u6210\u9879\u76ee\u89e3\u6790"}\uff0c\u5df2\u751f\u6210 ${importedCount} \u4e2a\u521d\u59cb\u4efb\u52a1\u2026`
        : (analysis.note || "\u89e3\u6790\u5b8c\u6210\uff0c\u6b63\u5728\u6253\u5f00\u9879\u76ee\u2026");
      await finishLlmProgress(progressMessage);
    }

    acceptDetail(detail);
    await refreshProjects();
    renderAll();
    closeDialog(els.projectDialog);
    els.projectForm.reset();
    setDefaultDates();

    let message = useLlmProgress ? "\u9879\u76ee\u5df2\u521b\u5efa\u5e76\u5bfc\u5165\u521d\u59cb\u4efb\u52a1" : "\u7a7a\u9879\u76ee\u5df2\u521b\u5efa";
    if (useLlmProgress) {
      if (importedCount) {
        message = `${analysis.source === "llm" ? `\u9879\u76ee\u5df2\u521b\u5efa\uff0c\u5df2\u4f7f\u7528 ${analysis.model || "LLM"} \u751f\u6210` : "\u9879\u76ee\u5df2\u521b\u5efa\uff0c\u5df2\u89e3\u6790\u5e76\u5bfc\u5165"} ${importedCount} \u4e2a\u4efb\u52a1`;
      } else if (analysis.source === "rules" || analysis.source === "structured") {
        message = analysis.note || "\u9879\u76ee\u5df2\u521b\u5efa\uff0c\u5df2\u56de\u9000\u5230\u672c\u5730\u89e3\u6790";
      }
    }
    showToast(message);
  } catch (error) {
    if (useLlmProgress) {
      failLlmProgress(error.message);
    }
    showToast(error.message, true);
  }
}


async function handleMeetingUpdateSubmit(event) {
  event.preventDefault();
  const project = getCurrentProject();
  if (!project) {
    showToast("请先选择一个项目", true);
    return;
  }

  const meetingText = els.meetingForm.elements.meeting_text.value.trim();
  if (!meetingText) {
    showToast("请先粘贴本次会议纪要", true);
    return;
  }

  const autoSchedule = Boolean(els.meetingForm.elements.auto_schedule.checked);
  setMeetingBusy(true);
  try {
    const detail = await api(`/api/projects/${project.id}/meeting-update`, {
      method: "POST",
      body: JSON.stringify({
        meeting_text: meetingText,
        auto_schedule: autoSchedule,
      }),
    });

    acceptDetail(detail);
    await refreshProjects();
    renderAll();

    const updatedCount = Number(detail.meeting_update?.updated_count || 0);
    const skippedCount = Number(detail.meeting_update?.skipped_count || 0);
    const summary = detail.meeting_update?.summary || "";

    setMeetingBusy(false);
    closeDialog(els.meetingDialog);
    els.meetingForm.reset();

    let message = updatedCount ? `会议更新完成：已更新 ${updatedCount} 个任务` : "会议更新完成";
    if (skippedCount) {
      message += `，跳过 ${skippedCount} 条`;
    }
    if (!updatedCount && summary) {
      message = summary;
    }
    showToast(message);
  } catch (error) {
    const friendlyMessage = error.message === "unknown endpoint"
      ? "当前后端还没更新到会议纪要更新接口，请重启 python app.py 后再试"
      : error.message;
    showToast(friendlyMessage, true);
  } finally {
    if (state.meetingBusy) {
      setMeetingBusy(false);
    }
  }
}
async function handleImportSubmit(event) {
  event.preventDefault();
  const file = els.importForm.elements.import_file.files[0];
  if (!file) {
    showToast("请选择文件", true);
    return;
  }

  setImportBusy(true);
  try {
    const projectName = els.importForm.elements.project_name.value.trim();
    const payload = {
      project_id: projectName ? null : state.selectedProjectId,
      project_name: projectName || null,
      format: file.name.split(".").pop().toLowerCase(),
      file_name: file.name,
      content_base64: await readFileAsBase64(file),
      start_date: els.importForm.elements.start_date.value || null,
      replace_existing: els.importForm.elements.replace_existing.checked,
    };

    const detail = await api("/api/import", {
      method: "POST",
      body: JSON.stringify(payload),
    });
    acceptDetail(detail);
    await refreshProjects();
    renderAll();
    closeDialog(els.importDialog);
    els.importForm.reset();
    setDefaultDates();
    showToast(`已导入 ${file.name}`);
  } catch (error) {
    showToast(error.message, true);
  } finally {
    setImportBusy(false);
  }
}

async function handleSmartImportSubmit() {
  const project = getCurrentProject();
  if (!project) {
    showToast("请先选中一个项目，再追加解析任务", true);
    return;
  }

  const text = els.importForm.elements.smart_text.value.trim();
  if (!text) {
    showToast("请先粘贴要解析的文本", true);
    return;
  }

  const replaceExisting = Boolean(els.importForm.elements.replace_existing.checked);

  startImportProgress();
  setImportBusy(true);
  try {
    const detail = await api(`/api/projects/${project.id}/smart-import`, {
      method: "POST",
      body: JSON.stringify({
        description: text,
        replace_existing: replaceExisting,
      }),
    });
    const importedCount = detail.imported_task_count || 0;
    const analysis = detail.analysis || {};
    const progressMessage = importedCount
      ? `已为当前项目新增 ${importedCount} 个任务，正在刷新甘特图…`
      : (analysis.note || "解析完成，正在刷新甘特图…");
    await finishImportProgress(progressMessage);
    acceptDetail(detail);
    await refreshProjects();
    renderAll();
    closeDialog(els.importDialog);
    els.importForm.reset();
    setDefaultDates();
    showToast(importedCount ? `已为当前项目新增 ${importedCount} 个任务` : "已追加解析任务");
  } catch (error) {
    const friendlyMessage = error.message === "unknown endpoint"
      ? "当前后端还没更新到智能追加任务接口，请重启 python app.py 后再试"
      : error.message;
    failImportProgress(friendlyMessage);
    showToast(friendlyMessage, true);
  } finally {
    setImportBusy(false);
  }
}

async function handleOpenSnapshotDialog() {
  const project = getCurrentProject();
  if (!project) {
    showToast("请先选择一个项目", true);
    return;
  }

  state.snapshots = [];
  renderSnapshotList();
  openDialog(els.snapshotDialog);
  setSnapshotBusy(true);
  try {
    const data = await api(`/api/projects/${project.id}/snapshots`);
    state.snapshots = data.snapshots || [];
    renderSnapshotList();
  } catch (error) {
    showToast(error.message, true);
    renderSnapshotList();
  } finally {
    setSnapshotBusy(false);
  }
}

async function handleRestoreSnapshot(snapshotId) {
  const project = getCurrentProject();
  if (!project) {
    showToast("请先选择一个项目", true);
    return;
  }
  const snapshot = state.snapshots.find((item) => Number(item.id) === Number(snapshotId));
  if (!snapshot) {
    showToast("未找到要恢复的快照", true);
    return;
  }
  if (!window.confirm(`确认将整个甘特图恢复到 ${formatDateTimeLabel(snapshot.created_at)} 吗？当前任务排期会被覆盖。`)) {
    return;
  }

  setSnapshotBusy(true);
  try {
    const detail = await api(`/api/projects/${project.id}/snapshots/${snapshotId}/restore`, {
      method: "POST",
      body: JSON.stringify({}),
    });
    acceptDetail(detail);
    await refreshProjects();
    renderAll();
    closeDialog(els.snapshotDialog);
    showToast(`已恢复到 ${formatDateTimeLabel(snapshot.created_at)} 的快照`);
  } catch (error) {
    showToast(error.message, true);
  } finally {
    setSnapshotBusy(false);
  }
}
async function handleAddTask() {
  if (!state.selectedProjectId) {
    showToast("请先创建或选择项目", true);
    return;
  }

  const project = getCurrentProject();
  try {
    const detail = await api("/api/tasks", {
      method: "POST",
      body: JSON.stringify({
        project_id: state.selectedProjectId,
        title: "新任务",
        start_date: project?.start_date,
        end_date: project?.start_date,
        estimate_hours: 6,
        estimate_basis: "默认模板",
        confidence: 0.56,
      }),
    });
    acceptDetail(detail);
    state.selectedTaskId = detail.tasks.at(-1)?.id ?? state.selectedTaskId;
    els.taskEditorDetails.open = true;
    await refreshProjects();
    renderAll();
    showToast("已新增任务");
  } catch (error) {
    showToast(error.message, true);
  }
}

async function handleReschedule() {
  if (!state.selectedProjectId) {
    return;
  }

  try {
    const detail = await api(`/api/projects/${state.selectedProjectId}/reschedule`, {
      method: "POST",
      body: JSON.stringify({}),
    });
    acceptDetail(detail);
    await refreshProjects();
    renderAll();
    showToast("已按依赖重算排期");
  } catch (error) {
    showToast(error.message, true);
  }
}

async function saveTask(autoSchedule) {
  if (!state.selectedProjectId) {
    showToast("请先选择项目", true);
    return;
  }

  const payload = readTaskForm();
  if (!payload.title) {
    showToast("任务标题不能为空", true);
    return;
  }

  try {
    const taskId = Number(els.taskForm.elements.task_id.value || 0);
    const detail = await api(taskId ? `/api/tasks/${taskId}` : "/api/tasks", {
      method: taskId ? "PUT" : "POST",
      body: JSON.stringify({
        ...payload,
        project_id: state.selectedProjectId,
        auto_schedule: autoSchedule,
      }),
    });
    acceptDetail(detail);
    if (taskId) {
      state.selectedTaskId = taskId;
    }
    await refreshProjects();
    renderAll();
    showToast(autoSchedule ? "已保存并重算排期" : "任务已保存");
  } catch (error) {
    showToast(error.message, true);
  }
}

async function handleDeleteTask() {
  const task = getSelectedTask();
  if (!task) {
    return;
  }
  if (!window.confirm(`确认删除任务“${task.title}”吗？`)) {
    return;
  }

  try {
    const detail = await api(`/api/tasks/${task.id}`, {
      method: "DELETE",
      body: JSON.stringify({}),
    });
    acceptDetail(detail);
    await refreshProjects();
    renderAll();
    showToast("任务已删除");
  } catch (error) {
    showToast(error.message, true);
  }
}

function exportCurrentProject(format) {
  if (!state.selectedProjectId) {
    showToast("请先选择项目", true);
    return;
  }
  window.open(`/api/export?project_id=${state.selectedProjectId}&format=${format}`, "_blank");
}

function exportAllProjectsXlsx() {
  if (!state.projects.length) {
    showToast("当前没有可导出的项目", true);
    return;
  }
  window.open("/api/export-all?format=xlsx", "_blank");
}

function readTaskForm() {
  const form = els.taskForm.elements;
  return {
    title: form.title.value.trim(),
    description: form.description.value.trim(),
    status: form.status.value,
    owner: form.owner.value.trim(),
    priority: Number(form.priority.value || 2),
    complexity: Number(form.complexity.value || 3),
    progress: Number(form.progress.value || 0),
    estimate_hours: Number(form.estimate_hours.value || 0),
    actual_hours: Number(form.actual_hours.value || 0),
    start_date: form.start_date.value || null,
    end_date: form.end_date.value || null,
    confidence: Number(form.confidence.value || 0.56),
    dependency_ids: form.dependency_ids.value
      .split(/[，,、]/)
      .map((item) => Number(item.trim()))
      .filter((item) => item > 0),
    estimate_basis: form.estimate_basis.value.trim(),
    notes: form.notes.value.trim(),
  };
}

function readFileAsBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || "").split(",").pop());
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function showToast(message, isError = false) {
  els.toast.textContent = message;
  els.toast.style.background = isError ? "rgba(128, 43, 52, 0.94)" : "rgba(8, 42, 74, 0.94)";
  els.toast.classList.add("show");
  clearTimeout(state.toastTimer);
  state.toastTimer = setTimeout(() => {
    els.toast.classList.remove("show");
  }, 2400);
}

function formatDateInput(value) {
  const date = value instanceof Date ? value : new Date(value);
  return date.toISOString().slice(0, 10);
}

function addDays(value, days) {
  const date = value instanceof Date ? new Date(value) : new Date(`${String(value).slice(0, 10)}T00:00:00`);
  date.setDate(date.getDate() + days);
  return date;
}

function parseDate(value) {
  if (!value) {
    return null;
  }
  const parsed = new Date(`${String(value).slice(0, 10)}T00:00:00`);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function diffDays(start, end) {
  const startDate = start instanceof Date ? start : parseDate(start);
  const endDate = end instanceof Date ? end : parseDate(end);
  if (!startDate || !endDate) {
    return 0;
  }
  return Math.round((endDate - startDate) / 86400000);
}

function formatDateTimeLabel(value) {
  if (!value) {
    return "未记录";
  }
  return String(value).replace("T", " ").slice(0, 16);
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function formatStatus(status) {
  return {
    planned: "计划中",
    in_progress: "进行中",
    blocked: "阻塞",
    done: "已完成",
  }[status] || status;
}

function formatHealth(health) {
  return {
    planned: "待启动",
    active: "推进中",
    at_risk: "有风险",
    done: "已完成",
  }[health] || health;
}

function formatAnalysisSource(source, model = "") {
  if (source === "llm") {
    return model ? `LLM - ${model}` : "LLM 智能分析";
  }
  if (source === "structured") {
    return "结构化解析";
  }
  if (source === "rules") {
    return "规则拆分";
  }
  if (source === "manual") {
    return "手动创建";
  }
  return source || "未标记";
}

function formatProjectCategory(category) {
  return {
    general: "通用项目",
    web_tool: "Web 工具",
    data_project: "数据项目",
    content_project: "内容项目",
  }[category] || category;
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function debounce(fn, wait) {
  let timer = null;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), wait);
  };
}






















