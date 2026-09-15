(function () {
  "use strict";

  const PAGE = "daily";
  const life = GovUI.beginPageLifecycle();
  let uploading = false;

  const state = {
    uploadedFiles: [],
    jobId: null,
    running: false,
    eventSource: null,
    excludeYears: [],
    excludeCurrent: {},
    excludePendingFile: null,
    defaultExcludeYear: 2026,
    batchId: null,
    databaseMode: "excel",
  };

  const dropZone = document.getElementById("dropZone");
  const fileInput = document.getElementById("fileInput");
  const uploadList = document.getElementById("uploadList");
  const importPreview = document.getElementById("importPreview");
  const importMode = document.getElementById("importMode");
  const importStartDate = document.getElementById("importStartDate");
  const importEndDate = document.getElementById("importEndDate");
  const importCounts = document.getElementById("importCounts");
  const importWarnings = document.getElementById("importWarnings");
  const importActivateBtn = document.getElementById("importActivateBtn");
  const importRejectBtn = document.getElementById("importRejectBtn");
  const excludeYear = document.getElementById("excludeYear");
  const excludeCurrent = document.getElementById("excludeCurrent");
  const excludeDropZone = document.getElementById("excludeDropZone");
  const excludeFileInput = document.getElementById("excludeFileInput");
  const excludeUploadBtn = document.getElementById("excludeUploadBtn");
  const excludePending = document.getElementById("excludePending");
  const runBtn = document.getElementById("runBtn");
  const stopBtn = document.getElementById("stopBtn");
  const resetBtn = document.getElementById("resetBtn");
  const runStatus = document.getElementById("runStatus");
  const frontendLog = document.getElementById("frontendLog");
  const scriptLog = document.getElementById("scriptLog");
  const resultSection = document.getElementById("resultSection");
  const resultPlaceholder = document.getElementById("resultPlaceholder");
  const completeBanner = document.getElementById("completeBanner");
  const reportComparison = document.getElementById("reportComparison");
  const comparisonPrevious = document.getElementById("comparisonPrevious");
  const comparisonUsers = document.getElementById("comparisonUsers");
  const comparisonLines = document.getElementById("comparisonLines");
  const comparisonCategories = document.getElementById("comparisonCategories");
  const summarySimple = document.getElementById("summarySimple");
  const summaryFull = document.getElementById("summaryFull");
  const dlResult = document.getElementById("dlResult");
  const dlPublish = document.getElementById("dlPublish");
  const dlStats = document.getElementById("dlStats");
  const dlSummarySimple = document.getElementById("dlSummarySimple");
  const dlSummaryFull = document.getElementById("dlSummaryFull");
  const dlAll = document.getElementById("dlAll");
  const statusBadge = document.getElementById("statusBadge");

  const progressTracker = GovUI.createProgressTracker({
    progressBar: document.getElementById("progressBar"),
    elapsedText: document.getElementById("elapsedText"),
    etaText: document.getElementById("etaText"),
    pctText: document.getElementById("pctText"),
    totalTimeText: document.getElementById("totalTimeText"),
  });

  const logFrontend = GovUI.createLogger(frontendLog, PAGE);

  function updateExcludeUploadBtn() {
    excludeUploadBtn.disabled = !state.excludePendingFile;
  }

  function formatExcludeLabel(name) {
    return String(name || "").replace(/表11/g, "").trim();
  }

  function renderExcludeCurrent() {
    const year = excludeYear.value;
    const info = state.excludeCurrent[year];
    if (info && info.filename) {
      const sizeKb = (info.size / 1024).toFixed(1);
      excludeCurrent.textContent = `当前最新：${formatExcludeLabel(info.filename)}（${sizeKb} KB）`;
    } else {
      excludeCurrent.textContent = "当前最新：—（目录暂无文件）";
    }
  }

  function populateExcludeYears(years, defaultYear) {
    excludeYear.innerHTML = "";
    years.forEach(function (y) {
      const opt = document.createElement("option");
      opt.value = String(y);
      opt.textContent = String(y);
      if (y === defaultYear) opt.selected = true;
      excludeYear.appendChild(opt);
    });
    renderExcludeCurrent();
  }

  async function loadExcludeYears() {
    try {
      const resp = await GovUI.fetchWithTimeout("/api/exclude/years", null, 8000);
      const data = await GovUI.parseJsonSafe(resp);
      if (!resp.ok) throw new Error(data.detail || "加载剔除配置失败");
      state.excludeYears = data.years || [];
      state.excludeCurrent = data.current || {};
      state.defaultExcludeYear = data.default_year || 2026;
      populateExcludeYears(state.excludeYears, state.defaultExcludeYear);
    } catch (err) {
      if (!GovUI.isGatewayError(err)) {
        logFrontend("加载剔除年份配置失败：" + err.message);
      }
      populateExcludeYears([2024, 2025, 2026, 2027, 2028, 2029, 2030], 2026);
    }
  }

  async function loadNewdataFiles() {
    // 不再列出服务器 Newdata 目录存量文件；列表仅展示本次会话成功导入项
  }

  function setSessionUploadedFile(fileInfo) {
    state.uploadedFiles = fileInfo ? [fileInfo] : [];
    GovUI.savePagePatch(PAGE, { uploadedFiles: state.uploadedFiles });
    renderUploadList();
  }

  function restoreSessionUploadedFiles() {
    if (life.reset) {
      setSessionUploadedFile(null);
      return;
    }
    const pageState = GovUI.loadPageState(PAGE);
    const files = Array.isArray(pageState.uploadedFiles) ? pageState.uploadedFiles : [];
    // 只保留最后一个
    state.uploadedFiles = files.length ? [files[files.length - 1]] : [];
    renderUploadList();
  }

  function setExcludePendingFile(file) {
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext !== "xlsx" && ext !== "xls") {
      logFrontend(`剔除文件须为 Excel：${file.name}`);
      return;
    }
    state.excludePendingFile = file;
    excludePending.textContent = `待上传：${formatExcludeLabel(file.name)}`;
    updateExcludeUploadBtn();
  }

  async function uploadExcludeFile() {
    if (!state.excludePendingFile) return;
    const year = excludeYear.value;
    const file = state.excludePendingFile;
    logFrontend(`正在上传剔除文件（${year} 年）：${file.name} ...`);
    try {
      const data = await GovUI.postFileWithRetry(
        `/api/upload/exclude?year=${encodeURIComponent(year)}`,
        file,
        { log: logFrontend }
      );
      state.excludePendingFile = null;
      excludePending.textContent = "";
      updateExcludeUploadBtn();
      excludeFileInput.value = "";
      logFrontend(
        `剔除文件上传成功：${year} 年 → ${formatExcludeLabel(data.filename)} (${(data.size / 1024).toFixed(1)} KB)`
      );
      await loadExcludeYears();
    } catch (err) {
      logFrontend(`剔除文件上传失败：${err.message}`);
    }
  }

  excludeYear.addEventListener("change", renderExcludeCurrent);
  excludeDropZone.addEventListener("dragover", function (e) {
    e.preventDefault();
    excludeDropZone.classList.add("dragover");
  });
  excludeDropZone.addEventListener("dragleave", function () {
    excludeDropZone.classList.remove("dragover");
  });
  excludeDropZone.addEventListener("drop", function (e) {
    e.preventDefault();
    excludeDropZone.classList.remove("dragover");
    const files = e.dataTransfer.files;
    if (files.length > 1) logFrontend("剔除文件每次仅支持 1 个，已取第一个文件。");
    if (files.length > 0) setExcludePendingFile(files[0]);
  });
  excludeFileInput.addEventListener("change", function () {
    const files = excludeFileInput.files;
    if (files.length > 1) logFrontend("剔除文件每次仅支持 1 个，已取第一个文件。");
    if (files.length > 0) setExcludePendingFile(files[0]);
    excludeFileInput.value = "";
  });
  excludeUploadBtn.addEventListener("click", uploadExcludeFile);

  function updateActionButtons() {
    runBtn.disabled = state.running || state.uploadedFiles.length === 0;
    stopBtn.disabled = !state.running;
    resetBtn.disabled = state.running;
  }

  function renderUploadList() {
    uploadList.innerHTML = "";
    state.uploadedFiles.forEach(function (f) {
      const li = document.createElement("li");
      const sizeKb = (f.size / 1024).toFixed(1);
      li.innerHTML = `${GovUI.escapeHtml(f.filename)}<span class="size">(${sizeKb} KB)</span>`;
      uploadList.appendChild(li);
    });
    updateActionButtons();
  }

  function hideImportPreview() {
    state.batchId = null;
    importPreview.hidden = true;
    importCounts.textContent = "";
    importWarnings.textContent = "";
  }

  async function loadImportPreview(batchId, databaseMode) {
    const query = new URLSearchParams();
    if (importStartDate.value) query.set("start_date", importStartDate.value);
    if (importEndDate.value) query.set("end_date", importEndDate.value);
    const suffix = query.toString() ? `?${query}` : "";
    const resp = await GovUI.fetchWithTimeout(`/api/imports/${batchId}/preview${suffix}`, null, 30000);
    const data = await GovUI.parseJsonSafe(resp);
    if (!resp.ok) throw new Error(data.detail || "差异预览失败");
    state.batchId = batchId;
    state.databaseMode = databaseMode || "shadow";
    importMode.textContent = state.databaseMode === "postgres" ? "正式库" : "影子库";
    importStartDate.value = data.replace_start_date || "";
    importEndDate.value = data.replace_end_date || "";
    importCounts.textContent = `新增 ${data.added} 条 · 删除 ${data.deleted} 条 · 修改 ${data.modified} 条 · 未变化 ${data.unchanged} 条`;
    const warningParts = [];
    if (data.missing_dates && data.missing_dates.length) {
      warningParts.push(`缺失日期 ${data.missing_dates.length} 天（激活后按 0 条处理）`);
    }
    const validation = data.validation || {};
    if (validation.duplicate_key_count) {
      warningParts.push(`重复键 ${validation.duplicate_key_count} 个，禁止激活`);
    }
    if (validation.errors && validation.errors.length) {
      warningParts.push(`校验错误 ${validation.errors.length} 条，禁止激活`);
    }
    importWarnings.textContent = warningParts.join("；") || "校验通过，可确认替换区间后激活。";
    importActivateBtn.disabled = data.status !== "pending_confirmation";
    importRejectBtn.disabled = !["pending_confirmation", "validation_error"].includes(data.status);
    importPreview.hidden = false;
  }

  importStartDate.addEventListener("change", function () {
    if (state.batchId) loadImportPreview(state.batchId, state.databaseMode).catch(err => logFrontend(err.message));
  });
  importEndDate.addEventListener("change", function () {
    if (state.batchId) loadImportPreview(state.batchId, state.databaseMode).catch(err => logFrontend(err.message));
  });
  importActivateBtn.addEventListener("click", async function () {
    if (!state.batchId) return;
    importActivateBtn.disabled = true;
    try {
      const resp = await fetch(`/api/imports/${state.batchId}/activate`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ start_date: importStartDate.value, end_date: importEndDate.value, confirmed_by: "admin" }),
      });
      const data = await GovUI.parseJsonSafe(resp);
      if (!resp.ok) throw new Error(data.detail || "激活失败");
      logFrontend(`数据库批次已激活：${state.batchId}`);
      importWarnings.textContent = "批次已激活，已写入审计记录。";
    } catch (err) {
      logFrontend("激活失败：" + err.message);
      importActivateBtn.disabled = false;
    }
  });
  importRejectBtn.addEventListener("click", async function () {
    if (!state.batchId) return;
    try {
      const resp = await fetch(`/api/imports/${state.batchId}/reject`, { method: "POST" });
      const data = await GovUI.parseJsonSafe(resp);
      if (!resp.ok) throw new Error(data.detail || "拒绝失败");
      logFrontend(`数据库批次已拒绝：${state.batchId}`);
      hideImportPreview();
    } catch (err) {
      logFrontend("拒绝失败：" + err.message);
    }
  });

  async function uploadFile(file) {
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext !== "xlsx" && ext !== "xls") {
      logFrontend(`跳过非 Excel 文件：${file.name}`);
      return false;
    }
    logFrontend(`正在上传：${file.name} ...`);
    try {
      const data = await GovUI.postFileWithRetry("/api/upload", file, { log: logFrontend });
      setSessionUploadedFile({ filename: data.filename, size: data.size });
      logFrontend(`导入成功：${data.filename} (${(data.size / 1024).toFixed(1)} KB)`);
      if (data.batch_id) {
        await loadImportPreview(data.batch_id, data.database_mode);
      } else {
        hideImportPreview();
        if (data.database_error) logFrontend(`数据库影子校验未完成：${data.database_error}`);
      }
      return true;
    } catch (err) {
      logFrontend(`上传失败：${file.name} — ${err.message}`);
      return false;
    }
  }

  async function handleFiles(fileList) {
    const files = Array.from(fileList || []);
    const excelFiles = files.filter(function (f) {
      const ext = f.name.split(".").pop().toLowerCase();
      return ext === "xlsx" || ext === "xls";
    });
    if (excelFiles.length === 0) {
      logFrontend("未选择有效的 Excel 文件（.xlsx / .xls）。");
      return;
    }
    if (excelFiles.length > 1) {
      logFrontend(`一次仅保留最后一个文件，已忽略前 ${excelFiles.length - 1} 个。`);
    }
    const last = excelFiles[excelFiles.length - 1];
    if (uploading) {
      logFrontend("已有文件正在上传，请稍候。");
      return;
    }
    uploading = true;
    try {
      await uploadFile(last);
    } finally {
      uploading = false;
    }
  }

  dropZone.addEventListener("dragover", function (e) {
    e.preventDefault();
    dropZone.classList.add("dragover");
  });
  dropZone.addEventListener("dragleave", function () {
    dropZone.classList.remove("dragover");
  });
  dropZone.addEventListener("drop", function (e) {
    e.preventDefault();
    dropZone.classList.remove("dragover");
    handleFiles(e.dataTransfer.files);
  });
  fileInput.addEventListener("change", function () {
    handleFiles(fileInput.files);
    fileInput.value = "";
  });

  GovUI.bindTabs(document);
  GovUI.bindCopyButtons(summarySimple, summaryFull, logFrontend);

  function setRunning(running) {
    state.running = running;
    updateActionButtons();
    GovUI.setBadge(statusBadge, running ? "running" : "idle");
    if (running) {
      runStatus.className = "run-status running";
      runStatus.innerHTML = '<span class="spinner"></span>脚本运行中...';
    }
    GovUI.savePagePatch(PAGE, { running: running, jobId: state.jobId });
  }

  function showResult(payload) {
    state.running = false;
    updateActionButtons();
    // 只展示本页 kind 的结果
    if (payload.kind && payload.kind !== "daily") {
      return;
    }
    GovUI.setBadge(statusBadge, GovUI.badgeFromStatus(payload.status));
    if (payload.status === "success" || payload.status === "failed" || payload.status === "cancelled") {
      progressTracker.finish(payload.duration_seconds);
    } else {
      progressTracker.reset();
    }
    if (resultSection) resultSection.style.display = "flex";
    GovUI.showResult({
      resultSection: resultSection,
      resultPlaceholder: resultPlaceholder,
      completeBanner: completeBanner,
      summarySimple: summarySimple,
      summaryFull: summaryFull,
      dlResult: dlResult,
      dlPublish: dlPublish,
      dlStats: dlStats,
      dlSummarySimple: dlSummarySimple,
      dlSummaryFull: dlSummaryFull,
      dlAll: dlAll,
      jobId: state.jobId,
      payload: payload,
      runStatus: runStatus,
    });
    if (reportComparison) {
      const diff = payload.status === "success" ? payload.comparison : null;
      reportComparison.hidden = !diff;
      if (diff) {
        comparisonPrevious.textContent = diff.available
          ? `上次报告：${String(diff.previous_report || "").split(/[\\/]/).pop()}；上次口径：${diff.previous_rule_version || "未记录"}；本次口径：${diff.rule_version || "—"}`
          : (diff.reason || "暂无可比上次报告");
        const users = (diff.metrics && diff.metrics["用户"]) || {};
        const lines = (diff.metrics && diff.metrics["线路"]) || {};
        comparisonUsers.textContent = diff.available
          ? `频繁停电用户统计表合计：上次 ${users.report_total_previous ?? users.previous_frequent} 户 → 本次 ${users.report_total_current ?? users.current_frequent} 户；净变化 ${users.report_total_delta ?? users.net_change}。实际退出频繁清单 ${users.left_frequent} 户，停电次数下降 ${users.decreased_outage_count} 户。`
          : "";
        comparisonLines.textContent = diff.available
          ? `频繁停电线路统计表合计：上次 ${lines.report_total_previous ?? lines.previous_frequent} 条 → 本次 ${lines.report_total_current ?? lines.current_frequent} 条；净变化 ${lines.report_total_delta ?? lines.net_change}。实际退出频繁清单 ${lines.left_frequent} 条。`
          : "";
        comparisonCategories.textContent = diff.available
          ? (diff.statistics_columns || []).map(function (item) {
              return `${item.column} ${item.label}：${item.previous} → ${item.current}（${item.delta > 0 ? "+" : ""}${item.delta}）`;
            }).join("\n")
          : "";
      }
    }
    GovUI.savePagePatch(PAGE, {
      jobId: state.jobId,
      lastResult: Object.assign({}, payload, { kind: "daily" }),
      finished: true,
    });
    if (!payload.summary_simple && !payload.summary_full && payload.status === "success") {
      logFrontend("未能自动提取摘要，请查看脚本日志。");
    }
  }

  function connectStream(jobId) {
    if (state.eventSource) state.eventSource.close();
    const streamOpts = {
      jobId: jobId,
      scriptLog: scriptLog,
      logFrontend: logFrontend,
      running: true,
      eventSource: null,
      kind: "daily",
      pageKind: PAGE,
      onRunningChange: function (v) {
        state.running = v;
        streamOpts.running = v;
        updateActionButtons();
        if (v) {
          GovUI.setBadge(statusBadge, "running");
        }
      },
      setProgressFn: function () {},
      onDone: function (payload) {
        showResult(payload);
      },
    };
    GovUI.connectStream(streamOpts);
    state.eventSource = streamOpts.eventSource;
  }

  runBtn.addEventListener("click", async function () {
    if (state.running) return;
    if (completeBanner) {
      completeBanner.style.display = "none";
      completeBanner.textContent = "";
    }
    if (resultPlaceholder) resultPlaceholder.style.display = "";
    if (reportComparison) reportComparison.hidden = true;
    summarySimple.value = "";
    summaryFull.value = "";
    GovUI.setDownloadLink(dlResult, null, "result", null);
    GovUI.setDownloadLink(dlPublish, null, "publish", null);
    GovUI.setDownloadLink(dlStats, null, "stats", null);
    GovUI.setDownloadLink(dlSummarySimple, null, "summary_simple", null);
    GovUI.setDownloadLink(dlSummaryFull, null, "summary_full", null);
    GovUI.setDownloadLink(dlAll, null, "all", null);
    scriptLog.textContent = "";
    progressTracker.reset();
    setRunning(true);
    logFrontend("正在启动脚本...");

    try {
      const expected = await GovUI.fetchExpectedSeconds("daily");
      const resp = await fetch("/api/run", { method: "POST" });
      const data = await resp.json();
      if (!resp.ok) throw new Error(data.detail || "启动失败");
      state.jobId = data.job_id;
      GovUI.saveSession({
        activeJobId: data.job_id,
        jobId: data.job_id,
        kind: "daily",
        status: "running",
        finished: false,
      });
      GovUI.savePagePatch(PAGE, { jobId: data.job_id, running: true });
      logFrontend(`任务已创建：${data.job_id}`);
      progressTracker.start({
        startedAtMs: data.created_at ? Number(data.created_at) * 1000 : Date.now(),
        expectedSeconds: data.expected_seconds || expected,
      });
      connectStream(data.job_id);
    } catch (err) {
      setRunning(false);
      GovUI.setBadge(statusBadge, "error");
      progressTracker.reset();
      runStatus.className = "run-status error";
      runStatus.textContent = err.message;
      logFrontend("启动失败：" + err.message);
    }
  });

  stopBtn.addEventListener("click", async function () {
    if (!state.running) return;
    stopBtn.disabled = true;
    try {
      await GovUI.stopJob({ jobId: state.jobId, logFrontend: logFrontend });
      runStatus.className = "run-status running";
      runStatus.textContent = "正在停止…";
    } catch (err) {
      logFrontend("停止失败：" + err.message);
      stopBtn.disabled = !state.running;
    }
  });

  resetBtn.addEventListener("click", function () {
    if (state.running) {
      logFrontend("任务运行中，请先停止后再重置页面。");
      return;
    }
    if (state.eventSource) {
      state.eventSource.close();
      state.eventSource = null;
    }
    GovUI.clearSession();
    state.jobId = null;
    state.uploadedFiles = [];
    state.excludePendingFile = null;
    hideImportPreview();
    excludePending.textContent = "";
    excludeUploadBtn.disabled = true;
    renderUploadList();
    GovUI.clearResultUi({
      frontendLog: frontendLog,
      scriptLog: scriptLog,
      summarySimple: summarySimple,
      summaryFull: summaryFull,
      completeBanner: completeBanner,
      resultPlaceholder: resultPlaceholder,
      dlResult: dlResult,
      dlPublish: dlPublish,
      dlStats: dlStats,
      dlSummarySimple: dlSummarySimple,
      dlSummaryFull: dlSummaryFull,
      dlAll: dlAll,
      runStatus: runStatus,
      statusBadge: statusBadge,
      progressTracker: progressTracker,
    });
    updateActionButtons();
    logFrontend("页面已重置。可重新上传文件后运行。");
  });

  // boot
  if (!life.reset) {
    GovUI.restoreFrontendLog(frontendLog, PAGE);
    const pageState = GovUI.loadPageState(PAGE);
    if (pageState.scriptLogText && scriptLog && !scriptLog.textContent) {
      scriptLog.textContent = pageState.scriptLogText;
    }
  } else {
    progressTracker.reset();
  }

  restoreSessionUploadedFiles();

  if (life.reset || !frontendLog.innerHTML.trim()) {
    logFrontend("页面已加载。请导入 Newdata 后点击「开始运行」。重叠日期以本次导入为准。");
    if (life.reset) logFrontend("检测到刷新，已重置本页会话状态。");
  }

  GovUI.setBadge(statusBadge, "idle");
  if (completeBanner) completeBanner.style.display = "none";
  if (resultSection) resultSection.style.display = "flex";

  loadExcludeYears().then(function () {
    return GovUI.restoreLatestJob({
      pageKind: PAGE,
      scriptLog: scriptLog,
      logFrontend: logFrontend,
      setProgressFn: function () {},
      onRunningChange: function (running, job) {
        state.running = running;
        if (job) state.jobId = job.job_id;
        updateActionButtons();
        GovUI.setBadge(statusBadge, running ? "running" : "idle");
        if (running) {
          runStatus.className = "run-status running";
          runStatus.innerHTML = '<span class="spinner"></span>脚本运行中...';
          GovUI.startProgressFromJob(progressTracker, job, GovUI.DEFAULT_EXPECTED);
        }
      },
      onDone: function (payload, job) {
        if (job) state.jobId = job.job_id;
        showResult(payload);
        updateActionButtons();
      },
      onRestored: function (job, streamOpts) {
        if (job) state.jobId = job.job_id;
        if (streamOpts) state.eventSource = streamOpts.eventSource;
        if (job && (job.status === "running" || job.status === "pending")) {
          GovUI.startProgressFromJob(progressTracker, job, job.expected_seconds);
        } else if (
          job &&
          (job.status === "success" || job.status === "failed" || job.status === "cancelled")
        ) {
          progressTracker.finish(job.duration_seconds);
        }
      },
    });
  });
})();
