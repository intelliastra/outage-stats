(function () {
  "use strict";

  var PAGE = "exclude";
  var life = GovUI.beginPageLifecycle();
  var uploading = false;

  var state = {
    uploadedFilename: null,
    uploadedSize: 0,
    jobId: null,
    running: false,
    eventSource: null,
  };

  var dropZone = document.getElementById("dropZone");
  var fileInput = document.getElementById("fileInput");
  var uploadList = document.getElementById("uploadList");
  var dataInfo = document.getElementById("dataInfo");
  var runBtn = document.getElementById("runBtn");
  var stopBtn = document.getElementById("stopBtn");
  var resetBtn = document.getElementById("resetBtn");
  var runStatus = document.getElementById("runStatus");
  var frontendLog = document.getElementById("frontendLog");
  var scriptLog = document.getElementById("scriptLog");
  var statusBadge = document.getElementById("statusBadge");
  var resultSection = document.getElementById("resultSection");
  var resultPlaceholder = document.getElementById("resultPlaceholder");
  var completeBanner = document.getElementById("completeBanner");
  var summarySimple = document.getElementById("summarySimple");
  var dlExcludeOutput = document.getElementById("dlExcludeOutput");

  var progressTracker = GovUI.createProgressTracker({
    progressBar: document.getElementById("progressBar"),
    elapsedText: document.getElementById("elapsedText"),
    etaText: document.getElementById("etaText"),
    pctText: document.getElementById("pctText"),
    totalTimeText: document.getElementById("totalTimeText"),
  });

  var logFrontend = GovUI.createLogger(frontendLog, PAGE);
  GovUI.bindTabs(document);
  if (completeBanner) completeBanner.style.display = "none";
  if (resultSection) resultSection.style.display = "flex";

  // Copy button
  document.querySelectorAll("[data-copy]").forEach(function (btn) {
    btn.addEventListener("click", function () {
      if (summarySimple && summarySimple.value) {
        GovUI.copyText(summarySimple.value);
        logFrontend("已复制摘要到剪贴板。");
      }
    });
  });

  function renderUploadList() {
    uploadList.innerHTML = "";
    if (state.uploadedFilename) {
      var li = document.createElement("li");
      var sizeKb = (state.uploadedSize / 1024).toFixed(1);
      li.innerHTML =
        GovUI.escapeHtml(state.uploadedFilename) +
        '<span class="size">(' + sizeKb + " KB)</span>";
      uploadList.appendChild(li);
    }
    updateActionButtons();
  }

  function updateActionButtons() {
    runBtn.disabled = state.running || !state.uploadedFilename;
    stopBtn.disabled = !state.running;
    resetBtn.disabled = state.running;
  }

  function setRunning(running) {
    state.running = running;
    updateActionButtons();
    GovUI.setBadge(statusBadge, running ? "running" : "idle");
    if (running) {
      runStatus.className = "run-status running";
      runStatus.innerHTML = '<span class="spinner"></span>脚本运行中…';
    }
    persistForm();
  }

  function persistForm() {
    GovUI.savePagePatch(PAGE, {
      uploadedFilename: state.uploadedFilename,
      uploadedSize: state.uploadedSize,
      jobId: state.jobId,
      running: state.running,
    });
  }

  function renderDataInfo(config) {
    if (!config) return;
    var scriptOk = config.exclude_script_exists;
    var fileCount = config.count || 0;
    var html =
      '<strong>剔除脚本</strong>：' + GovUI.escapeHtml(config.exclude_script || "—") +
      (scriptOk ? ' <span style="color:var(--gov-success)">（就绪）</span>' : ' <span style="color:var(--gov-error)">（未找到）</span>') +
      "<br>Newdata 目录文件数：" + fileCount;
    dataInfo.innerHTML = html;
  }

  function showResult(payload) {
    state.running = false;
    updateActionButtons();
    if (payload.kind && payload.kind !== "exclude") return;
    GovUI.setBadge(statusBadge, GovUI.badgeFromStatus(payload.status));
    if (payload.status === "success" || payload.status === "failed" || payload.status === "cancelled") {
      progressTracker.finish(payload.duration_seconds);
    } else {
      progressTracker.reset();
    }
    if (resultSection) resultSection.style.display = "flex";

    if (completeBanner) {
      var status = payload.status || "";
      var ok = status === "success";
      var cancelled = status === "cancelled";
      if (!status || status === "running" || status === "pending") {
        completeBanner.style.display = "none";
      } else {
        completeBanner.style.display = "";
        completeBanner.className =
          "complete-banner " + (ok ? "success" : cancelled ? "cancelled" : "error");
        var durationNote =
          payload.duration_seconds != null
            ? " 总用时 " + GovUI.formatDuration(payload.duration_seconds) + "。"
            : "";
        if (ok) completeBanner.textContent = "运行完成！摘要与文件已就绪。" + durationNote;
        else if (cancelled) completeBanner.textContent = "任务已停止，后台脚本进程已终止。" + durationNote;
        else completeBanner.textContent = "运行失败（退出码 " + (payload.exit_code ?? "?") + "），请查看脚本日志。" + durationNote;
      }
    }

    if (runStatus) {
      if (payload.status === "success") {
        runStatus.className = "run-status success";
        runStatus.textContent = "运行完成";
      } else if (payload.status === "cancelled") {
        runStatus.className = "run-status error";
        runStatus.textContent = "已停止";
      } else if (payload.status === "failed") {
        runStatus.className = "run-status error";
        runStatus.textContent = "运行失败";
      }
    }

    if (resultPlaceholder) {
      resultPlaceholder.style.display =
        payload.status === "success" || payload.status === "failed" || payload.status === "cancelled" ? "none" : "";
    }

    if (summarySimple) {
      summarySimple.value = GovUI.trimSummaryText(payload.summary_simple || "");
    }

    GovUI.setDownloadLink(dlExcludeOutput, state.jobId, "exclude_output", payload.files);

    GovUI.savePagePatch(PAGE, {
      jobId: state.jobId,
      lastResult: Object.assign({}, payload, { kind: "exclude" }),
      finished: true,
    });
  }

  async function loadConfig() {
    try {
      var resp = await GovUI.fetchWithTimeout("/api/exclude/config", null, 8000);
      var data = await GovUI.parseJsonSafe(resp);
      if (!resp.ok) throw new Error(data.detail || "加载配置失败");
      renderDataInfo(data);
      logFrontend("已加载剔除外力破坏配置。");
      return data;
    } catch (err) {
      if (!GovUI.isGatewayError(err)) {
        dataInfo.textContent = "配置加载失败：" + err.message;
        logFrontend("配置加载失败：" + err.message);
      }
      return null;
    }
  }

  async function uploadFile(file) {
    var ext = file.name.split(".").pop().toLowerCase();
    if (ext !== "xlsx" && ext !== "xls") {
      logFrontend("跳过非 Excel 文件：" + file.name);
      return false;
    }
    logFrontend("正在上传：" + file.name + " ...");
    try {
      var data = await GovUI.postFileWithRetry("/api/upload", file, { log: logFrontend });
      state.uploadedFilename = data.filename;
      state.uploadedSize = data.size;
      renderUploadList();
      logFrontend("导入成功：" + data.filename + " (" + (data.size / 1024).toFixed(1) + " KB)");
      persistForm();
      return true;
    } catch (err) {
      logFrontend("上传失败：" + file.name + " — " + err.message);
      return false;
    }
  }

  async function handleFiles(fileList) {
    var files = Array.from(fileList || []);
    var excelFiles = files.filter(function (f) {
      var ext = f.name.split(".").pop().toLowerCase();
      return ext === "xlsx" || ext === "xls";
    });
    if (excelFiles.length === 0) {
      logFrontend("未选择有效的 Excel 文件（.xlsx / .xls）。");
      return;
    }
    if (excelFiles.length > 1) {
      logFrontend("一次仅保留最后一个文件，已忽略前 " + (excelFiles.length - 1) + " 个。");
    }
    var last = excelFiles[excelFiles.length - 1];
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

  // Drag & drop
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

  // Run
  runBtn.addEventListener("click", async function () {
    if (state.running || !state.uploadedFilename) return;
    if (completeBanner) {
      completeBanner.style.display = "none";
      completeBanner.textContent = "";
    }
    if (resultPlaceholder) resultPlaceholder.style.display = "";
    summarySimple.value = "";
    GovUI.setDownloadLink(dlExcludeOutput, null, "exclude_output", null);
    scriptLog.textContent = "";
    progressTracker.reset();
    setRunning(true);
    GovUI.setBadge(statusBadge, "running");
    logFrontend("正在启动剔除外力破坏统计：" + state.uploadedFilename + " …");

    try {
      var expected = await GovUI.fetchExpectedSeconds("exclude");
      var resp = await fetch("/api/exclude/run", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ input_filename: state.uploadedFilename }),
      });
      var data = await resp.json();
      if (!resp.ok) throw new Error(data.detail || "启动失败");
      state.jobId = data.job_id;
      GovUI.saveSession({
        activeJobId: data.job_id,
        jobId: data.job_id,
        kind: "exclude",
        status: "running",
        finished: false,
      });
      persistForm();
      logFrontend("任务已创建：" + data.job_id);
      progressTracker.start({
        startedAtMs: data.created_at ? Number(data.created_at) * 1000 : Date.now(),
        expectedSeconds: data.expected_seconds || expected,
      });

      var streamOpts = {
        jobId: data.job_id,
        scriptLog: scriptLog,
        logFrontend: logFrontend,
        running: true,
        eventSource: null,
        kind: "exclude",
        pageKind: PAGE,
        onRunningChange: function (v) {
          state.running = v;
          streamOpts.running = v;
          updateActionButtons();
          if (v) GovUI.setBadge(statusBadge, "running");
        },
        setProgressFn: function () {},
        onDone: function (payload) {
          showResult(payload);
        },
      };
      GovUI.connectStream(streamOpts);
      state.eventSource = streamOpts.eventSource;
    } catch (err) {
      setRunning(false);
      GovUI.setBadge(statusBadge, "error");
      progressTracker.reset();
      runStatus.className = "run-status error";
      runStatus.textContent = err.message;
      logFrontend("启动失败：" + err.message);
    }
  });

  // Stop
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

  // Reset
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
    state.uploadedFilename = null;
    state.uploadedSize = 0;
    renderUploadList();
    GovUI.clearResultUi({
      frontendLog: frontendLog,
      scriptLog: scriptLog,
      summarySimple: summarySimple,
      summaryFull: null,
      completeBanner: completeBanner,
      resultPlaceholder: resultPlaceholder,
      dlResult: dlExcludeOutput,
      dlPublish: null,
      dlStats: null,
      dlSummarySimple: null,
      dlSummaryFull: null,
      dlAll: null,
      runStatus: runStatus,
      statusBadge: statusBadge,
      progressTracker: progressTracker,
    });
    updateActionButtons();
    logFrontend("页面已重置。可重新上传文件后开始运行。");
  });

  // Init
  if (life.reset || !frontendLog.innerHTML.trim()) {
    logFrontend("剔除外力破坏页面已加载。");
    if (life.reset) logFrontend("检测到刷新，已重置本页会话状态。");
  }
  GovUI.setBadge(statusBadge, "idle");

  // Restore form state
  if (!life.reset) {
    var ps = GovUI.loadPageState(PAGE);
    if (ps.uploadedFilename) {
      state.uploadedFilename = ps.uploadedFilename;
      state.uploadedSize = ps.uploadedSize || 0;
      renderUploadList();
    }
    GovUI.restoreFrontendLog(frontendLog, PAGE);
    if (ps.scriptLogText && scriptLog) scriptLog.textContent = ps.scriptLogText;
  }

  loadConfig().then(function () {
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
          runStatus.innerHTML = '<span class="spinner"></span>脚本运行中…';
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
