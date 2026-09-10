(function () {
  "use strict";

  const PAGE = "historical";
  const life = GovUI.beginPageLifecycle();

  const state = {
    config: null,
    jobId: null,
    running: false,
    eventSource: null,
  };

  const startMonth = document.getElementById("startMonth");
  const startDay = document.getElementById("startDay");
  const endMonth = document.getElementById("endMonth");
  const endDay = document.getElementById("endDay");
  const dataInfo = document.getElementById("dataInfo");
  const runBtn = document.getElementById("runBtn");
  const stopBtn = document.getElementById("stopBtn");
  const resetBtn = document.getElementById("resetBtn");
  const runStatus = document.getElementById("runStatus");
  const frontendLog = document.getElementById("frontendLog");
  const scriptLog = document.getElementById("scriptLog");
  const statusBadge = document.getElementById("statusBadge");
  const resultSection = document.getElementById("resultSection");
  const resultPlaceholder = document.getElementById("resultPlaceholder");
  const completeBanner = document.getElementById("completeBanner");
  const summarySimple = document.getElementById("summarySimple");
  const summaryFull = document.getElementById("summaryFull");
  const dlResult = document.getElementById("dlResult");
  const dlPublish = document.getElementById("dlPublish");
  const dlStats = document.getElementById("dlStats");
  const dlSummarySimple = document.getElementById("dlSummarySimple");
  const dlSummaryFull = document.getElementById("dlSummaryFull");
  const dlAll = document.getElementById("dlAll");

  const progressTracker = GovUI.createProgressTracker({
    progressBar: document.getElementById("progressBar"),
    elapsedText: document.getElementById("elapsedText"),
    etaText: document.getElementById("etaText"),
    pctText: document.getElementById("pctText"),
    totalTimeText: document.getElementById("totalTimeText"),
  });

  const logFrontend = GovUI.createLogger(frontendLog, PAGE);
  GovUI.bindTabs(document);
  GovUI.bindCopyButtons(summarySimple, summaryFull, logFrontend);
  if (completeBanner) completeBanner.style.display = "none";
  if (resultSection) resultSection.style.display = "flex";

  function fillSelect(el, from, to) {
    el.innerHTML = "";
    for (let i = from; i <= to; i++) {
      const opt = document.createElement("option");
      opt.value = String(i);
      opt.textContent = String(i);
      el.appendChild(opt);
    }
  }

  function daysInMonth(year, month) {
    return new Date(year, month, 0).getDate();
  }

  function selectedYear() {
    const el = document.querySelector('input[name="dataYear"]:checked');
    return el ? parseInt(el.value, 10) : 2025;
  }

  function persistForm() {
    GovUI.savePagePatch(PAGE, {
      year: selectedYear(),
      startMonth: startMonth.value,
      startDay: startDay.value,
      endMonth: endMonth.value,
      endDay: endDay.value,
      jobId: state.jobId,
      running: state.running,
    });
  }

  function refreshDayOptions() {
    const year = selectedYear();
    const sm = parseInt(startMonth.value, 10);
    const em = parseInt(endMonth.value, 10);
    const prevSd = parseInt(startDay.value, 10) || 1;
    const prevEd = parseInt(endDay.value, 10) || 1;
    fillSelect(startDay, 1, daysInMonth(year, sm));
    const endYearGuess = em < sm || (em === sm && prevEd < prevSd) ? year + 1 : year;
    fillSelect(endDay, 1, daysInMonth(endYearGuess, em));
    startDay.value = String(Math.min(prevSd, parseInt(startDay.lastChild.value, 10)));
    endDay.value = String(Math.min(prevEd, parseInt(endDay.lastChild.value, 10)));
    persistForm();
  }

  function buildPeriodPreview() {
    const year = selectedYear();
    const sm = parseInt(startMonth.value, 10);
    const sd = parseInt(startDay.value, 10);
    const em = parseInt(endMonth.value, 10);
    const ed = parseInt(endDay.value, 10);
    let endYear = year;
    if (em < sm || (em === sm && ed < sd)) endYear = year + 1;
    return {
      year: year,
      start_month: sm,
      start_day: sd,
      end_month: em,
      end_day: ed,
      label: `${year}-${sm}-${sd} ~ ${endYear}-${em}-${ed}`,
    };
  }

  function renderDataInfo() {
    if (!state.config) return;
    const year = String(selectedYear());
    const info = state.config.year_info[year] || {};
    const ready = !!info.data_ready && !!info.exclude_ready;
    dataInfo.innerHTML =
      `<strong>数据源</strong>：${GovUI.escapeHtml(info.source_label || "—")}<br>` +
      `文件：${GovUI.escapeHtml(info.data_filename || (info.data_ready ? "可自动准备" : "未就绪"))}<br>` +
      `剔除文件：${GovUI.escapeHtml(info.exclude_filename || "未找到表11")}<br>` +
      `<span style="color:${ready ? "var(--gov-success)" : "var(--gov-error)"}">${
        ready ? "就绪，可获取统计" : "数据或剔除文件未就绪，无法运行"
      }</span>`;
    updateActionButtons();
  }

  function updateActionButtons() {
    if (!state.config) {
      runBtn.disabled = true;
      stopBtn.disabled = !state.running;
      resetBtn.disabled = state.running;
      return;
    }
    const year = String(selectedYear());
    const info = state.config.year_info[year] || {};
    runBtn.disabled = state.running || !info.data_ready || !info.exclude_ready;
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

  function showResult(payload) {
    state.running = false;
    updateActionButtons();
    // 只展示本页 kind 的结果
    if (payload.kind && payload.kind !== "historical") {
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
    GovUI.savePagePatch(PAGE, {
      jobId: state.jobId,
      lastResult: Object.assign({}, payload, { kind: "historical" }),
      finished: true,
    });
  }

  async function loadConfig() {
    try {
      const resp = await GovUI.fetchWithTimeout("/api/historical/config", null, 8000);
      state.config = await GovUI.parseJsonSafe(resp);
      renderDataInfo();
      logFrontend("已加载历史统计配置。");
      if (state.config.hint) logFrontend(state.config.hint);
    } catch (err) {
      if (!GovUI.isGatewayError(err)) {
        dataInfo.textContent = "配置加载失败：" + err.message;
        logFrontend("配置加载失败：" + err.message);
      }
    }
  }

  fillSelect(startMonth, 1, 12);
  fillSelect(endMonth, 1, 12);
  startMonth.value = "1";
  endMonth.value = "12";
  refreshDayOptions();
  startDay.value = "1";
  endDay.value = "31";

  // restore form from session (page switch)
  if (!life.reset) {
    const ps = GovUI.loadPageState(PAGE);
    if (ps.year) {
      const radio = document.querySelector('input[name="dataYear"][value="' + ps.year + '"]');
      if (radio) radio.checked = true;
    }
    if (ps.startMonth) startMonth.value = String(ps.startMonth);
    if (ps.endMonth) endMonth.value = String(ps.endMonth);
    refreshDayOptions();
    if (ps.startDay) startDay.value = String(ps.startDay);
    if (ps.endDay) endDay.value = String(ps.endDay);
    GovUI.restoreFrontendLog(frontendLog, PAGE);
    if (ps.scriptLogText && scriptLog) scriptLog.textContent = ps.scriptLogText;
  }

  startMonth.addEventListener("change", refreshDayOptions);
  endMonth.addEventListener("change", refreshDayOptions);
  startDay.addEventListener("change", persistForm);
  endDay.addEventListener("change", persistForm);
  document.querySelectorAll('input[name="dataYear"]').forEach(function (radio) {
    radio.addEventListener("change", function () {
      refreshDayOptions();
      renderDataInfo();
      persistForm();
    });
  });

  runBtn.addEventListener("click", async function () {
    if (state.running) return;
    const period = buildPeriodPreview();
    if (completeBanner) {
      completeBanner.style.display = "none";
      completeBanner.textContent = "";
    }
    if (resultPlaceholder) resultPlaceholder.style.display = "";
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
    GovUI.setBadge(statusBadge, "running");
    logFrontend(`正在启动历史统计：${period.label}（${period.year}）…`);

    try {
      const expected = await GovUI.fetchExpectedSeconds("historical");
      const resp = await fetch("/api/historical/run", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          year: period.year,
          start_month: period.start_month,
          start_day: period.start_day,
          end_month: period.end_month,
          end_day: period.end_day,
        }),
      });
      const data = await resp.json();
      if (!resp.ok) throw new Error(data.detail || "启动失败");
      state.jobId = data.job_id;
      GovUI.saveSession({
        activeJobId: data.job_id,
        jobId: data.job_id,
        kind: "historical",
        status: "running",
        finished: false,
      });
      persistForm();
      logFrontend(
        `任务已创建：${data.job_id}，区间 ${data.period_start} ~ ${data.period_end}`
      );
      progressTracker.start({
        startedAtMs: data.created_at ? Number(data.created_at) * 1000 : Date.now(),
        expectedSeconds: data.expected_seconds || expected,
      });

      const streamOpts = {
        jobId: data.job_id,
        scriptLog: scriptLog,
        logFrontend: logFrontend,
        running: true,
        eventSource: null,
        kind: "historical",
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
    document.querySelector('input[name="dataYear"][value="2025"]').checked = true;
    startMonth.value = "1";
    endMonth.value = "12";
    refreshDayOptions();
    startDay.value = "1";
    endDay.value = "31";
    renderDataInfo();
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
    logFrontend("页面已重置。可重新选择参数后获取统计。");
  });

  if (life.reset || !frontendLog.innerHTML.trim()) {
    logFrontend("历史区间统计页面已加载。");
    if (life.reset) logFrontend("检测到刷新，已重置本页会话状态。");
  }
  GovUI.setBadge(statusBadge, "idle");

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
