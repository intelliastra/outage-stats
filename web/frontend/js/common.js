/** Shared helpers for daily + historical pages. */
(function (global) {
  "use strict";

  var SESSION_KEY = "outage_stats_session_v2";
  var DEFAULT_EXPECTED = 1800;

  function now() {
    return new Date().toLocaleTimeString("zh-CN", { hour12: false });
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function formatDuration(seconds) {
    var s = Math.max(0, Math.floor(Number(seconds) || 0));
    var h = Math.floor(s / 3600);
    var m = Math.floor((s % 3600) / 60);
    var sec = s % 60;
    if (h > 0) return h + "小时" + m + "分" + sec + "秒";
    if (m > 0) return m + "分" + sec + "秒";
    return sec + "秒";
  }

  function isReloadNavigation() {
    try {
      var nav = performance.getEntriesByType("navigation")[0];
      if (nav && nav.type === "reload") return true;
      // legacy
      if (performance.navigation && performance.navigation.type === 1) return true;
    } catch (_) {
      /* ignore */
    }
    return false;
  }

  function loadSession() {
    try {
      return JSON.parse(sessionStorage.getItem(SESSION_KEY) || "{}") || {};
    } catch (_) {
      return {};
    }
  }

  function saveSession(partial) {
    try {
      var next = Object.assign({}, loadSession(), partial || {});
      sessionStorage.setItem(SESSION_KEY, JSON.stringify(next));
    } catch (_) {
      /* ignore */
    }
  }

  function clearSession() {
    try {
      sessionStorage.removeItem(SESSION_KEY);
    } catch (_) {
      /* ignore */
    }
  }

  /** Refresh/re-login should reset; in-tab page switch keeps state. */
  function beginPageLifecycle() {
    if (isReloadNavigation()) {
      clearSession();
      return { reset: true };
    }
    return { reset: false, session: loadSession() };
  }

  function savePagePatch(pageKind, patch) {
    var all = loadSession();
    var page = Object.assign({}, all[pageKind] || {}, patch || {});
    all[pageKind] = page;
    try {
      sessionStorage.setItem(SESSION_KEY, JSON.stringify(all));
    } catch (_) {
      /* ignore */
    }
  }

  function loadPageState(pageKind) {
    var all = loadSession();
    return all[pageKind] || {};
  }

  function createLogger(frontendLogEl, pageKind) {
    return function logFrontend(message) {
      if (!frontendLogEl) return;
      var line = document.createElement("div");
      line.className = "log-line";
      line.innerHTML = '<span class="ts">[' + now() + "]</span>" + escapeHtml(message);
      frontendLogEl.appendChild(line);
      frontendLogEl.scrollTop = frontendLogEl.scrollHeight;
      if (pageKind) {
        savePagePatch(pageKind, { frontendLogHtml: frontendLogEl.innerHTML });
      }
    };
  }

  function restoreFrontendLog(frontendLogEl, pageKind) {
    if (!frontendLogEl) return;
    var page = loadPageState(pageKind);
    if (page.frontendLogHtml) {
      frontendLogEl.innerHTML = page.frontendLogHtml;
      frontendLogEl.scrollTop = frontendLogEl.scrollHeight;
    }
  }

  function bindTabs(root) {
    var scope = root || document;
    scope.querySelectorAll(".tab").forEach(function (tab) {
      tab.addEventListener("click", function () {
        var parent = tab.closest(".gov-panel-monitor, .gov-card") || scope;
        parent.querySelectorAll(".tab").forEach(function (t) {
          t.classList.remove("active");
        });
        parent.querySelectorAll(".log-panel").forEach(function (p) {
          p.classList.remove("active");
        });
        tab.classList.add("active");
        var targetId = tab.dataset.tab === "frontend" ? "frontendLog" : "scriptLog";
        var panel = parent.querySelector("#" + targetId);
        if (panel) panel.classList.add("active");
      });
    });
  }

  function setBadge(badgeEl, status) {
    if (!badgeEl) return;
    badgeEl.className = "gov-badge " + (status || "idle");
    var labels = {
      idle: "等待",
      running: "运行中",
      success: "成功",
      error: "失败",
      cancelled: "已停止",
    };
    var spinner = status === "running" ? '<span class="spinner"></span> ' : "";
    badgeEl.innerHTML = spinner + (labels[status] || "等待");
  }

  function setProgress(progressEl, active) {
    if (!progressEl) return;
    progressEl.classList.toggle("active", !!active);
  }

  function setDownloadLink(el, jobId, key, files) {
    if (!el) return;
    var enabled = false;
    if (key === "all") {
      enabled = !!(jobId && files && (files.result || files.publish || files.stats || files.summary_simple || files.summary_full));
      if (enabled) el.href = "/api/download/" + jobId + "/all";
    } else {
      var path = files && files[key];
      enabled = !!(jobId && path);
      if (enabled) el.href = "/api/download/" + jobId + "/" + key;
    }
    if (enabled) {
      el.removeAttribute("aria-disabled");
      el.classList.remove("is-disabled");
    } else {
      el.href = "#";
      el.setAttribute("aria-disabled", "true");
      el.classList.add("is-disabled");
    }
  }

  function createProgressTracker(els) {
    var timer = null;
    var startedAt = null;
    var expected = DEFAULT_EXPECTED;
    var finished = false;

    function ensureEls() {
      return {
        progressBar: els.progressBar || document.getElementById("progressBar"),
        elapsedText: els.elapsedText || document.getElementById("elapsedText"),
        etaText: els.etaText || document.getElementById("etaText"),
        pctText: els.pctText || document.getElementById("pctText"),
        totalTimeText: els.totalTimeText || document.getElementById("totalTimeText"),
      };
    }

    function paint(elapsed, pct, remaining, doneSeconds) {
      var e = ensureEls();
      if (e.progressBar) e.progressBar.style.width = Math.max(0, Math.min(100, pct)).toFixed(1) + "%";
      if (e.pctText) e.pctText.textContent = Math.floor(pct) + "%";
      if (e.elapsedText) e.elapsedText.textContent = "已用：" + formatDuration(elapsed);
      if (e.etaText) {
        if (remaining == null) e.etaText.textContent = "预计剩余：—";
        else if (remaining <= 0) e.etaText.textContent = "预计剩余：即将完成";
        else e.etaText.textContent = "预计剩余：" + formatDuration(remaining);
      }
      if (e.totalTimeText) {
        e.totalTimeText.textContent =
          doneSeconds != null ? "本次运行总用时：" + formatDuration(doneSeconds) : "";
      }
    }

    function tick() {
      if (!startedAt || finished) return;
      var elapsed = (Date.now() - startedAt) / 1000;
      var pct = expected > 0 ? (elapsed / expected) * 100 : 0;
      if (pct >= 99) pct = 99;
      paint(elapsed, pct, Math.max(0, expected - elapsed), null);
    }

    return {
      start: function (opts) {
        opts = opts || {};
        finished = false;
        startedAt = opts.startedAtMs || Date.now();
        expected = Number(opts.expectedSeconds) > 0 ? Number(opts.expectedSeconds) : DEFAULT_EXPECTED;
        var e = ensureEls();
        if (e.totalTimeText) e.totalTimeText.textContent = "";
        tick();
        if (timer) clearInterval(timer);
        timer = setInterval(tick, 1000);
      },
      finish: function (durationSeconds) {
        finished = true;
        if (timer) {
          clearInterval(timer);
          timer = null;
        }
        var secs =
          durationSeconds != null
            ? Number(durationSeconds)
            : startedAt
              ? (Date.now() - startedAt) / 1000
              : 0;
        paint(secs, 100, null, secs);
      },
      reset: function () {
        finished = true;
        if (timer) {
          clearInterval(timer);
          timer = null;
        }
        startedAt = null;
        paint(0, 0, null, null);
        var e = ensureEls();
        if (e.elapsedText) e.elapsedText.textContent = "已用：—";
        if (e.etaText) e.etaText.textContent = "预计剩余：—";
        if (e.pctText) e.pctText.textContent = "0%";
        if (e.totalTimeText) e.totalTimeText.textContent = "";
        if (e.progressBar) e.progressBar.style.width = "0%";
      },
    };
  }

  async function fetchExpectedSeconds(kind) {
    try {
      var resp = await fetch("/api/timing");
      var data = await resp.json();
      var val = data && data[kind];
      if (val && Number(val) > 0) return Number(val);
      if (data && data.default_expected_seconds) return Number(data.default_expected_seconds);
    } catch (_) {
      /* ignore */
    }
    return DEFAULT_EXPECTED;
  }

  function trimSummaryText(text) {
    var value = String(text || "");
    var cut = value.search(
      /\[OK\]\s*步骤\s*7[：:]|>>\s*步骤\s*8|步骤\s*8[：:]|\[OK\]\s*全部步骤执行完毕/
    );
    if (cut >= 0) value = value.slice(0, cut);
    return value.replace(/\s+$/g, "");
  }

  function showResult(opts) {
    var resultSection = opts.resultSection;
    var completeBanner = opts.completeBanner;
    var resultPlaceholder = opts.resultPlaceholder;
    var summarySimple = opts.summarySimple;
    var summaryFull = opts.summaryFull;
    var dlResult = opts.dlResult;
    var dlPublish = opts.dlPublish;
    var dlStats = opts.dlStats;
    var dlSummarySimple = opts.dlSummarySimple;
    var dlSummaryFull = opts.dlSummaryFull;
    var dlAll = opts.dlAll;
    var jobId = opts.jobId;
    var payload = opts.payload || {};
    var runStatus = opts.runStatus;

    if (!resultSection) return;
    resultSection.classList.remove("hidden");
    resultSection.style.display = "flex";

    var status = payload.status || "";
    var ok = status === "success";
    var cancelled = status === "cancelled";

    if (resultPlaceholder) {
      resultPlaceholder.style.display =
        ok || cancelled || status === "failed" ? "none" : "";
    }

    if (completeBanner) {
      if (!status || status === "running" || status === "pending") {
        completeBanner.className = "complete-banner";
        completeBanner.textContent = "";
        completeBanner.style.display = "none";
      } else {
        completeBanner.style.display = "";
        completeBanner.className =
          "complete-banner " + (ok ? "success" : cancelled ? "cancelled" : "error");
        var durationNote =
          payload.duration_seconds != null
            ? " 总用时 " + formatDuration(payload.duration_seconds) + "。"
            : "";
        if (ok) completeBanner.textContent = "运行完成！摘要与文件已就绪。" + durationNote;
        else if (cancelled)
          completeBanner.textContent = "任务已停止，后台脚本进程已终止。" + durationNote;
        else
          completeBanner.textContent =
            "运行失败（退出码 " +
            (payload.exit_code ?? "?") +
            "），请查看脚本日志。" +
            durationNote;
      }
    }

    if (runStatus) {
      if (ok) {
        runStatus.className = "run-status success";
        runStatus.textContent = "运行完成";
      } else if (cancelled) {
        runStatus.className = "run-status error";
        runStatus.textContent = "已停止";
      } else if (status === "failed") {
        runStatus.className = "run-status error";
        runStatus.textContent = "运行失败";
      }
    }

    if (summarySimple) summarySimple.value = trimSummaryText(payload.summary_simple || "");
    if (summaryFull) summaryFull.value = trimSummaryText(payload.summary_full || "");

    setDownloadLink(dlResult, jobId, "result", payload.files);
    setDownloadLink(dlPublish, jobId, "publish", payload.files);
    setDownloadLink(dlStats, jobId, "stats", payload.files);
    setDownloadLink(dlSummarySimple, jobId, "summary_simple", payload.files);
    setDownloadLink(dlSummaryFull, jobId, "summary_full", payload.files);
    setDownloadLink(dlAll, jobId, "all", payload.files);
  }

  function connectStream(opts) {
    var jobId = opts.jobId;
    var scriptLog = opts.scriptLog;
    var logFrontend = opts.logFrontend;
    var onDone = opts.onDone;
    var onRunningChange = opts.onRunningChange;
    var setProgressFn = opts.setProgressFn;
    var clearScript = opts.clearScript !== false;
    var pageKind = opts.pageKind;
    var finished = false;
    var pollTimer = null;

    if (opts.eventSource) opts.eventSource.close();
    if (clearScript && scriptLog) scriptLog.textContent = "";
    if (logFrontend) logFrontend("已连接 SSE 日志流，任务 ID：" + jobId);

    function stopPoll() {
      if (pollTimer) {
        clearInterval(pollTimer);
        pollTimer = null;
      }
    }

    function finishOnce(payload) {
      if (finished) return;
      finished = true;
      opts.running = false;
      stopPoll();
      if (opts.eventSource) {
        try {
          opts.eventSource.close();
        } catch (_) {
          /* ignore */
        }
        opts.eventSource = null;
      }
      if (onRunningChange) onRunningChange(false);
      if (setProgressFn) setProgressFn(false);
      saveSession({
        activeJobId: jobId,
        jobId: jobId,
        kind: (payload && payload.kind) || opts.kind || null,
        status: payload && payload.status,
        finished: true,
      });
      if (pageKind) {
        savePagePatch(pageKind, {
          jobId: jobId,
          status: payload && payload.status,
          finished: true,
          lastResult: payload || null,
        });
      }
      if (onDone) onDone(payload || {});
    }

    var es = new EventSource("/api/jobs/" + jobId + "/stream");
    opts.eventSource = es;
    opts.running = true;

    es.onmessage = function (event) {
      try {
        var data = JSON.parse(event.data);
        if (data.line && scriptLog) {
          scriptLog.textContent += data.line;
          scriptLog.scrollTop = scriptLog.scrollHeight;
          if (pageKind) savePagePatch(pageKind, { scriptLogText: scriptLog.textContent });
          saveSession({ lastScriptLog: scriptLog.textContent, activeJobId: jobId });
          // 日志出现全部完成标记时，主动拉取最终状态，避免 SSE done 丢失
          if (
            !finished &&
            (/全部步骤执行完毕/.test(data.line) || /\[WEB\].*脚本进程已终止/.test(data.line))
          ) {
            pollJobOnce();
          }
        }
      } catch (_) {
        if (scriptLog) {
          scriptLog.textContent += event.data;
          scriptLog.scrollTop = scriptLog.scrollHeight;
        }
      }
    };

    es.addEventListener("done", function (event) {
      try {
        var payload = JSON.parse(event.data);
        finishOnce(payload);
        if (logFrontend) {
          if (payload.status === "success") {
            logFrontend(
              "脚本执行成功。" +
                (payload.duration_seconds != null
                  ? " 总用时 " + formatDuration(payload.duration_seconds) + "。"
                  : "")
            );
          } else if (payload.status === "cancelled") {
            logFrontend("脚本已被停止。");
          } else {
            logFrontend("脚本执行失败，退出码：" + payload.exit_code);
          }
        }
      } catch (err) {
        if (logFrontend) logFrontend("解析完成事件失败：" + err.message);
        pollJobOnce();
      }
    });

    es.addEventListener("error", function () {
      if (!finished && opts.running) {
        if (logFrontend) logFrontend("SSE 连接中断，尝试轮询任务状态...");
        try {
          es.close();
        } catch (_) {
          /* ignore */
        }
        opts.eventSource = null;
        pollJobOnce();
        if (!pollTimer) {
          pollTimer = setInterval(pollJobOnce, 2000);
        }
      }
    });

    function pollJobOnce() {
      if (finished) {
        stopPoll();
        return;
      }
      fetch("/api/jobs/" + jobId + "?include_logs=true")
        .then(parseJsonSafe)
        .then(function (data) {
          if (
            data.status === "success" ||
            data.status === "failed" ||
            data.status === "cancelled"
          ) {
            if (opts.scriptLog && data.logs != null) {
              opts.scriptLog.textContent = data.logs;
            }
            finishOnce({
              status: data.status,
              exit_code: data.exit_code,
              summary_simple: data.summary_simple,
              summary_full: data.summary_full,
              files: data.files,
              kind: data.kind,
              created_at: data.created_at,
              finished_at: data.finished_at,
              duration_seconds: data.duration_seconds,
              expected_seconds: data.expected_seconds,
            });
          }
        })
        .catch(function (err) {
          if (logFrontend && !finished) logFrontend("轮询失败：" + err.message);
        });
    }

    // 运行中兜底轮询，防止 done 事件未到达导致进度条卡住
    pollTimer = setInterval(pollJobOnce, 3000);

    return es;
  }

  async function pollJob(jobId, opts) {
    try {
      var resp = await fetch("/api/jobs/" + jobId + "?include_logs=true");
      var data = await parseJsonSafe(resp);
      if (
        data.status === "success" ||
        data.status === "failed" ||
        data.status === "cancelled"
      ) {
        if (opts.scriptLog && data.logs != null) opts.scriptLog.textContent = data.logs;
        if (opts.onRunningChange) opts.onRunningChange(false);
        if (opts.setProgressFn) opts.setProgressFn(false);
        saveSession({ jobId: jobId, status: data.status, finished: true });
        if (opts.onDone) {
          opts.onDone({
            status: data.status,
            exit_code: data.exit_code,
            summary_simple: data.summary_simple,
            summary_full: data.summary_full,
            files: data.files,
            kind: data.kind,
            created_at: data.created_at,
            finished_at: data.finished_at,
            duration_seconds: data.duration_seconds,
            expected_seconds: data.expected_seconds,
          });
        }
      } else {
        setTimeout(function () {
          pollJob(jobId, opts);
        }, 2000);
      }
    } catch (err) {
      if (opts.logFrontend) opts.logFrontend("轮询失败：" + err.message);
      if (opts.onRunningChange) opts.onRunningChange(false);
    }
  }

  async function stopJob(opts) {
    var jobId = opts.jobId;
    var logFrontend = opts.logFrontend;
    var url = jobId ? "/api/jobs/" + jobId + "/stop" : "/api/stop";
    if (logFrontend) logFrontend("正在请求停止后台任务…");
    var resp = await fetch(url, { method: "POST" });
    var data = await resp.json().catch(function () {
      return {};
    });
    if (!resp.ok) throw new Error(data.detail || "停止失败");
    if (logFrontend) logFrontend("已发送停止请求，等待进程退出…");
    return data;
  }

  async function restoreLatestJob(opts) {
    var logFrontend = opts.logFrontend;
    var pageKind = opts.pageKind || "daily";
    var expectedKind = pageKind === "historical" ? "historical" : "daily";
    try {
      var resp = await fetchWithTimeout(
        "/api/jobs/latest?include_logs=true&kind=" + encodeURIComponent(expectedKind),
        null,
        8000
      );
      var data = await parseJsonSafe(resp);
      var job = data.job;

      // 另一类任务在跑：只提示忙碌，不把对方输出灌进本页
      if (data.busy && data.running_kind && data.running_kind !== expectedKind) {
        if (logFrontend) {
          logFrontend(
            "系统正忙：另一页面的" +
              (data.running_kind === "historical" ? "历史统计" : "日常统计") +
              "任务运行中，本页输出保持独立。"
          );
        }
        if (opts.onOtherBusy) opts.onOtherBusy(data.running_kind);
      }

      if (!job) {
        // 尝试恢复本页会话中缓存的独立结果
        var cached = loadPageState(pageKind).lastResult;
        var cachedJobId = loadPageState(pageKind).jobId;
        if (
          cached &&
          cachedJobId &&
          opts.onDone &&
          cached.kind === expectedKind
        ) {
          if (logFrontend) logFrontend("已恢复本页缓存的输出结果。");
          opts.onDone(cached, { job_id: cachedJobId, kind: expectedKind });
          if (opts.onRestored) opts.onRestored({ job_id: cachedJobId, kind: expectedKind }, null);
          return { job_id: cachedJobId, kind: expectedKind };
        }
        if (opts.onRestored) opts.onRestored(null);
        return null;
      }

      // 双保险：绝不恢复其它 kind 的任务
      if (job.kind && job.kind !== expectedKind) {
        if (opts.onRestored) opts.onRestored(null);
        return null;
      }

      if (opts.scriptLog && job.logs) {
        opts.scriptLog.textContent = job.logs;
        opts.scriptLog.scrollTop = opts.scriptLog.scrollHeight;
      }

      savePagePatch(pageKind, {
        jobId: job.job_id,
        status: job.status,
        finished: job.status !== "running" && job.status !== "pending",
      });
      saveSession({
        activeJobId: job.job_id,
        kind: job.kind,
        status: job.status,
        finished: job.status !== "running" && job.status !== "pending",
      });

      var isRunning = job.status === "running" || job.status === "pending";
      if (isRunning) {
        if (logFrontend) {
          logFrontend(
            "检测到本页任务仍在运行（" +
              (job.kind === "historical" ? "历史统计" : "日常统计") +
              "），已恢复日志流与进度。"
          );
        }
        if (opts.onRunningChange) opts.onRunningChange(true, job);
        var streamOpts = {
          jobId: job.job_id,
          scriptLog: opts.scriptLog,
          logFrontend: logFrontend,
          running: true,
          eventSource: null,
          clearScript: false,
          kind: job.kind,
          pageKind: pageKind,
          onRunningChange: function (v) {
            if (opts.onRunningChange) opts.onRunningChange(v, job);
            streamOpts.running = v;
          },
          setProgressFn: opts.setProgressFn,
          onDone: function (payload) {
            savePagePatch(pageKind, {
              jobId: job.job_id,
              lastResult: payload,
              status: payload.status,
              finished: true,
            });
            if (opts.onDone) opts.onDone(payload, job);
          },
        };
        connectStream(streamOpts);
        if (opts.onRestored) opts.onRestored(job, streamOpts);
        return job;
      }

      if (logFrontend) {
        logFrontend(
          "已恢复本页最近一次任务结果（" +
            (job.kind === "historical" ? "历史统计" : "日常统计") +
            "，状态：" +
            job.status +
            "）。"
        );
      }
      var payload = {
        status: job.status,
        exit_code: job.exit_code,
        summary_simple: job.summary_simple,
        summary_full: job.summary_full,
        files: job.files,
        kind: job.kind,
        created_at: job.created_at,
        finished_at: job.finished_at,
        duration_seconds: job.duration_seconds,
        expected_seconds: job.expected_seconds,
      };
      savePagePatch(pageKind, {
        jobId: job.job_id,
        lastResult: payload,
        status: job.status,
        finished: true,
      });
      if (opts.onDone) opts.onDone(payload, job);
      if (opts.onRestored) opts.onRestored(job, null);
      return job;
    } catch (err) {
      if (logFrontend && !isGatewayError(err)) {
        logFrontend("恢复任务状态失败：" + err.message);
      }
      if (opts.onRestored) opts.onRestored(null);
      return null;
    }
  }

  function copyText(text) {
    var value = String(text || "");
    if (!value.trim()) return Promise.reject(new Error("暂无内容可复制"));
    if (navigator.clipboard && typeof navigator.clipboard.writeText === "function") {
      return navigator.clipboard.writeText(value).catch(function () {
        return fallbackCopy(value);
      });
    }
    return fallbackCopy(value);
  }

  function fallbackCopy(text) {
    return new Promise(function (resolve, reject) {
      var ta = document.createElement("textarea");
      ta.value = text;
      ta.setAttribute("readonly", "");
      ta.style.position = "fixed";
      ta.style.left = "-9999px";
      document.body.appendChild(ta);
      ta.focus();
      ta.select();
      ta.setSelectionRange(0, ta.value.length);
      var ok = false;
      try {
        ok = document.execCommand("copy");
      } catch (_) {
        ok = false;
      }
      document.body.removeChild(ta);
      if (ok) resolve();
      else reject(new Error("浏览器拒绝复制，请手动选择文本复制"));
    });
  }

  function bindCopyButtons(summarySimple, summaryFull, logFrontend) {
    document.querySelectorAll("[data-copy]").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var key = btn.dataset.copy;
        var el = key === "simple" ? summarySimple : summaryFull;
        var text = el ? el.value : "";
        copyText(text)
          .then(function () {
            if (logFrontend) {
              logFrontend("已复制" + (key === "simple" ? "简版" : "全量版") + "摘要到剪贴板");
            }
            var old = btn.textContent;
            btn.textContent = "已复制";
            setTimeout(function () {
              btn.textContent = old || "复制";
            }, 1500);
          })
          .catch(function (err) {
            if (logFrontend) logFrontend("复制失败：" + (err.message || err));
            if (el) {
              el.focus();
              el.select();
            }
          });
      });
    });
  }

  function badgeFromStatus(status) {
    if (status === "success") return "success";
    if (status === "cancelled") return "cancelled";
    if (status === "failed") return "error";
    if (status === "running" || status === "pending") return "running";
    return "idle";
  }

  function clearResultUi(opts) {
    if (opts.frontendLog) opts.frontendLog.innerHTML = "";
    if (opts.scriptLog) opts.scriptLog.textContent = "";
    if (opts.summarySimple) opts.summarySimple.value = "";
    if (opts.summaryFull) opts.summaryFull.value = "";
    if (opts.completeBanner) {
      opts.completeBanner.style.display = "none";
      opts.completeBanner.textContent = "";
    }
    if (opts.resultPlaceholder) opts.resultPlaceholder.style.display = "";
    setDownloadLink(opts.dlResult, null, "result", null);
    setDownloadLink(opts.dlPublish, null, "publish", null);
    setDownloadLink(opts.dlStats, null, "stats", null);
    setDownloadLink(opts.dlSummarySimple, null, "summary_simple", null);
    setDownloadLink(opts.dlSummaryFull, null, "summary_full", null);
    setDownloadLink(opts.dlAll, null, "all", null);
    if (opts.runStatus) {
      opts.runStatus.className = "run-status";
      opts.runStatus.textContent = "";
    }
    setBadge(opts.statusBadge, "idle");
    if (opts.progressTracker) opts.progressTracker.reset();
  }

  function fetchWithTimeout(url, init, ms) {
    var ctrl = new AbortController();
    var timer = setTimeout(function () {
      ctrl.abort();
    }, ms || 12000);
    var opts = Object.assign({}, init || {}, { signal: ctrl.signal });
    return fetch(url, opts).finally(function () {
      clearTimeout(timer);
    });
  }

  function parseJsonSafe(resp) {
    return resp.text().then(function (text) {
      if (!text) return {};
      try {
        return JSON.parse(text);
      } catch (_) {
        var err = new Error(
          resp.status === 502 || resp.status === 503 || resp.status === 504
            ? "连接中断（HTTP " + resp.status + "）"
            : "服务器返回了无法解析的响应（HTTP " + resp.status + "）"
        );
        err.status = resp.status;
        throw err;
      }
    });
  }

  function isGatewayError(err) {
    var status = err && err.status;
    if (status === 502 || status === 503 || status === 504 || status === 499) return true;
    var name = String((err && err.name) || "");
    var msg = String((err && err.message) || "");
    return (
      name === "AbortError" ||
      /Failed to fetch|NetworkError|网关中断|连接中断|network|ERR_EMPTY|ECONNRESET|aborted/i.test(msg)
    );
  }

  function parseUploadResponse(resp) {
    return resp.text().then(function (text) {
      var data = null;
      if (text) {
        try {
          data = JSON.parse(text);
        } catch (_) {
          data = null;
        }
      }
      return { resp: resp, data: data, text: text || "" };
    });
  }

  function uploadErrorMessage(parsed, fallback) {
    var status = parsed && parsed.resp ? parsed.resp.status : 0;
    var data = parsed && parsed.data;
    var detail = data && (data.detail || data.message);
    if (typeof detail === "string" && detail) return detail;
    if (status === 413) return "文件过大，超过服务器 105MB 限制";
    if (status === 401) return "登录已过期，请刷新页面后重新认证";
    if (status === 502 || status === 503 || status === 504) {
      return "连接中断（HTTP " + status + "）";
    }
    if (status) return fallback || ("上传失败（HTTP " + status + "）");
    return fallback || "上传失败";
  }

  function isRetryableUploadError(err) {
    return isGatewayError(err);
  }

  function sleep(ms) {
    return new Promise(function (resolve) {
      setTimeout(resolve, ms);
    });
  }

  function newUploadId() {
    if (typeof crypto !== "undefined" && crypto.randomUUID) return crypto.randomUUID();
    return (
      "u" +
      Date.now().toString(16) +
      "-" +
      Math.random().toString(16).slice(2, 10) +
      Math.random().toString(16).slice(2, 6)
    );
  }

  function destFromUploadUrl(url) {
    if (String(url).indexOf("/api/upload/exclude") >= 0) {
      var m = String(url).match(/[?&]year=(\d+)/);
      return { dest: "exclude", year: m ? m[1] : "2026" };
    }
    return { dest: "newdata", year: null };
  }

  function postChunkOnce(chunkUrl, blob) {
    var form = new FormData();
    form.append("file", blob, "chunk.bin");
    return fetch(chunkUrl, { method: "POST", body: form })
      .then(parseUploadResponse)
      .then(function (parsed) {
        if (parsed.resp.ok) return parsed.data || {};
        var err = new Error(uploadErrorMessage(parsed, "分片上传失败"));
        err.status = parsed.resp.status;
        throw err;
      });
  }

  function postChunkWithRetry(chunkUrl, blob, log, attempts, label) {
    attempts = attempts || 5;
    function once(n) {
      return postChunkOnce(chunkUrl, blob).catch(function (err) {
        if (n < attempts && isRetryableUploadError(err)) {
          var waitSec = Math.min(8, n * 2);
          log((label || "分片") + " 中断（第 " + n + " 次），" + waitSec + " 秒后重试…");
          return sleep(waitSec * 1000).then(function () {
            return once(n + 1);
          });
        }
        throw err;
      });
    }
    return once(1);
  }

  function postFileWithRetry(url, file, opts) {
    opts = opts || {};
    var log = opts.log || function () {};
    var maxBytes = opts.maxBytes || 150 * 1024 * 1024;
    var chunkSize = opts.chunkSize || 2 * 1024 * 1024;
    if (file && typeof file.size === "number" && file.size > maxBytes) {
      return Promise.reject(new Error("文件大小超过 150MB 限制"));
    }
    if (!file || !file.size) {
      return Promise.reject(new Error("文件为空"));
    }
    var meta = destFromUploadUrl(url);
    var total = Math.max(1, Math.ceil(file.size / chunkSize));
    var uploadId = newUploadId();
    log("开始分片上传（" + total + " 片，每片约 " + Math.round(chunkSize / 1024) + "KB）…");
    var i = 0;
    function sendNext() {
      var start = i * chunkSize;
      var blob = file.slice(start, Math.min(start + chunkSize, file.size));
      var chunkUrl =
        "/api/upload/chunk?upload_id=" +
        encodeURIComponent(uploadId) +
        "&index=" +
        i +
        "&total=" +
        total +
        "&filename=" +
        encodeURIComponent(file.name) +
        "&dest=" +
        encodeURIComponent(meta.dest);
      if (meta.year) chunkUrl += "&year=" + encodeURIComponent(meta.year);
      if (i + 1 === total) {
        log("最后一片已收齐，正在写入服务器…");
      }
      return postChunkWithRetry(chunkUrl, blob, log, 5, "片 " + (i + 1) + "/" + total).then(function (data) {
        i += 1;
        if (i < total) {
          log("已上传 " + i + "/" + total + " 片");
          return sendNext();
        }
        if (!data || !data.filename) {
          throw new Error("分片已传完，但未收到合并结果");
        }
        return data;
      });
    }
    return sendNext();
  }

  function renderRuntimeVersions(scriptVersion, webVersion, scriptName) {
    var el = document.getElementById("runtimeVersions");
    if (!el) return;
    var script = String(scriptVersion || "").replace(/^v/i, "");
    var web = String(webVersion || "").trim();
    var name = String(scriptName || "").trim();
    var parts = [];
    parts.push(script ? "分析脚本 v" + script : "分析脚本 …");
    if (web) parts.push("页面 " + web);
    el.textContent = parts.join(" · ");
    el.title = name
      ? "当前调用 " + name
      : script
        ? "当前调用 v" + script
        : "正在读取分析脚本版本";
  }

  function initRuntimeVersions() {
    var pageVersion =
      (document.documentElement && document.documentElement.getAttribute("data-web-version")) || "";
    renderRuntimeVersions("", pageVersion, "");
    fetch("/api/health")
      .then(function (resp) {
        return resp.ok ? resp.json() : null;
      })
      .then(function (data) {
        if (!data) return;
        renderRuntimeVersions(
          data.script_version || "",
          data.web_version || pageVersion,
          data.script_name || ""
        );
      })
      .catch(function () {
        /* keep page version from HTML */
      });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initRuntimeVersions);
  } else {
    initRuntimeVersions();
  }

  function startProgressFromJob(tracker, job, fallbackExpected) {
    if (!tracker) return;
    var expected =
      (job && Number(job.expected_seconds) > 0 && Number(job.expected_seconds)) ||
      fallbackExpected ||
      DEFAULT_EXPECTED;
    var startedAtMs =
      job && job.created_at ? Number(job.created_at) * 1000 : Date.now();
    tracker.start({ startedAtMs: startedAtMs, expectedSeconds: expected });
  }

  global.GovUI = {
    now: now,
    escapeHtml: escapeHtml,
    formatDuration: formatDuration,
    createLogger: createLogger,
    restoreFrontendLog: restoreFrontendLog,
    bindTabs: bindTabs,
    setBadge: setBadge,
    setProgress: setProgress,
    showResult: showResult,
    connectStream: connectStream,
    pollJob: pollJob,
    stopJob: stopJob,
    restoreLatestJob: restoreLatestJob,
    saveSession: saveSession,
    loadSession: loadSession,
    clearSession: clearSession,
    beginPageLifecycle: beginPageLifecycle,
    savePagePatch: savePagePatch,
    loadPageState: loadPageState,
    bindCopyButtons: bindCopyButtons,
    copyText: copyText,
    badgeFromStatus: badgeFromStatus,
    setDownloadLink: setDownloadLink,
    createProgressTracker: createProgressTracker,
    fetchExpectedSeconds: fetchExpectedSeconds,
    fetchWithTimeout: fetchWithTimeout,
    parseJsonSafe: parseJsonSafe,
    isGatewayError: isGatewayError,
    clearResultUi: clearResultUi,
    startProgressFromJob: startProgressFromJob,
    postFileWithRetry: postFileWithRetry,
    trimSummaryText: trimSummaryText,
    initRuntimeVersions: initRuntimeVersions,
    DEFAULT_EXPECTED: DEFAULT_EXPECTED,
  };
})(window);
