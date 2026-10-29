(function () {
  "use strict";

  const $ = (sel, root) => (root || document).querySelector(sel);
  const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));

  const state = {
    file: null,
    report: null,
    filter: "all",
  };

  function esc(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function setStatus(message, kind) {
    const el = $("#status");
    el.textContent = message || "";
    el.className = "status" + (kind ? " is-" + kind : "");
  }

  function bytesLabel(n) {
    if (!n) return "";
    if (n < 1024) return n + " B";
    if (n < 1024 * 1024) return Math.round(n / 1024) + " KB";
    return (n / (1024 * 1024)).toFixed(1) + " MB";
  }

  function validateFile(file) {
    if (!file) return "Choose a .docx or .pptx file first.";
    const name = (file.name || "").toLowerCase();
    if (!name.endsWith(".docx") && !name.endsWith(".pptx")) {
      return "Only .docx and .pptx files are supported in this version.";
    }
    if (file.size === 0) return "That file is empty.";
    if (file.size > 15 * 1024 * 1024) return "That file is larger than 15 MB.";
    return null;
  }

  function selectFile(file) {
    const error = validateFile(file);
    if (error) {
      state.file = null;
      $("#file-name").textContent = error;
      setStatus(error, "error");
      return;
    }
    state.file = file;
    $("#file-name").textContent = file.name + " · " + bytesLabel(file.size);
    setStatus("Ready. Check the consent box, then run the preflight.", "");
  }

  function initUpload() {
    const input = $("#file");
    const dropzone = $("#dropzone");

    $("#choose-file").addEventListener("click", function () {
      input.click();
    });

    input.addEventListener("change", function () {
      if (input.files && input.files[0]) selectFile(input.files[0]);
    });

    ["dragenter", "dragover"].forEach(function (name) {
      dropzone.addEventListener(name, function (event) {
        event.preventDefault();
        dropzone.classList.add("is-dragover");
      });
    });
    ["dragleave", "drop"].forEach(function (name) {
      dropzone.addEventListener(name, function (event) {
        event.preventDefault();
        dropzone.classList.remove("is-dragover");
      });
    });
    dropzone.addEventListener("drop", function (event) {
      const files = event.dataTransfer && event.dataTransfer.files;
      if (files && files[0]) selectFile(files[0]);
    });

    $("#scan-form").addEventListener("submit", function (event) {
      event.preventDefault();
      const error = validateFile(state.file);
      if (error) {
        setStatus(error, "error");
        return;
      }
      if (!$("#consent").checked) {
        setStatus("Please confirm you have the right to scan this file.", "error");
        $("#consent").focus();
        return;
      }
      runScan();
    });
  }

  async function runScan() {
    const button = $("#scan-button");
    button.disabled = true;
    setStatus("Scanning locally in your browser. No file is uploaded.", "loading");
    try {
      const data = await window.PreflightScan.scanFile(state.file);
      state.report = data;
      renderReport(data);
      setStatus("Done. The file never left this tab. Review the findings and download the report.", "");
      $("#results").hidden = false;
      $("#results").scrollIntoView({ behavior: "smooth", block: "start" });
    } catch (err) {
      setStatus((err && err.message) || "The scan failed. Please try another file.", "error");
    } finally {
      button.disabled = false;
    }
  }

  function severityLabel(value) {
    return { critical: "Critical", serious: "Serious", moderate: "Moderate", minor: "Minor" }[value] || value;
  }

  function renderReport(report) {
    $("#results-file").textContent = report.file.name + " · " + bytesLabel(report.file.bytes) + " · scanned " + new Date(report.meta.generated_at).toLocaleString();
    $("#score").textContent = report.score;
    $("#grade").textContent = report.grade;
    $("#count-critical").textContent = report.counts.critical;
    $("#count-serious").textContent = report.counts.serious;
    $("#count-moderate").textContent = report.counts.moderate;
    $("#count-minor").textContent = report.counts.minor;
    $("#score-bar-fill").style.width = Math.max(0, Math.min(100, report.score)) + "%";

    const ai = report.ai || {};
    $("#ai-summary").textContent = ai.summary || "";
    $("#ai-badge").textContent = ai.enabled ? "AI summary" : "On-device rules";
    const priorities = $("#ai-priorities");
    priorities.innerHTML = "";
    (ai.priorities || []).forEach(function (item) {
      const li = document.createElement("li");
      li.textContent = item;
      priorities.appendChild(li);
    });
    $("#ai-limits").textContent = ai.limitations || "";

    const instanceCount = report.issues.reduce(function (sum, issue) { return sum + issue.count; }, 0);
    $("#issue-total").textContent = "(" + report.issues.length + " rules · " + instanceCount + " instances)";
    state.filter = "all";
    $$(".filter").forEach(function (btn) { btn.classList.toggle("is-active", btn.dataset.filter === "all"); });
    renderIssues();
  }

  function renderIssues() {
    const list = $("#issue-list");
    list.innerHTML = "";
    if (!state.report) return;
    const issues = state.report.issues.filter(function (issue) {
      return state.filter === "all" || issue.severity === state.filter;
    });
    $("#empty-filter").hidden = issues.length !== 0;

    issues.forEach(function (issue) {
      const article = document.createElement("article");
      article.className = "issue " + issue.severity;

      const top = document.createElement("div");
      top.className = "issue-top";
      const pill = document.createElement("span");
      pill.className = "pill " + issue.severity;
      pill.textContent = severityLabel(issue.severity);
      const title = document.createElement("h4");
      title.textContent = issue.title;
      top.appendChild(pill);
      top.appendChild(title);
      article.appendChild(top);

      const meta = document.createElement("p");
      meta.className = "issue-meta";
      meta.textContent = "WCAG " + issue.wcag + " · " + issue.count + (issue.count === 1 ? " instance" : " instances");
      article.appendChild(meta);

      const why = document.createElement("p");
      why.textContent = issue.why;
      article.appendChild(why);

      const fix = document.createElement("p");
      fix.className = "fix";
      fix.textContent = "Fix: " + issue.fix;
      article.appendChild(fix);

      if (issue.locations && issue.locations.length) {
        const ul = document.createElement("ul");
        ul.className = "locations";
        issue.locations.forEach(function (loc) {
          const li = document.createElement("li");
          li.textContent = loc;
          ul.appendChild(li);
        });
        article.appendChild(ul);
      }
      list.appendChild(article);
    });
  }

  function initFilters() {
    $$(".filter").forEach(function (button) {
      button.addEventListener("click", function () {
        state.filter = button.dataset.filter;
        $$(".filter").forEach(function (btn) { btn.classList.toggle("is-active", btn === button); });
        renderIssues();
      });
    });
  }

  function reportHtml(report) {
    const issues = report.issues.map(function (issue) {
      const locations = (issue.locations || []).map(function (loc) { return "<li>" + esc(loc) + "</li>"; }).join("");
      return (
        '<section class="finding ' + esc(issue.severity) + '">' +
          '<div class="finding-top"><span class="pill ' + esc(issue.severity) + '">' + esc(severityLabel(issue.severity)) + "</span>" +
          "<h3>" + esc(issue.title) + "</h3></div>" +
          '<p class="meta">WCAG ' + esc(issue.wcag) + " · " + esc(issue.count) + " instance(s)</p>" +
          "<p>" + esc(issue.why) + "</p>" +
          '<p class="fix"><strong>Fix:</strong> ' + esc(issue.fix) + "</p>" +
          (locations ? '<ul class="loc">' + locations + "</ul>" : "") +
        "</section>"
      );
    }).join("");

    const priorities = (report.ai.priorities || []).map(function (item) { return "<li>" + esc(item) + "</li>"; }).join("");
    const counts = report.counts;

    return "<!doctype html><html lang=\"en\"><head><meta charset=\"utf-8\">" +
      "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">" +
      "<title>Accessibility preflight — " + esc(report.file.name) + "</title>" +
      "<style>" +
      "body{font-family:ui-sans-serif,system-ui,-apple-system,Segoe UI,sans-serif;max-width:820px;margin:40px auto;padding:0 24px;color:#17140f;line-height:1.55}" +
      "h1,h2,h3{font-family:ui-serif,Georgia,serif;line-height:1.15}h1{font-size:30px;margin-bottom:4px}h2{font-size:22px;margin-top:32px}" +
      ".muted{color:#5c564b}.score{display:flex;gap:24px;align-items:baseline;border-top:3px solid #17140f;border-bottom:1px solid #d9d0bf;padding:14px 0;margin:20px 0}" +
      ".score strong{font-size:44px;line-height:1}.counts{display:flex;gap:18px;flex-wrap:wrap}.counts span{font-weight:700}" +
      ".finding{border-left:6px solid #254e70;background:#fffdf7;border:1px solid #d9d0bf;padding:14px 16px;margin:12px 0;border-radius:8px}" +
      ".finding.critical{border-left-color:#a61b1b}.finding.serious{border-left-color:#c2410c}.finding.moderate{border-left-color:#8a5a00}" +
      ".finding-top{display:flex;gap:10px;align-items:center;flex-wrap:wrap}.finding h3{margin:0;font-size:19px}" +
      ".pill{font-size:11px;text-transform:uppercase;letter-spacing:.08em;font-weight:800;border-radius:999px;padding:3px 8px;background:#ede7db}" +
      ".pill.critical{background:#f8d7d7;color:#7c1010}.pill.serious{background:#fbe0cc;color:#7a2e0a}.pill.moderate{background:#f5e8c8;color:#5f4100}.pill.minor{background:#dce7f1;color:#1d3f5c}" +
      ".meta{font-size:13px;color:#5c564b;margin:4px 0 8px}.fix{border-top:1px dashed #d9d0bf;padding-top:8px}.loc{font-size:13px;color:#5c564b}" +
      ".disclaimer{background:#fff8ea;border:1px solid #e8d5b8;padding:14px 16px;border-radius:8px;font-size:14px}" +
      "@media print{body{margin:0}.finding{break-inside:avoid}}" +
      "</style></head><body>" +
      "<p class=\"muted\">Preflight A11y · document accessibility preflight</p>" +
      "<h1>Accessibility preflight report</h1>" +
      "<p class=\"muted\">" + esc(report.file.name) + " · " + esc(report.file.type.toUpperCase()) + " · " + esc(new Date(report.meta.generated_at).toLocaleString()) + "</p>" +
      "<div class=\"score\"><strong>" + esc(report.score) + "/100</strong><div><strong>" + esc(report.grade) + "</strong><br><span class=\"muted\">Preflight score — a sorting aid, not a compliance certificate.</span></div></div>" +
      "<div class=\"counts\"><span>Critical: " + esc(counts.critical) + "</span><span>Serious: " + esc(counts.serious) + "</span><span>Moderate: " + esc(counts.moderate) + "</span><span>Minor: " + esc(counts.minor) + "</span></div>" +
      "<h2>What this means</h2><p>" + esc(report.ai.summary) + "</p>" +
      "<h3>Do these first</h3><ol>" + priorities + "</ol>" +
      "<h2>Findings</h2>" + (issues || "<p>No automatically detectable issues were found.</p>") +
      "<h2>Limitations</h2><p class=\"disclaimer\">" + esc(report.ai.limitations) + "</p>" +
      "<p class=\"muted\">Generated by Preflight A11y rules-v0.1. This report is not legal advice and does not certify WCAG, ADA, EAA or Section 508 conformance.</p>" +
      "</body></html>";
  }

  function downloadReport() {
    if (!state.report) return;
    const html = reportHtml(state.report);
    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const base = (state.report.file.name || "report").replace(/\.[^.]+$/, "");
    a.href = url;
    a.download = base + "-preflight.html";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  function initMisc() {
    $("#download-report").addEventListener("click", downloadReport);
    $("#scan-another").addEventListener("click", function () {
      $("#results").hidden = true;
      state.report = null;
      state.file = null;
      $("#file").value = "";
      $("#file-name").textContent = "No file selected.";
      $("#consent").checked = false;
      setStatus("", "");
      $("#scan").scrollIntoView({ behavior: "smooth", block: "start" });
      $("#choose-file").focus();
    });
  }

  document.addEventListener("DOMContentLoaded", function () {
    initUpload();
    initFilters();
    initMisc();
  });
})();
