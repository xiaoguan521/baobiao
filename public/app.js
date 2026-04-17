const state = {
  authRequired: false,
  sheets: []
};

function toMonthValue(monthKey) {
  if (!monthKey || !/^\d{4}-\d{2}$/.test(monthKey)) return "";
  return monthKey;
}

function getSavedToken() {
  return window.localStorage.getItem("report-api-token") || "";
}

function saveToken(value) {
  window.localStorage.setItem("report-api-token", value);
}

function setTokenHint(message, tone = "neutral") {
  const tokenHint = document.querySelector("#tokenHint");
  tokenHint.textContent = message;
  tokenHint.dataset.tone = tone;
}

function requestHeaders(needsJson = false) {
  const headers = {};
  const token = document.querySelector("#apiToken").value.trim();
  if (token) headers.authorization = `Bearer ${token}`;
  if (needsJson) headers["content-type"] = "application/json";
  return headers;
}

function renderHealth(payload) {
  document.querySelector("#healthText").textContent = payload.ok ? "在线" : "异常";
  document.querySelector("#authText").textContent = payload.authRequired ? "需要 Token" : "无需 Token";
  state.authRequired = Boolean(payload.authRequired);
  setTokenHint(
    state.authRequired ? "当前服务要求 Token，请填写后再发起生成或排查。" : "当前服务未要求 Token。",
    state.authRequired ? "warn" : "ok"
  );
}

function renderSheets(payload) {
  state.sheets = payload.sheets || [];
  const select = document.querySelector("#sheetName");
  select.innerHTML = '<option value="">整本报表</option>';
  state.sheets.forEach((sheet) => {
    const option = document.createElement("option");
    option.value = sheet;
    option.textContent = sheet;
    select.appendChild(option);
  });
}

function formatJsonBlock(payload) {
  const pre = document.createElement("pre");
  pre.textContent = JSON.stringify(payload, null, 2);
  return pre;
}

function renderMessage(container, message, tone = "neutral") {
  container.className = `result-card ${tone}`;
  container.innerHTML = "";
  const paragraph = document.createElement("p");
  paragraph.textContent = message;
  container.appendChild(paragraph);
}

async function downloadWithAuth(file) {
  const response = await fetch(file.downloadUrl, {
    headers: requestHeaders(false)
  });

  if (!response.ok) {
    let message = "下载失败";
    try {
      const payload = await response.json();
      message = payload.error || message;
    } catch (_error) {
      message = await response.text();
    }
    throw new Error(message);
  }

  const blob = await response.blob();
  const blobUrl = window.URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = blobUrl;
  anchor.download = file.name;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => window.URL.revokeObjectURL(blobUrl), 1000);
}

function renderGenerateResult(payload) {
  const container = document.querySelector("#generateResult");
  container.className = "result-card success";
  container.innerHTML = "";

  const headline = document.createElement("div");
  headline.className = "result-header";
  const copy = document.createElement("div");
  copy.innerHTML = `<strong>${payload.file.name}</strong><span>${payload.sheetOnly ? "单 sheet 导出" : "整本导出"}</span>`;
  const downloadButton = document.createElement("button");
  downloadButton.type = "button";
  downloadButton.className = "download-link";
  downloadButton.textContent = "下载文件";
  downloadButton.addEventListener("click", async () => {
    try {
      downloadButton.disabled = true;
      downloadButton.textContent = "下载中...";
      await downloadWithAuth(payload.file);
      downloadButton.textContent = "下载完成";
      window.setTimeout(() => {
        downloadButton.textContent = "下载文件";
        downloadButton.disabled = false;
      }, 1200);
    } catch (error) {
      downloadButton.textContent = "下载文件";
      downloadButton.disabled = false;
      renderMessage(container, error.message, "error");
    }
  });
  headline.appendChild(copy);
  headline.appendChild(downloadButton);
  container.appendChild(headline);

  const meta = document.createElement("div");
  meta.className = "stats-grid";
  const entries = [
    ["月份", payload.month],
    ["Sheet", payload.sheetName || "整本"],
    ["文件大小", `${payload.file.sizeBytes} bytes`],
    ["扫描行数", payload.stats?.scannedRows ?? "-"],
    ["分类命中", payload.stats?.classifiedRows ?? "-"],
    ["未命中网点", payload.stats?.unmatchedOutletRows ?? "-"]
  ];
  entries.forEach(([label, value]) => {
    const item = document.createElement("div");
    item.className = "stat-item";
    item.innerHTML = `<span>${label}</span><strong>${value}</strong>`;
    meta.appendChild(item);
  });
  container.appendChild(meta);
  container.appendChild(formatJsonBlock(payload));
}

function renderTopList(title, items) {
  const block = document.createElement("section");
  block.className = "debug-section";
  const heading = document.createElement("h3");
  heading.textContent = title;
  block.appendChild(heading);

  if (!items || !items.length) {
    const empty = document.createElement("p");
    empty.textContent = "暂无数据。";
    block.appendChild(empty);
    return block;
  }

  const list = document.createElement("ol");
  items.forEach((item) => {
    const row = document.createElement("li");
    row.innerHTML = `<strong>${item.key}</strong><span>${item.count}</span>`;
    list.appendChild(row);
  });
  block.appendChild(list);
  return block;
}

function renderDebugResult(payload) {
  const container = document.querySelector("#debugResult");
  container.className = "result-card success";
  container.innerHTML = "";

  const summary = document.createElement("div");
  summary.className = "stats-grid";
  [
    ["月份", payload.month],
    ["扫描行数", payload.scannedRows],
    ["分类命中", payload.classifiedRows],
    ["未匹配", payload.unmatchedRows]
  ].forEach(([label, value]) => {
    const item = document.createElement("div");
    item.className = "stat-item";
    item.innerHTML = `<span>${label}</span><strong>${value}</strong>`;
    summary.appendChild(item);
  });
  container.appendChild(summary);
  container.appendChild(renderTopList("经办人 Top", payload.unmatchedCreators));
  container.appendChild(renderTopList("地区 Top", payload.unmatchedByRegion));
  container.appendChild(renderTopList("事项 Top", payload.unmatchedByItem));
  container.appendChild(renderTopList("描述 Top", payload.unmatchedDxms));
}

async function loadBootstrap() {
  const [healthResponse, sheetsResponse] = await Promise.all([
    fetch("/api/health"),
    fetch("/api/reports/sheets")
  ]);

  const health = await healthResponse.json();
  const sheets = await sheetsResponse.json();
  renderHealth(health);
  renderSheets(sheets);
}

async function submitGenerate(event) {
  event.preventDefault();
  const result = document.querySelector("#generateResult");
  renderMessage(result, "正在生成报表，请稍候...", "loading");

  const month = document.querySelector("#month").value;
  const sheetName = document.querySelector("#sheetName").value;
  const sheetOnly = document.querySelector("#sheetOnly").checked;

  try {
    const response = await fetch("/api/reports/generate", {
      method: "POST",
      headers: requestHeaders(true),
      body: JSON.stringify({
        month: toMonthValue(month),
        sheetName: sheetName || undefined,
        sheetOnly
      })
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "生成失败");
    renderGenerateResult(payload);
  } catch (error) {
    renderMessage(result, error.message, "error");
  }
}

async function submitDebug(event) {
  event.preventDefault();
  const result = document.querySelector("#debugResult");
  renderMessage(result, "正在查询未匹配网点，请稍候...", "loading");

  const month = document.querySelector("#debugMonth").value;
  const limit = document.querySelector("#debugLimit").value;
  const params = new URLSearchParams({
    month: toMonthValue(month),
    limit
  });

  try {
    const response = await fetch(`/api/reports/debug/unmatched?${params.toString()}`, {
      headers: requestHeaders(false)
    });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || "查询失败");
    renderDebugResult(payload);
  } catch (error) {
    renderMessage(result, error.message, "error");
  }
}

function wireEvents() {
  const tokenInput = document.querySelector("#apiToken");
  tokenInput.value = getSavedToken();
  tokenInput.addEventListener("input", () => {
    saveToken(tokenInput.value.trim());
    setTokenHint(tokenInput.value.trim() ? "Token 已保存在当前浏览器。" : "未保存到服务器，仅保存在当前浏览器。");
  });

  document.querySelector("#generateForm").addEventListener("submit", submitGenerate);
  document.querySelector("#debugForm").addEventListener("submit", submitDebug);
  document.querySelector("#resetResult").addEventListener("click", () => {
    renderMessage(document.querySelector("#generateResult"), "生成结果会显示在这里。");
  });
}

async function main() {
  const today = new Date();
  const defaultMonth = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}`;
  document.querySelector("#month").value = defaultMonth;
  document.querySelector("#debugMonth").value = defaultMonth;

  wireEvents();

  try {
    await loadBootstrap();
  } catch (error) {
    renderMessage(document.querySelector("#generateResult"), `初始化失败：${error.message}`, "error");
    renderMessage(document.querySelector("#debugResult"), `初始化失败：${error.message}`, "error");
  }
}

main();
