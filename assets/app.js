const state = {
  templates: [],
  templateMap: new Map(),
  activeTemplateId: null,
  rows: [],
  headers: [],
  datasetName: "불러온 데이터 없음",
  datasetSheet: "-",
  activeRowIndex: 0,
  selectedRowKeys: new Set(),
  searchQuery: "",
};

const dom = {};

window.addEventListener("load", async () => {
  if (!window.XLSX || !window.html2pdf || !window.JSZip) {
    alert("필수 라이브러리를 불러오지 못했습니다. 네트워크 상태를 확인하세요.");
    return;
  }

  cacheDom();
  bindEvents();
  await loadTemplates();
  await loadSampleData();
});

function cacheDom() {
  dom.datasetName = document.querySelector("#dataset-name");
  dom.datasetCount = document.querySelector("#dataset-count");
  dom.datasetSheet = document.querySelector("#dataset-sheet");
  dom.templateSelect = document.querySelector("#template-select");
  dom.recordList = document.querySelector("#record-list");
  dom.recordSearch = document.querySelector("#record-search");
  dom.previewCanvas = document.querySelector("#preview-canvas");
  dom.activeTemplateName = document.querySelector("#active-template-name");
  dom.activeRecordName = document.querySelector("#active-record-name");
  dom.selectionSummary = document.querySelector("#selection-summary");
  dom.fileInput = document.querySelector("#data-file-input");
  dom.loadSampleButton = document.querySelector("#load-sample-btn");
  dom.exportCurrentButton = document.querySelector("#export-current-btn");
  dom.exportBatchButton = document.querySelector("#export-batch-btn");
  dom.printPreviewButton = document.querySelector("#print-preview-btn");
  dom.selectVisibleButton = document.querySelector("#select-visible-btn");
  dom.clearSelectionButton = document.querySelector("#clear-selection-btn");
}

function bindEvents() {
  dom.fileInput.addEventListener("change", handleFileUpload);
  dom.loadSampleButton.addEventListener("click", loadSampleData);
  dom.recordSearch.addEventListener("input", (event) => {
    state.searchQuery = event.target.value.trim().toLowerCase();
    renderDataGrid();
  });
  dom.exportCurrentButton.addEventListener("click", exportCurrentPdf);
  dom.exportBatchButton.addEventListener("click", exportBatchZip);
  dom.printPreviewButton.addEventListener("click", () => window.print());
  dom.selectVisibleButton.addEventListener("click", selectVisibleRecords);
  dom.clearSelectionButton.addEventListener("click", () => {
    state.selectedRowKeys.clear();
    renderDataGrid();
    updateStatus();
  });
  dom.templateSelect.addEventListener("change", (event) => {
    state.activeTemplateId = event.target.value;
    updateTemplateSelection();
    renderPreview();
    updateStatus();
  });
}

async function loadTemplates() {
  const response = await fetch("./templates/catalog.json");
  const catalog = await response.json();

  const templates = await Promise.all(
    catalog.map(async (entry) => {
      const [metaResponse, templateResponse] = await Promise.all([
        fetch(`./templates/${entry.meta.replace("./", "")}`),
        fetch(`./templates/${entry.template.replace("./", "")}`),
      ]);

      return {
        ...(await metaResponse.json()),
        html: await templateResponse.text(),
      };
    })
  );

  state.templates = templates;
  state.templateMap = new Map(templates.map((template) => [template.id, template]));
  state.activeTemplateId = templates[0]?.id ?? null;
  renderTemplateSelect();
  updateTemplateSelection();
}

async function loadSampleData() {
  const response = await fetch("./data/dummy-data-namok.xlsx");
  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
  applyRows(rows, "dummy-data-namok.xlsx", sheetName);
}

async function handleFileUpload(event) {
  const [file] = event.target.files || [];
  if (!file) return;

  const extension = file.name.split(".").pop().toLowerCase();
  const arrayBuffer = await file.arrayBuffer();
  let rows = [];
  let sheetName = "-";

  if (extension === "json") {
    rows = JSON.parse(new TextDecoder().decode(arrayBuffer));
    sheetName = "json";
  } else {
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    sheetName = workbook.SheetNames.includes("MuseumData") ? "MuseumData" : workbook.SheetNames[0];
    rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
  }

  applyRows(rows, file.name, sheetName);
  event.target.value = "";
}

function applyRows(inputRows, datasetName, sheetName) {
  const rows = Array.isArray(inputRows) ? inputRows : [];
  const normalized = rows
    .map((row, index) => normalizeRow(row, index))
    .filter((row) => Object.keys(row).some((key) => !key.startsWith("__") && String(row[key]).trim() !== ""));

  state.rows = normalized;
  state.headers = collectHeaders(normalized);
  state.datasetName = datasetName;
  state.datasetSheet = sheetName;
  state.activeRowIndex = normalized.length ? 0 : -1;
  state.selectedRowKeys = new Set(normalized.length ? [normalized[0].__key] : []);
  updateDataInfo();
  renderDataGrid();
  renderPreview();
  updateStatus();
}

function normalizeRow(row, index) {
  const normalized = {};
  for (const [key, value] of Object.entries(row || {})) {
    const cleanKey = String(key).trim();
    if (!cleanKey) continue;
    normalized[cleanKey] = value == null ? "" : String(value);
  }

  normalized.__key = `row-${index + 1}`;
  normalized.__sourceIndex = index + 2;
  return normalized;
}

function collectHeaders(rows) {
  const set = new Set();
  rows.forEach((row) => {
    Object.keys(row).forEach((key) => {
      if (!key.startsWith("__")) set.add(key);
    });
  });
  return Array.from(set);
}

function updateDataInfo() {
  dom.datasetName.textContent = state.datasetName;
  dom.datasetCount.textContent = String(state.rows.length);
  dom.datasetSheet.textContent = state.datasetSheet;
}

function renderTemplateSelect() {
  dom.templateSelect.innerHTML = state.templates
    .map(
      (template) =>
        `<option value="${escapeHtmlAttr(template.id)}">${escapeHtml(template.name)} · ${escapeHtml(
          template.paper
        )} · ${escapeHtml(template.orientation)}</option>`
    )
    .join("");
}

function updateTemplateSelection() {
  const activeTemplate = getActiveTemplate();
  dom.activeTemplateName.textContent = activeTemplate ? activeTemplate.name : "템플릿 미선택";
  if (dom.templateSelect) dom.templateSelect.value = state.activeTemplateId || "";
}

function getActiveTemplate() {
  return state.templateMap.get(state.activeTemplateId) || null;
}

function getActiveRow() {
  if (!state.rows.length || state.activeRowIndex < 0) return null;
  return state.rows[state.activeRowIndex] || null;
}

function getVisibleRows() {
  if (!state.searchQuery) return state.rows;

  return state.rows.filter((row) => {
    const haystack = state.headers
      .map((header) => String(row[header] ?? ""))
      .join(" ")
      .toLowerCase();
    return haystack.includes(state.searchQuery);
  });
}

function renderDataGrid() {
  const visibleRows = getVisibleRows();

  if (!visibleRows.length) {
    dom.recordList.innerHTML = `
      <div class="empty-state">
        <h3>표시할 데이터가 없습니다</h3>
        <p>검색어를 바꾸거나 다른 데이터를 불러오세요.</p>
      </div>
    `;
    return;
  }

  const displayHeaders = state.headers;
  const rowsHtml = visibleRows
    .map((row) => {
      const cells = displayHeaders
        .map((header) => {
          const value = String(row[header] ?? "");
          const isLong = value.length > 80 || value.includes("\n");
          const control = isLong
            ? `<textarea class="sheet-textarea" data-row-key="${row.__key}" data-header="${escapeHtmlAttr(header)}">${escapeHtml(value)}</textarea>`
            : `<input class="sheet-input" data-row-key="${row.__key}" data-header="${escapeHtmlAttr(header)}" value="${escapeHtmlAttr(value)}" />`;
          return `<td>${control}</td>`;
        })
        .join("");

      return `
        <tr class="${getActiveRow()?.__key === row.__key ? "active-row" : ""}" data-row-key="${row.__key}">
          <td>
            <div class="row-marker">
              <input class="row-select" type="checkbox" data-check-key="${row.__key}" ${
                state.selectedRowKeys.has(row.__key) ? "checked" : ""
              } />
            </div>
          </td>
          ${cells}
        </tr>
      `;
    })
    .join("");

  dom.recordList.innerHTML = `
    <table class="sheet-table">
      <thead>
        <tr>
          <th>선택</th>
          ${displayHeaders.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}
        </tr>
      </thead>
      <tbody>${rowsHtml}</tbody>
    </table>
  `;

  dom.recordList.querySelectorAll(".sheet-input, .sheet-textarea").forEach((control) => {
    control.addEventListener("focus", () => {
      const rowKey = control.dataset.rowKey;
      state.activeRowIndex = state.rows.findIndex((row) => row.__key === rowKey);
      syncActiveRowHighlight();
      renderPreview();
      updateStatus();
    });

    control.addEventListener("input", (event) => {
      const rowKey = event.target.dataset.rowKey;
      const header = event.target.dataset.header;
      const row = state.rows.find((item) => item.__key === rowKey);
      if (!row || !header) return;
      row[header] = event.target.value;
      state.activeRowIndex = state.rows.findIndex((item) => item.__key === rowKey);
      syncActiveRowHighlight();
      renderPreview();
      updateStatus();
    });
  });

  dom.recordList.querySelectorAll("[data-check-key]").forEach((checkbox) => {
    checkbox.addEventListener("change", () => {
      const rowKey = checkbox.dataset.checkKey;
      if (checkbox.checked) state.selectedRowKeys.add(rowKey);
      else state.selectedRowKeys.delete(rowKey);
      updateStatus();
    });
  });
}

function renderPreview() {
  const activeTemplate = getActiveTemplate();
  const activeRow = getActiveRow();

  if (!activeTemplate || !activeRow) {
    dom.previewCanvas.innerHTML = `
      <div class="empty-state">
        <h3>데이터와 템플릿을 선택하세요</h3>
        <p>왼쪽은 데이터 시트, 오른쪽은 선택한 양식의 출력 결과를 보여줍니다.</p>
      </div>
    `;
    return;
  }

  const paper = document.createElement("div");
  paper.className = `preview-paper ${activeTemplate.orientation === "landscape" ? "landscape" : ""}`;
  paper.innerHTML = renderTemplate(activeTemplate.html, activeRow);
  dom.previewCanvas.innerHTML = "";
  dom.previewCanvas.appendChild(paper);
}

function renderTemplate(html, row) {
  return html.replace(/\[\[(.+?)\]\]/g, (_, token) => renderToken(token.trim(), row));
}

function renderToken(token, row) {
  if (token.startsWith("CHECKMAP:")) return renderCheckMapToken(token.slice(9), row);
  if (token.startsWith("CHECK:")) return renderCheckToken(token.slice(6), row);
  if (token.startsWith("CHOICE:")) return renderChoiceToken(token.slice(7), row);
  if (token.startsWith("CHOICEMAP:")) return renderChoiceMapToken(token.slice(10), row);
  return escapeHtml(getRowValue(row, token));
}

function renderCheckToken(body, row) {
  const [header, labelText] = body.split("|");
  if (!header || !labelText) return "";
  const actual = getRowValue(row, header.trim()).trim();
  const label = labelText.trim();
  const selected = actual === label;
  return `<span class="check-item ${selected ? "selected" : ""}">${escapeHtml(label)}</span>`;
}

function renderCheckMapToken(body, row) {
  const [header, mapText] = body.split("|");
  if (!header || !mapText) return "";
  const actual = getRowValue(row, header.trim()).trim();
  const items = mapText
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean)
    .map((item) => {
      const [key, label] = item.split("=");
      return { key: (key ?? "").trim(), label: (label ?? key ?? "").trim() };
    });

  return items
    .map((item) => `<span class="check-item ${item.key === actual ? "selected" : ""}">${escapeHtml(item.label)}</span>`)
    .join(" ");
}

function renderChoiceToken(body, row) {
  const [header, optionsText] = body.split("|");
  if (!header || !optionsText) return "";

  const actual = getRowValue(row, header.trim()).trim();
  const options = optionsText
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean)
    .map((item) => {
      const [key, label] = item.includes("=") ? item.split("=") : [item, item];
      return {
        key: (key ?? "").trim(),
        label: (label ?? key ?? "").trim(),
      };
    });

  return `<span class="token-choice">${options
    .map(
      (option) =>
        `<span class="choice-pill ${option.key === actual ? "selected" : ""}">${escapeHtml(option.label)}</span>`
    )
    .join("")}</span>`;
}

function renderChoiceMapToken(body, row) {
  const [header, mapText] = body.split("|");
  if (!header || !mapText) return "";

  const actual = getRowValue(row, header.trim()).trim();
  const pairs = mapText
    .split(",")
    .map((pair) => pair.trim())
    .filter(Boolean)
    .map((pair) => {
      const [key, label] = pair.split("=");
      return { key: key?.trim() ?? "", label: label?.trim() ?? "" };
    });

  return `<span class="token-choicemap">${pairs
    .map(({ key, label }) => `<span class="map-pill ${key === actual ? "selected" : ""}">${escapeHtml(label || key)}</span>`)
    .join("")}</span>`;
}

function getRowValue(row, header) {
  return row?.[header] == null ? "" : String(row[header]);
}

function extractTokens(html) {
  return Array.from(html.matchAll(/\[\[(.+?)\]\]/g), (match) => match[1].trim());
}

function updateStatus() {
  const activeRow = getActiveRow();
  dom.activeRecordName.textContent = activeRow
    ? `${activeRow["명칭"] || activeRow["자료번호"] || "선택 항목"}`
    : "레코드 미선택";
  dom.selectionSummary.textContent = `선택 ${state.selectedRowKeys.size}건`;
}

function syncActiveRowHighlight() {
  const activeKey = getActiveRow()?.__key;
  dom.recordList.querySelectorAll("tbody tr[data-row-key]").forEach((row) => {
    row.classList.toggle("active-row", row.dataset.rowKey === activeKey);
  });
}

async function exportCurrentPdf() {
  const activeRow = getActiveRow();
  const activeTemplate = getActiveTemplate();
  if (!activeRow || !activeTemplate) {
    alert("데이터와 템플릿을 먼저 선택하세요.");
    return;
  }
  await exportRowToPdf(activeRow, activeTemplate, buildFileName(activeTemplate.fileNamePattern, activeRow));
}

async function exportBatchZip() {
  const activeTemplate = getActiveTemplate();
  if (!activeTemplate || !state.rows.length) {
    alert("데이터와 템플릿을 먼저 선택하세요.");
    return;
  }

  const selected = state.rows.filter((row) => state.selectedRowKeys.has(row.__key));
  const rowsToExport = selected.length ? selected : state.rows;
  const zip = new JSZip();

  dom.exportBatchButton.disabled = true;
  dom.exportBatchButton.textContent = "ZIP 생성 중...";

  try {
    for (const row of rowsToExport) {
      const blob = await renderPdfBlob(row, activeTemplate);
      zip.file(buildFileName(activeTemplate.fileNamePattern, row), blob);
    }

    downloadBlob(await zip.generateAsync({ type: "blob" }), `${activeTemplate.id}-batch.zip`);
  } finally {
    dom.exportBatchButton.disabled = false;
    dom.exportBatchButton.textContent = "선택 항목 ZIP";
  }
}

async function exportRowToPdf(row, template, fileName) {
  const blob = await renderPdfBlob(row, template);
  downloadBlob(blob, fileName);
}

async function renderPdfBlob(row, template) {
  const sandbox = document.createElement("div");
  sandbox.style.position = "fixed";
  sandbox.style.left = "-99999px";
  sandbox.style.top = "0";
  sandbox.innerHTML = `
    <div class="preview-paper ${template.orientation === "landscape" ? "landscape" : ""}">
      ${renderTemplate(template.html, row)}
    </div>
  `;
  document.body.appendChild(sandbox);

  const node = sandbox.firstElementChild;
  const options = {
    margin: 0,
    filename: "temp.pdf",
    image: { type: "jpeg", quality: 0.98 },
    html2canvas: { scale: 2, useCORS: true, backgroundColor: "#ffffff" },
    jsPDF: {
      unit: "mm",
      format: "a4",
      orientation: template.orientation === "landscape" ? "landscape" : "portrait",
    },
  };

  try {
    const worker = window.html2pdf().set(options).from(node).toPdf();
    const pdf = await worker.get("pdf");
    return pdf.output("blob");
  } finally {
    sandbox.remove();
  }
}

function buildFileName(pattern, row) {
  const base = (pattern || "NAMOK-{{자료번호}}-{{세부번호}}.pdf").replace(/\{\{(.+?)\}\}/g, (_, key) => {
    const value = getRowValue(row, key.trim()).trim();
    return value || "Unknown";
  });

  return sanitizeFileName(base);
}

function sanitizeFileName(fileName) {
  return fileName
    .replace(/\//g, "-")
    .replace(/\\/g, "-")
    .replace(/:/g, "-")
    .replace(/\*/g, "-")
    .replace(/\?/g, "-")
    .replace(/"/g, "'")
    .replace(/</g, "(")
    .replace(/>/g, ")")
    .replace(/\|/g, "-");
}

function selectVisibleRecords() {
  getVisibleRows().forEach((row) => state.selectedRowKeys.add(row.__key));
  renderDataGrid();
  updateStatus();
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function escapeHtmlAttr(value) {
  return escapeHtml(value).replace(/\n/g, "&#10;");
}
