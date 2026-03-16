const SAMPLE_CSV = `満足度,来場回数,年齢層,性別,知ったきっかけ,感想
5,初めて,20代,女性,SNS,とても面白かったです
4,2回目,30代,男性,友人,役者さんの演技が良かった
5,初めて,20代,女性,チラシ,また見たいです
3,3回以上,40代,女性,劇団HP,少し長く感じました
4,初めて,30代,女性,SNS,世界観が好きでした`;

const COLUMN_TYPES = {
  category: "カテゴリ",
  numeric: "数値",
  text: "自由記述",
  hidden: "非表示",
};

const GRAPH_TYPES = {
  category: [
    { value: "pie", label: "円グラフ" },
    { value: "bar", label: "棒グラフ" },
    { value: "bar-horizontal", label: "横棒グラフ" },
  ],
  numeric: [
    { value: "bar", label: "棒グラフ" },
    { value: "bar-horizontal", label: "横棒グラフ" },
    { value: "mean-only", label: "平均値のみ表示" },
  ],
};
const TEXT_HINT_WORDS = ["感想", "コメント", "意見", "自由記述", "メッセージ", "note", "comment"];
const NUMERIC_HINT_WORDS = ["年齢", "点数", "満足", "評価", "score", "rate", "数"];
const CATEGORY_HINT_WORDS = ["性別", "きっかけ", "回数", "年代", "区分", "種別"];
const TEXT_STOP_WORDS = new Set(["です", "ます", "でした", "した", "こと", "それ", "ため", "よう", "ので", "から", "また", "とても", "少し", "ような", "ある", "いる", "今回", "のでした"]);

const appState = {
  sourceType: "",
  sourceName: "",
  loadedAt: "",
  headers: [],
  rows: [],
  columns: [],
};

const chartMap = new Map();
let renderToken = 0;

const ui = {
  fileInput: document.getElementById("csvFileInput"),
  textInput: document.getElementById("csvTextInput"),
  loadButton: document.getElementById("loadCsvButton"),
  sampleButton: document.getElementById("loadSampleButton"),
  errorBox: document.getElementById("errorBox"),
  statusBox: document.getElementById("statusBox"),
  overviewGrid: document.getElementById("overviewGrid"),
  analysisCards: document.getElementById("analysisCards"),
  hiddenColumnsBar: document.getElementById("hiddenColumnsBar"),
  hiddenColumnsList: document.getElementById("hiddenColumnsList"),
  emptyState: document.getElementById("emptyState"),
  pdfLayoutSelect: document.getElementById("pdfLayoutSelect"),
  exportPdfButton: document.getElementById("exportPdfButton"),
  exportArea: document.getElementById("resultsExportArea"),
};

init();

function init() {
  bindEvents();
  renderOverview();
}

function bindEvents() {
  ui.loadButton.addEventListener("click", handleLoadCsv);
  ui.sampleButton.addEventListener("click", () => {
    ui.textInput.value = SAMPLE_CSV;
    showStatus("サンプルCSVを入力しました。");
  });
  ui.exportPdfButton.addEventListener("click", exportAllAsPdf);
  ui.analysisCards.addEventListener("change", handleCardControls);
  ui.analysisCards.addEventListener("click", handleCardActions);
  ui.hiddenColumnsList.addEventListener("click", handleCardActions);
}

async function handleLoadCsv() {
  clearMessages();

  try {
    const source = await getCsvSource();
    const parsed = parseCsvText(source.text);
    const dataset = buildDataset(parsed, source);
    Object.assign(appState, dataset);
    renderOverview();
    renderAnalysisCards();
    showStatus("CSVを読み込みました。");
  } catch (error) {
    resetResults();
    showError(error.message || "CSVの読み込みに失敗しました。");
  }
}

async function getCsvSource() {
  const file = ui.fileInput.files[0];
  const text = ui.textInput.value.trim();

  if (file) {
    const fileText = await file.text();
    return {
      text: fileText,
      sourceType: "ファイル",
      sourceName: file.name,
    };
  }

  if (text) {
    return {
      text,
      sourceType: "貼り付け",
      sourceName: "",
    };
  }

  throw new Error("CSVファイルを選ぶか、CSVテキストを貼り付けてください。");
}

function parseCsvText(text) {
  const normalizedText = preprocessCsvText(text);
  if (!normalizedText.trim()) {
    throw new Error("CSVが空です。データを入れてから読み込んでください。");
  }

  const result = Papa.parse(normalizedText, {
    skipEmptyLines: true,
    delimiter: detectDelimiter(normalizedText),
  });

  if (result.errors && result.errors.length) {
    throw new Error("CSVの読み込みに失敗しました。区切りや改行を確認してください。");
  }

  return result.data;
}

function buildDataset(rawRows, source) {
  if (!rawRows.length) {
    throw new Error("CSVが空です。データを入れてから読み込んでください。");
  }

  const headers = normalizeHeaders(rawRows[0]);
  if (!headers.length || headers.every((header) => !header)) {
    throw new Error("ヘッダーが見つかりません。1行目に項目名を入れてください。");
  }

  const bodyRows = rawRows.slice(1).filter((row) => row.some((cell) => String(cell ?? "").trim() !== ""));
  if (!bodyRows.length) {
    throw new Error("データ行がありません。ヘッダーの下に回答データを入れてください。");
  }

  const rows = bodyRows.map((row) => {
    const record = {};
    headers.forEach((header, index) => {
      const key = header || `列${index + 1}`;
      record[key] = String(row[index] ?? "").trim();
    });
    return record;
  });

  const normalizedHeaders = headers.map((header, index) => header || `列${index + 1}`);
  const columns = normalizedHeaders.map((header, index) => buildColumnDefinition(header, index, rows));

  return {
    sourceType: source.sourceType,
    sourceName: source.sourceName,
    loadedAt: formatDateTime(new Date()),
    headers: normalizedHeaders,
    rows,
    columns,
  };
}

function buildColumnDefinition(header, index, rows) {
  const values = rows.map((row) => row[header] ?? "");
  const prepared = prepareColumnValues(values);
  const detection = detectColumnType(prepared, header);
  return {
    id: `col-${index}`,
    name: header,
    values,
    prepared,
    detection,
    inferredType: detection.type,
    type: detection.type,
    graphType: getDefaultGraphType(detection.type),
    analysisCache: {},
  };
}

function prepareColumnValues(values) {
  const trimmed = values.map((value) => normalizeCellValue(String(value ?? "")));
  const filled = trimmed.filter(Boolean);
  const numeric = filled
    .map((value) => Number(sanitizeNumericValue(value)))
    .filter((value) => Number.isFinite(value));

  return {
    trimmed,
    filled,
    numeric,
  };
}

function detectColumnType(prepared, header) {
  const normalized = prepared.filled;
  const lowerHeader = String(header || "").toLowerCase();
  const numericRatio = normalized.length ? prepared.numeric.length / normalized.length : 0;
  const uniqueCount = new Set(normalized).size;
  const uniqueRatio = normalized.length ? uniqueCount / normalized.length : 0;
  const averageLength = normalized.length
    ? normalized.reduce((sum, value) => sum + value.length, 0) / normalized.length
    : 0;
  const hints = [];

  if (!normalized.length) {
    return {
      type: "category",
      hints: ["空欄が多いためカテゴリ扱い"],
      metrics: { numericRatio, uniqueRatio, averageLength },
    };
  }

  if (NUMERIC_HINT_WORDS.some((word) => lowerHeader.includes(word.toLowerCase()))) {
    hints.push("列名から数値寄りと判定");
  }
  if (TEXT_HINT_WORDS.some((word) => lowerHeader.includes(word.toLowerCase()))) {
    hints.push("列名から自由記述寄りと判定");
  }
  if (CATEGORY_HINT_WORDS.some((word) => lowerHeader.includes(word.toLowerCase()))) {
    hints.push("列名からカテゴリ寄りと判定");
  }

  if (prepared.numeric.length >= 3 && numericRatio >= 0.75) {
    hints.push(`数値率 ${Math.round(numericRatio * 100)}%`);
    return {
      type: "numeric",
      hints,
      metrics: { numericRatio, uniqueRatio, averageLength },
    };
  }

  if ((uniqueRatio >= 0.8 && averageLength >= 8) || averageLength >= 18) {
    hints.push(`ユニーク率 ${Math.round(uniqueRatio * 100)}%`);
    hints.push(`平均文字数 ${averageLength.toFixed(1)}`);
    return {
      type: "text",
      hints,
      metrics: { numericRatio, uniqueRatio, averageLength },
    };
  }

  hints.push(`選択肢候補 ${uniqueCount}件`);
  return {
    type: "category",
    hints,
    metrics: { numericRatio, uniqueRatio, averageLength },
  };
}

function getDefaultGraphType(type) {
  if (type === "category") {
    return "pie";
  }
  if (type === "numeric") {
    return "bar";
  }
  return "";
}

function handleCardControls(event) {
  const target = event.target;
  if (!(target instanceof HTMLSelectElement)) {
    return;
  }

  const card = target.closest("[data-column-id]");
  if (!card) {
    return;
  }

  const column = appState.columns.find((item) => item.id === card.dataset.columnId);
  if (!column) {
    return;
  }

  if (target.dataset.control === "type") {
    column.type = target.value;
    column.analysisCache = {};
    if (target.value === "category" || target.value === "numeric") {
      column.graphType = getDefaultGraphType(target.value);
    } else {
      column.graphType = "";
    }
  }

  if (target.dataset.control === "graph") {
    column.graphType = target.value;
  }

  updateColumnCard(column);
}

function handleCardActions(event) {
  const target = event.target;
  if (!(target instanceof HTMLButtonElement)) {
    return;
  }

  if (target.dataset.action !== "png") {
    if (target.dataset.action === "restore-column") {
      const column = appState.columns.find((item) => item.id === target.dataset.columnId);
      if (!column) {
        return;
      }
      column.type = column.inferredType;
      column.graphType = getDefaultGraphType(column.type);
      column.analysisCache = {};
      updateColumnCard(column);
    }
    return;
  }

  const card = target.closest("[data-column-id]");
  if (!card) {
    return;
  }

  saveCardAsPng(card);
}

function renderOverview() {
  const items = [
    { label: "総回答数", value: appState.rows.length ? String(appState.rows.length) : "-" },
    { label: "列数", value: appState.headers.length ? String(appState.headers.length) : "-" },
    { label: "ファイル名", value: appState.sourceName || "-" },
    { label: "読み込み方法", value: appState.sourceType || "-" },
    { label: "読込日時", value: appState.loadedAt || "-" },
  ];

  ui.overviewGrid.innerHTML = items.map((item) => `
    <div class="overview-item">
      <span>${escapeHtml(item.label)}</span>
      <strong>${escapeHtml(item.value)}</strong>
    </div>
  `).join("");
}

function renderAnalysisCards() {
  destroyAllCharts();
  renderToken += 1;
  const currentToken = renderToken;

  const visibleColumns = appState.columns.filter((column) => column.type !== "hidden");
  ui.emptyState.style.display = visibleColumns.length ? "none" : "block";
  ui.analysisCards.innerHTML = visibleColumns.map((column) => createCardMarkup(column)).join("");
  ui.exportPdfButton.disabled = !visibleColumns.length;
  renderHiddenColumnsBar();

  renderChartsIncrementally(visibleColumns, currentToken);
}

function updateColumnCard(column) {
  destroyChartForColumn(column.id);
  const currentCard = ui.analysisCards.querySelector(`[data-column-id="${column.id}"]`);

  if (column.type === "hidden") {
    currentCard?.remove();
    refreshResultsVisibility();
    return;
  }

  const nextCard = htmlToElement(createCardMarkup(column));
  if (currentCard) {
    currentCard.replaceWith(nextCard);
  } else {
    const nextVisibleSibling = findNextVisibleColumn(column.id);
    if (nextVisibleSibling) {
      const siblingElement = ui.analysisCards.querySelector(`[data-column-id="${nextVisibleSibling.id}"]`);
      if (siblingElement) {
        siblingElement.before(nextCard);
      } else {
        ui.analysisCards.appendChild(nextCard);
      }
    } else {
      ui.analysisCards.appendChild(nextCard);
    }
  }

  if (column.type === "category" || column.type === "numeric") {
    renderChartForColumn(column);
  }

  refreshResultsVisibility();
}

function findNextVisibleColumn(columnId) {
  const currentIndex = appState.columns.findIndex((item) => item.id === columnId);
  if (currentIndex === -1) {
    return null;
  }

  for (let index = currentIndex + 1; index < appState.columns.length; index += 1) {
    if (appState.columns[index].type !== "hidden") {
      return appState.columns[index];
    }
  }

  return null;
}

function refreshResultsVisibility() {
  const visibleColumns = appState.columns.filter((column) => column.type !== "hidden");
  ui.emptyState.style.display = visibleColumns.length ? "none" : "block";
  ui.exportPdfButton.disabled = !visibleColumns.length;
  renderHiddenColumnsBar();
}

function renderHiddenColumnsBar() {
  const hiddenColumns = appState.columns.filter((column) => column.type === "hidden");
  ui.hiddenColumnsBar.classList.toggle("is-hidden", hiddenColumns.length === 0);
  ui.hiddenColumnsList.innerHTML = hiddenColumns.map((column) => `
    <div class="hidden-column-chip">
      <span>${escapeHtml(column.name)}</span>
      <button type="button" data-action="restore-column" data-column-id="${escapeHtml(column.id)}">戻す</button>
    </div>
  `).join("");
}

function createCardMarkup(column) {
  const analysis = getColumnAnalysis(column);
  const graphOptions = getGraphOptionsMarkup(column.type, column.graphType);
  const graphHiddenClass = column.type === "text" ? "is-hidden" : "";
  const hintItems = column.detection?.hints?.length
    ? column.detection.hints.map((hint) => `<li>${escapeHtml(hint)}</li>`).join("")
    : "";

  return `
    <article class="analysis-card" data-column-id="${escapeHtml(column.id)}">
      <div class="card-head">
        <div>
          <h3>${escapeHtml(column.name)}</h3>
          <p>自動判定: ${escapeHtml(COLUMN_TYPES[column.inferredType])}</p>
        </div>
      </div>

      ${hintItems ? `<ul class="hint-list">${hintItems}</ul>` : ""}

      <div class="card-controls">
        <label class="control">
          <span>列タイプ</span>
          <select data-control="type">
            ${getTypeOptionsMarkup(column.type)}
          </select>
        </label>
        <label class="control ${column.type === "text" ? "is-hidden" : ""}">
          <span>グラフ形式</span>
          <select data-control="graph">
            ${graphOptions}
          </select>
        </label>
      </div>

      ${renderAnalysisMarkup(column, analysis)}

      <div class="chart-wrap ${graphHiddenClass}">
        <canvas id="chart-${escapeHtml(column.id)}"></canvas>
      </div>

      <div class="card-actions">
        <button type="button" class="card-button" data-action="png">PNG保存</button>
      </div>
    </article>
  `;
}

function renderAnalysisMarkup(column, analysis) {
  if (column.type === "category") {
    return `
      <ul class="stats-list">
        <li>有効回答数: ${analysis.total}</li>
        <li>選択肢数: ${analysis.labels.length}</li>
        <li>最多回答: ${analysis.labels[0] ? `${escapeHtml(analysis.labels[0])} (${analysis.counts[0]}件 / ${analysis.percentages[0]}%)` : "-"}</li>
      </ul>
    `;
  }

  if (column.type === "numeric") {
    return `
      <ul class="stats-list">
        <li>件数: ${analysis.count}</li>
        <li>平均値: ${analysis.mean}</li>
        <li>最小値: ${analysis.min}</li>
        <li>最大値: ${analysis.max}</li>
        <li>中央値: ${analysis.median}</li>
        <li>標準偏差: ${analysis.stdDev}</li>
        <li>分布区間数: ${analysis.labels.length}</li>
      </ul>
    `;
  }

  const samples = analysis.samples.length
    ? analysis.samples.map((sample) => `<li>${escapeHtml(sample)}</li>`).join("")
    : "<li>サンプルはありません。</li>";
  const keywords = analysis.keywords.length
    ? analysis.keywords.map((item) => `<li>${escapeHtml(item.word)} (${item.count})</li>`).join("")
    : "<li>頻出語は見つかりませんでした</li>";

  return `
    <ul class="stats-list">
      <li>回答件数: ${analysis.responseCount}</li>
      <li>空欄件数: ${analysis.emptyCount}</li>
      <li>平均文字数: ${analysis.averageLength}</li>
    </ul>
    <ul class="keyword-list">
      ${keywords}
    </ul>
    <ul class="sample-list">
      ${samples}
    </ul>
  `;
}

function getTypeOptionsMarkup(selectedType) {
  return Object.entries(COLUMN_TYPES).map(([value, label]) => `
    <option value="${value}" ${value === selectedType ? "selected" : ""}>${label}</option>
  `).join("");
}

function getGraphOptionsMarkup(type, selectedGraphType) {
  const options = GRAPH_TYPES[type] || [];
  return options.map((option) => `
    <option value="${option.value}" ${option.value === selectedGraphType ? "selected" : ""}>${option.label}</option>
  `).join("");
}

function getColumnAnalysis(column) {
  if (column.analysisCache[column.type]) {
    return column.analysisCache[column.type];
  }

  let analysis;
  if (column.type === "numeric") {
    analysis = analyzeNumericColumn(column.prepared);
  } else if (column.type === "text") {
    analysis = analyzeTextColumn(column.prepared);
  } else {
    analysis = analyzeCategoryColumn(column.prepared);
  }

  column.analysisCache[column.type] = analysis;
  return analysis;
}

function analyzeCategoryColumn(prepared) {
  const map = new Map();
  prepared.trimmed.forEach((normalized, index) => {
    if (!normalized) {
      return;
    }

    if (!map.has(normalized)) {
      map.set(normalized, { label: normalized, count: 0, firstIndex: index });
    }
    map.get(normalized).count += 1;
  });

  const items = Array.from(map.values()).sort((a, b) => {
    if (b.count !== a.count) {
      return b.count - a.count;
    }
    return a.firstIndex - b.firstIndex;
  });

  const total = items.reduce((sum, item) => sum + item.count, 0);
  return {
    total,
    labels: items.map((item) => item.label),
    counts: items.map((item) => item.count),
    percentages: items.map((item) => total ? ((item.count / total) * 100).toFixed(1) : "0.0"),
  };
}

function analyzeNumericColumn(values) {
  const numbers = preparedNumbers(values);

  if (!numbers.length) {
    return {
      count: 0,
      mean: "-",
      min: "-",
      max: "-",
      median: "-",
      stdDev: "-",
      labels: [],
      counts: [],
      averageOnly: 0,
    };
  }

  const sorted = [...numbers].sort((a, b) => a - b);
  const frequencyMap = new Map();
  sorted.forEach((value) => {
    const key = String(value);
    frequencyMap.set(key, (frequencyMap.get(key) || 0) + 1);
  });

  const mean = numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
  const middle = Math.floor(sorted.length / 2);
  const median = sorted.length % 2 === 0 ? (sorted[middle - 1] + sorted[middle]) / 2 : sorted[middle];
  const variance = numbers.reduce((sum, value) => sum + ((value - mean) ** 2), 0) / numbers.length;
  const distribution = buildNumericDistribution(sorted);

  return {
    count: numbers.length,
    mean: formatNumber(mean),
    min: formatNumber(sorted[0]),
    max: formatNumber(sorted[sorted.length - 1]),
    median: formatNumber(median),
    stdDev: formatNumber(Math.sqrt(variance)),
    labels: distribution.labels,
    counts: distribution.counts,
    averageOnly: mean,
  };
}

function analyzeTextColumn(prepared) {
  const samples = prepared.filled.slice(0, 5);
  const totalLength = prepared.filled.reduce((sum, value) => sum + value.length, 0);
  return {
    responseCount: prepared.filled.length,
    emptyCount: prepared.trimmed.length - prepared.filled.length,
    averageLength: prepared.filled.length ? formatNumber(totalLength / prepared.filled.length) : "-",
    keywords: extractFrequentWords(prepared.filled),
    samples,
  };
}

function preparedNumbers(prepared) {
  return prepared.numeric;
}

function renderChartForColumn(column) {
  const canvas = document.getElementById(`chart-${column.id}`);
  if (!canvas) {
    return;
  }

  const analysis = getColumnAnalysis(column);
  const config = buildChartConfig(column, analysis);
  if (!config) {
    return;
  }

  const existing = chartMap.get(column.id);
  if (existing) {
    existing.destroy();
  }

  const chart = new Chart(canvas, config);
  chartMap.set(column.id, chart);
}

function buildChartConfig(column, analysis) {
  if (column.type === "category") {
    const type = column.graphType === "pie" ? "pie" : "bar";
    const horizontal = column.graphType === "bar-horizontal";
    return {
      type,
      data: {
        labels: analysis.labels,
        datasets: [{
          label: column.name,
          data: analysis.counts,
          backgroundColor: getChartColors(analysis.labels.length),
          borderColor: "rgba(17, 17, 17, 0.15)",
          borderWidth: 1,
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: horizontal ? "y" : "x",
        plugins: {
          legend: { display: type === "pie" },
        },
        scales: type === "pie" ? {} : {
          y: { beginAtZero: true },
        },
      },
    };
  }

  if (column.type === "numeric") {
    if (column.graphType === "mean-only") {
      return {
        type: "bar",
        data: {
          labels: ["平均値"],
          datasets: [{
            label: column.name,
            data: [analysis.averageOnly],
            backgroundColor: ["rgba(70, 70, 70, 0.78)"],
            borderWidth: 0,
          }],
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
          },
          scales: {
            y: { beginAtZero: true },
          },
        },
      };
    }

    return {
      type: "bar",
      data: {
        labels: analysis.labels,
        datasets: [{
          label: `${column.name} の頻度`,
          data: analysis.counts,
          backgroundColor: getChartColors(analysis.labels.length),
          borderWidth: 0,
        }],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: column.graphType === "bar-horizontal" ? "y" : "x",
        plugins: {
          legend: { display: false },
        },
        scales: {
          y: { beginAtZero: true },
        },
      },
    };
  }

  return null;
}

function getChartColors(count) {
  const palette = [
    "rgba(78, 121, 167, 0.82)",
    "rgba(242, 142, 44, 0.82)",
    "rgba(225, 87, 89, 0.82)",
    "rgba(118, 183, 178, 0.82)",
    "rgba(89, 161, 79, 0.82)",
    "rgba(237, 201, 72, 0.82)",
    "rgba(176, 122, 161, 0.82)",
    "rgba(255, 157, 167, 0.82)",
    "rgba(156, 117, 95, 0.82)",
    "rgba(186, 176, 172, 0.82)",
  ];

  return Array.from({ length: count }, (_, index) => palette[index % palette.length]);
}

function renderChartsIncrementally(columns, token, startIndex = 0) {
  if (token !== renderToken) {
    return;
  }

  const batch = columns.slice(startIndex, startIndex + 4);
  batch.forEach((column) => {
    if (column.type === "category" || column.type === "numeric") {
      renderChartForColumn(column);
    }
  });

  if (startIndex + 4 < columns.length) {
    window.requestAnimationFrame(() => renderChartsIncrementally(columns, token, startIndex + 4));
  }
}

async function saveCardAsPng(cardElement) {
  try {
    cardElement.classList.add("is-exporting");
    await nextFrame();
    const canvas = await html2canvas(cardElement, {
      backgroundColor: "#ffffff",
      scale: 2,
    });
    const columnId = cardElement.dataset.columnId;
    const column = appState.columns.find((item) => item.id === columnId);
    downloadDataUrl(canvas.toDataURL("image/png"), `${sanitizeFileName(column?.name || "analysis")}.png`);
    showStatus("PNGを保存しました。");
  } catch (error) {
    console.error(error);
    showError("PNG保存に失敗しました。もう一度お試しください。");
  } finally {
    cardElement.classList.remove("is-exporting");
  }
}

async function exportAllAsPdf() {
  if (!appState.columns.some((column) => column.type !== "hidden")) {
    showError("PDFに出力できる分析結果がありません。");
    return;
  }

  try {
    showStatus("PDFを作成中です。列数が多い場合は少し時間がかかります。");
    ui.exportPdfButton.disabled = true;
    ui.exportArea.classList.add("is-exporting");
    await nextFrame();
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF("p", "mm", "a4");
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const margin = 12;
    const contentWidth = pageWidth - margin * 2;
    const contentHeight = pageHeight - margin * 2;
    const gap = 6;
    const layoutColumns = getPdfLayoutColumns();
    let cursorY = margin;

    const summaryCanvas = await renderSummaryCanvas();
    const summaryWidth = contentWidth;
    const summaryHeight = (summaryCanvas.height * summaryWidth) / summaryCanvas.width;
    pdf.addImage(summaryCanvas.toDataURL("image/png"), "PNG", margin, cursorY, summaryWidth, summaryHeight);
    cursorY += summaryHeight + 6;

    const cardElements = Array.from(ui.analysisCards.querySelectorAll(".analysis-card"));
    const renderedCards = [];
    for (let index = 0; index < cardElements.length; index += 1) {
      renderedCards.push(await renderCardForPdf(cardElements[index]));
    }

    if (layoutColumns === 1) {
      for (let index = 0; index < renderedCards.length; index += 1) {
        const card = renderedCards[index];
        const imageWidth = contentWidth;
        let imageHeight = (card.canvas.height * imageWidth) / card.canvas.width;

        if (imageHeight > contentHeight) {
          const ratio = contentHeight / imageHeight;
          imageHeight *= ratio;
        }

        if (cursorY + imageHeight > pageHeight - margin) {
          pdf.addPage();
          cursorY = margin;
        }

        pdf.addImage(card.imageData, "PNG", margin, cursorY, imageWidth, imageHeight);
        cursorY += imageHeight + gap;
      }
    } else {
      const columnWidth = (contentWidth - gap) / 2;

      for (let index = 0; index < renderedCards.length; index += 2) {
        const rowCards = renderedCards.slice(index, index + 2).map((card, rowIndex) => {
          let imageHeight = (card.canvas.height * columnWidth) / card.canvas.width;
          const maxRowHeight = contentHeight * 0.7;
          if (imageHeight > maxRowHeight) {
            const ratio = maxRowHeight / imageHeight;
            imageHeight *= ratio;
          }

          return {
            ...card,
            x: margin + rowIndex * (columnWidth + gap),
            width: columnWidth,
            height: imageHeight,
          };
        });

        const rowHeight = Math.max(...rowCards.map((card) => card.height));
        if (cursorY + rowHeight > pageHeight - margin) {
          pdf.addPage();
          cursorY = margin;
        }

        rowCards.forEach((card) => {
          pdf.addImage(card.imageData, "PNG", card.x, cursorY, card.width, card.height);
        });

        cursorY += rowHeight + gap;
      }
    }

    pdf.save("csv-survey-analysis.pdf");
    showStatus("PDFを出力しました。");
  } catch (error) {
    console.error(error);
    showError("PDF出力に失敗しました。しばらく待ってから再度お試しください。");
  } finally {
    ui.exportArea.classList.remove("is-exporting");
    refreshResultsVisibility();
  }
}

async function renderCardForPdf(cardElement) {
  const canvas = await html2canvas(cardElement, {
    backgroundColor: "#ffffff",
    scale: 2,
    useCORS: true,
  });

  return {
    canvas,
    imageData: canvas.toDataURL("image/png"),
  };
}

function nextFrame() {
  return new Promise((resolve) => {
    window.requestAnimationFrame(() => resolve());
  });
}

function buildPdfSummaryLines() {
  return [
    `総回答数: ${appState.rows.length}`,
    `列数: ${appState.headers.length}`,
    `読み込み方法: ${appState.sourceType || "-"}`,
    `ファイル名: ${appState.sourceName || "-"}`,
    `読込日時: ${appState.loadedAt || "-"}`,
    `PDFレイアウト: ${getPdfLayoutColumns() === 2 ? "2カラム" : "1カラム"}`,
  ];
}

function getPdfLayoutColumns() {
  return ui.pdfLayoutSelect.value === "double" ? 2 : 1;
}

function preprocessCsvText(text) {
  return String(text)
    .replace(/^\uFEFF/, "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .trim();
}

function detectDelimiter(text) {
  const firstLine = text.split("\n").find((line) => line.trim()) || "";
  const candidates = [",", "\t", ";"];
  const best = candidates
    .map((delimiter) => ({ delimiter, count: firstLine.split(delimiter).length }))
    .sort((a, b) => b.count - a.count)[0];
  return best && best.count > 1 ? best.delimiter : ",";
}

function normalizeHeaders(rawHeaders) {
  const seen = new Map();
  return rawHeaders.map((header, index) => {
    const base = normalizeCellValue(String(header ?? "")) || `列${index + 1}`;
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);
    return count === 0 ? base : `${base}_${count + 1}`;
  });
}

function normalizeCellValue(value) {
  return value
    .replace(/\u3000/g, " ")
    .replace(/[０-９．，－]/g, (char) => String.fromCharCode(char.charCodeAt(0) - 65248))
    .trim();
}

function sanitizeNumericValue(value) {
  return String(value).replace(/,/g, "");
}

function buildNumericDistribution(sortedValues) {
  const uniqueValues = Array.from(new Set(sortedValues));
  if (uniqueValues.length <= 12) {
    const counts = uniqueValues.map((value) => sortedValues.filter((item) => item === value).length);
    return {
      labels: uniqueValues.map((value) => formatNumber(value)),
      counts,
    };
  }

  const min = sortedValues[0];
  const max = sortedValues[sortedValues.length - 1];
  const binCount = Math.min(8, Math.max(5, Math.round(Math.sqrt(sortedValues.length))));
  const binSize = (max - min || 1) / binCount;
  const bins = Array.from({ length: binCount }, (_, index) => ({
    start: min + binSize * index,
    end: index === binCount - 1 ? max : min + binSize * (index + 1),
    count: 0,
  }));

  sortedValues.forEach((value) => {
    const index = Math.min(binCount - 1, Math.floor((value - min) / binSize));
    bins[index].count += 1;
  });

  return {
    labels: bins.map((bin) => `${formatNumber(bin.start)}-${formatNumber(bin.end)}`),
    counts: bins.map((bin) => bin.count),
  };
}

function extractFrequentWords(texts) {
  const wordMap = new Map();

  texts.forEach((text) => {
    const tokens = text
      .toLowerCase()
      .split(/[^\p{L}\p{N}]+/u)
      .map((token) => token.trim())
      .filter((token) => token.length >= 2 && !TEXT_STOP_WORDS.has(token));

    const uniqueTokens = new Set(tokens);
    uniqueTokens.forEach((token) => {
      wordMap.set(token, (wordMap.get(token) || 0) + 1);
    });
  });

  return Array.from(wordMap.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5)
    .map(([word, count]) => ({ word, count }));
}

async function renderSummaryCanvas() {
  const summaryElement = document.createElement("div");
  summaryElement.className = "export-summary";
  summaryElement.innerHTML = `
    <h1>CSVアンケート分析ツール</h1>
    <ul>
      ${buildPdfSummaryLines().map((line) => `<li>${escapeHtml(line)}</li>`).join("")}
    </ul>
  `;

  summaryElement.style.position = "fixed";
  summaryElement.style.left = "-9999px";
  summaryElement.style.top = "0";
  document.body.appendChild(summaryElement);

  try {
    return await html2canvas(summaryElement, {
      backgroundColor: "#ffffff",
      scale: 2,
      useCORS: true,
    });
  } finally {
    summaryElement.remove();
  }
}

function destroyAllCharts() {
  chartMap.forEach((chart) => chart.destroy());
  chartMap.clear();
}

function destroyChartForColumn(columnId) {
  const chart = chartMap.get(columnId);
  if (chart) {
    chart.destroy();
    chartMap.delete(columnId);
  }
}

function resetResults() {
  destroyAllCharts();
  appState.sourceType = "";
  appState.sourceName = "";
  appState.loadedAt = "";
  appState.headers = [];
  appState.rows = [];
  appState.columns = [];
  renderOverview();
  ui.analysisCards.innerHTML = "";
  ui.emptyState.style.display = "block";
  ui.exportPdfButton.disabled = true;
}

function showError(message) {
  ui.errorBox.textContent = message;
  ui.errorBox.classList.add("is-visible");
}

function showStatus(message) {
  ui.statusBox.textContent = message;
  ui.statusBox.classList.add("is-visible");
}

function clearMessages() {
  ui.errorBox.textContent = "";
  ui.errorBox.classList.remove("is-visible");
  ui.statusBox.textContent = "";
  ui.statusBox.classList.remove("is-visible");
}

function formatDateTime(date) {
  return new Intl.DateTimeFormat("ja-JP", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

function formatNumber(value) {
  return Number(value).toLocaleString("ja-JP", {
    maximumFractionDigits: 2,
  });
}

function downloadDataUrl(dataUrl, fileName) {
  const link = document.createElement("a");
  link.href = dataUrl;
  link.download = fileName;
  link.click();
}

function sanitizeFileName(value) {
  return String(value).replace(/[\\/:*?"<>|]/g, "_");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function htmlToElement(markup) {
  const template = document.createElement("template");
  template.innerHTML = markup.trim();
  return template.content.firstElementChild;
}
