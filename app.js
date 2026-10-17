(() => {
  "use strict";

  let outputFileName = "Notas na fila.xlsx";
  const SHEET_NAME = "Filtrado";
  const DATE_FORMAT = "dd/mm/yyyy hh:mm:ss";
  const CHUNK_SIZE = 500;
  const HISTORY_SHEET = "Histórico TSKs";
  const HISTORY_HEADER = ["Data da comparação", "TSK", "Movimentação", "Situação", "Resumo anterior", "Arquivo anterior", "Arquivo atual"];
  const { normalizeKey, meaningful, isTotalOutage, outageCellIndexes, mergeSummary, compareActive, needsReview } = window.NotaSyncCore;

  const INITIAL_OUTPUT_HINTS = [
    "Número de Ordem / TSK",
    "Estado",
    "Demais colunas da origem",
    "Status/Resumo / OBS"
  ];

  // Mapa central de compatibilidade. Para adicionar novos formatos no futuro,
  // basta incluir um novo alias no campo correspondente.
  const COLUMN_SCHEMA = [
    {
      key: "order",
      output: "Número de Ordem",
      standard: ["Número de Ordem"],
      aliases: ["TSK"],
      required: true,
      operational: true
    },
    {
      key: "state",
      output: "Estado",
      standard: ["Estado"],
      aliases: [],
      required: true,
      operational: true
    },
    {
      key: "slaEnd",
      output: "Fim SLA",
      standard: ["Fim SLA"],
      aliases: [],
      required: false,
      operational: true
    },
    {
      key: "cm",
      output: "CM",
      standard: ["CM"],
      aliases: [],
      required: false,
      operational: true
    },
    {
      key: "nttCreated",
      output: "Criação do NTT",
      standard: ["Criação do NTT"],
      aliases: [],
      required: false,
      operational: true
    },
    {
      key: "endId",
      output: "END_ID",
      standard: ["END_ID"],
      aliases: [],
      required: false,
      operational: true
    },
    {
      key: "neId",
      output: "NE ID",
      standard: ["NE ID"],
      aliases: [],
      required: false,
      operational: true
    },
    {
      key: "creatorRule",
      output: "Regra usuário criador",
      standard: ["Regra usuário criador"],
      aliases: [],
      required: false,
      operational: true
    },
    {
      key: "failureType",
      output: "Tipo da Falha",
      standard: ["Tipo da Falha"],
      aliases: [],
      required: false,
      operational: true
    },
    {
      key: "alarmTitle",
      output: "Título do Alarme",
      standard: ["Título do Alarme"],
      aliases: [],
      required: false,
      operational: true
    },
    {
      key: "priority",
      output: "Faixa",
      standard: ["Faixa Priorização Dispatching", "Faixa"],
      aliases: ["P"],
      required: false,
      operational: true
    },
    {
      key: "status",
      output: "Status/Resumo",
      standard: ["Status/Resumo"],
      aliases: ["OBS"],
      required: false,
      operational: false
    }
  ];

  const $ = id => document.getElementById(id);
  const dropZone = $("dropZone");
  const fileInput = $("fileInput");
  const convertBtn = $("convertBtn");
  const downloadBtn = $("downloadBtn");
  const fileInfo = $("fileInfo");
  const fileName = $("fileName");
  const fileSize = $("fileSize");
  const previousDropZone = $("previousDropZone");
  const previousFileInput = $("previousFileInput");
  const previousFileInfo = $("previousFileInfo");
  const previousFileName = $("previousFileName");
  const previousFileSize = $("previousFileSize");
  const progressFill = $("progressFill");
  const progressPct = $("progressPct");
  const progressText = $("progressText");
  const logEl = $("log");
  const currentYear = $("currentYear");
  const actionHint = $("actionHint");

  let selectedFile = null;
  let previousFile = null;
  let generatedWorkbook = null;

  currentYear.textContent = String(new Date().getFullYear());

  function renderOutputColumns(columns) {
    const container = $("chips");
    container.innerHTML = "";
    columns.forEach(c => {
      const el = document.createElement("span");
      el.className = "chip";
      el.textContent = c;
      container.appendChild(el);
    });
  }

  renderOutputColumns(INITIAL_OUTPUT_HINTS);

  function setActionHint(message = "") {
    actionHint.innerHTML = message;
  }

  function updateActionState(state) {
    convertBtn.classList.remove("btn-attention");
    downloadBtn.classList.remove("btn-attention", "btn-download-ready");

    if (state === "process") {
      convertBtn.classList.add("btn-attention");
      setActionHint('Arquivo pronto para processamento. <strong>Clique em "Processar planilha"</strong> para continuar.');
      return;
    }

    if (state === "download") {
      downloadBtn.classList.add("btn-attention", "btn-download-ready");
      setActionHint('Processamento concluído. <strong>Clique em "Baixar arquivo XLSX"</strong> para salvar o resultado.');
      return;
    }

    if (state === "processing") {
      setActionHint("Processamento em andamento. Aguarde a conclusão da análise.");
      return;
    }

    setActionHint("");
  }

  updateActionState();

  function log(message, type = "") {
    const div = document.createElement("div");
    if (type) div.className = type;
    div.textContent = `[${new Date().toLocaleTimeString("pt-BR")}] ${message}`;
    if (logEl.firstElementChild && logEl.firstElementChild.classList.contains("muted")) {
      logEl.innerHTML = "";
    }
    logEl.appendChild(div);
    logEl.scrollTop = logEl.scrollHeight;
  }

  function progress(value, text) {
    const v = Math.max(0, Math.min(100, Math.round(value)));
    progressFill.style.width = v + "%";
    progressPct.textContent = v + "%";
    progressText.textContent = text;
  }

  const normalizeText = window.NotaSyncCore.normalizeText;

  function normalizeHeaderForLookup(value) {
    return String(value ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  function yieldToBrowser() {
    return new Promise(resolve => setTimeout(resolve, 0));
  }

  function formatBytes(bytes) {
    if (!Number.isFinite(bytes)) return "";
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 ** 2) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1024 ** 2).toFixed(2)} MB`;
  }

  function parseBrazilDate(value) {
    if (value === null || value === undefined) return null;
    const s = String(value).trim();
    if (!s) return null;

    const m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2}|\d{4})(?:\s+|T)(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (!m) return null;

    let day = Number(m[1]);
    let month = Number(m[2]);
    let year = Number(m[3]);
    const hour = Number(m[4]);
    const minute = Number(m[5]);
    const second = m[6] === undefined ? 0 : Number(m[6]);

    if (year < 100) year = 2000 + year;

    if (
      !Number.isInteger(day) || !Number.isInteger(month) || !Number.isInteger(year) ||
      hour > 23 || minute > 59 || second > 59 ||
      month < 1 || month > 12 || day < 1 || day > 31
    ) return null;

    const date = new Date(year, month - 1, day, hour, minute, second);
    if (
      date.getFullYear() !== year ||
      date.getMonth() !== month - 1 ||
      date.getDate() !== day ||
      date.getHours() !== hour ||
      date.getMinutes() !== minute ||
      date.getSeconds() !== second
    ) return null;

    return date;
  }

  function parseSupportedDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return new Date(value.getTime());
    }
    return parseBrazilDate(value);
  }

  function chooseEncoding(arrayBuffer) {
    return arrayBuffer;
  }

  function getInputTypeLabel(name = "") {
    const lower = String(name).toLowerCase();
    if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) return "XLSX";
    return "CSV/CVS";
  }

  function getOutputFileName(originalName) {
    const name = String(originalName || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();

    if (name.includes("merielem") || name.includes("notas na fila es")) return "Notas na fila ES.xlsx";
    if (name.includes("vinicius") || name.includes("notas na fila bxd")) return "Notas na fila Bxd.xlsx";
    return "Notas na fila.xlsx";
  }

  function buildHeaderIndex(originalHeader) {
    const indexByName = new Map();
    originalHeader.forEach((value, index) => {
      const normalized = normalizeHeaderForLookup(value);
      if (normalized && !indexByName.has(normalized)) indexByName.set(normalized, index);
    });
    return indexByName;
  }

  function resolveColumn(indexByName, definition) {
    for (const name of definition.standard) {
      const index = indexByName.get(normalizeHeaderForLookup(name));
      if (index !== undefined) {
        return { index, sourceName: name, matchType: "standard" };
      }
    }

    for (const name of definition.aliases) {
      const index = indexByName.get(normalizeHeaderForLookup(name));
      if (index !== undefined) {
        return { index, sourceName: name, matchType: "alias" };
      }
    }

    return null;
  }

  function levenshteinDistance(a, b) {
    const s = normalizeText(a);
    const t = normalizeText(b);
    const rows = s.length + 1;
    const cols = t.length + 1;
    const matrix = Array.from({ length: rows }, () => Array(cols).fill(0));

    for (let i = 0; i < rows; i++) matrix[i][0] = i;
    for (let j = 0; j < cols; j++) matrix[0][j] = j;

    for (let i = 1; i < rows; i++) {
      for (let j = 1; j < cols; j++) {
        const cost = s[i - 1] === t[j - 1] ? 0 : 1;
        matrix[i][j] = Math.min(
          matrix[i - 1][j] + 1,
          matrix[i][j - 1] + 1,
          matrix[i - 1][j - 1] + cost
        );
      }
    }

    return matrix[rows - 1][cols - 1];
  }

  function similarityScore(a, b) {
    const x = normalizeText(a);
    const y = normalizeText(b);
    if (!x || !y) return 0;
    if (x === y) return 1;
    if (x.includes(y) || y.includes(x)) return 0.9;

    const maxLen = Math.max(x.length, y.length);
    if (!maxLen) return 1;
    return 1 - (levenshteinDistance(x, y) / maxLen);
  }

  function findPossibleMatch(definition, extras) {
    const candidates = [...definition.standard, ...definition.aliases];
    let best = null;

    for (const extra of extras) {
      for (const expected of candidates) {
        const score = similarityScore(extra.name, expected);
        if (!best || score > best.score) {
          best = { sourceName: extra.name, expected, score };
        }
      }
    }

    return best && best.score >= 0.72 ? best : null;
  }

  function analyzeCurrentColumns(originalHeader) {
    const indexByName = buildHeaderIndex(originalHeader);
    const columns = {};
    const usedIndexes = new Set();
    const aliasesUsed = [];
    const missingRequired = [];
    const missingOptional = [];

    for (const definition of COLUMN_SCHEMA) {
      const resolved = resolveColumn(indexByName, definition);
      columns[definition.key] = resolved;

      if (resolved) {
        usedIndexes.add(resolved.index);
        if (resolved.matchType === "alias") {
          aliasesUsed.push({
            key: definition.key,
            sourceName: originalHeader[resolved.index] || resolved.sourceName,
            output: definition.output
          });
        }
      } else if (definition.required) {
        missingRequired.push(definition);
      } else if (definition.operational) {
        missingOptional.push(definition);
      }
    }

    const extras = originalHeader
      .map((name, index) => ({ name: String(name ?? "").trim(), index }))
      .filter(item => item.name && !usedIndexes.has(item.index));

    const operationalDefinitions = COLUMN_SCHEMA.filter(definition => definition.operational);
    const operationalMatched = operationalDefinitions.filter(definition => columns[definition.key]).length;

    const suggestions = missingRequired
      .map(definition => ({ definition, suggestion: findPossibleMatch(definition, extras) }))
      .filter(item => item.suggestion);

    return {
      columns,
      aliasesUsed,
      missingRequired,
      missingOptional,
      extras,
      suggestions,
      operationalMatched,
      operationalTotal: operationalDefinitions.length,
      alternative: aliasesUsed.length > 0
    };
  }

  function logCurrentColumnDiagnostics(analysis, originalHeader) {
    log("Analisando estrutura da planilha atual...");

    if (analysis.alternative) {
      log("Estrutura alternativa detectada. O NotaSync aplicará o mapa de compatibilidade.", "ok");
    } else {
      log("Estrutura reconhecida no padrão NotaSync/compatível.", "ok");
    }

    for (const alias of analysis.aliasesUsed) {
      log(`✓ "${alias.sourceName}" reconhecida internamente como "${alias.output}".`, "ok");
    }

    for (const item of analysis.suggestions) {
      log(`⚠ Possível correspondência para "${item.definition.output}": "${item.suggestion.sourceName}". Confirme o cabeçalho antes de processar.`, "err");
    }

    if (analysis.missingRequired.length) {
      const requiredNames = analysis.missingRequired.map(item => item.output).join(", ");
      const found = originalHeader.filter(Boolean).join(", ");
      throw new Error(`Estrutura incompatível. Colunas essenciais ausentes: ${requiredNames}. Cabeçalhos encontrados: ${found}`);
    }

    const recognizedCount = COLUMN_SCHEMA.filter(definition => analysis.columns[definition.key]).length;
    const ignoredCount = analysis.extras.length;
    log(`${recognizedCount} campos compatíveis reconhecidos na origem.`, "ok");

    if (ignoredCount > 0) {
      const preview = analysis.extras.slice(0, 8).map(item => `"${item.name}"`).join(", ");
      const suffix = ignoredCount > 8 ? ` e mais ${ignoredCount - 8}` : "";
      log(`ℹ ${ignoredCount} colunas extras serão ignoradas na saída: ${preview}${suffix}.`);
    }

    if (analysis.columns.status) {
      const statusName = originalHeader[analysis.columns.status.index];
      log(`✓ "${statusName}" será utilizada como coluna de resumo/status.`, "ok");
    } else {
      log('ℹ A origem não possui "Status/Resumo" ou "OBS" — somente "Status/Resumo" será adicionada ao final.');
    }

    if (!analysis.columns.priority) {
      log('ℹ Coluna de faixa/prioridade não encontrada — o cruzamento continuará funcionando normalmente.');
    }

    log("Estrutura válida para processamento operacional.", "ok");
  }

  function cellValue(row, analysis, key) {
    const resolved = analysis.columns[key];
    if (!resolved) return "";
    return row[resolved.index] ?? "";
  }

  function buildDynamicOutputLayout(originalHeader, analysis) {
    // Mantém somente os campos reconhecidos pelo NotaSync que realmente existem
    // na planilha atual. Colunas extras da origem são ignoradas.
    // A ordem original dos campos reconhecidos é preservada.
    const selectedColumns = COLUMN_SCHEMA
      .map(definition => {
        const resolved = analysis.columns[definition.key];
        if (!resolved) return null;
        return {
          key: definition.key,
          sourceIndex: resolved.index,
          headerName: originalHeader[resolved.index] || definition.output
        };
      })
      .filter(Boolean)
      .sort((a, b) => a.sourceIndex - b.sourceIndex);

    const header = selectedColumns.map(item => item.headerName);
    const outputIndexByKey = new Map();
    selectedColumns.forEach((item, outputIndex) => {
      item.outputIndex = outputIndex;
      outputIndexByKey.set(item.key, outputIndex);
    });

    let statusIndex = outputIndexByKey.has("status") ? outputIndexByKey.get("status") : -1;
    if (statusIndex < 0) {
      statusIndex = header.length;
      header.push("Status/Resumo");
      outputIndexByKey.set("status", statusIndex);
    }

    return {
      header,
      selectedColumns,
      statusIndex,
      orderSourceIndex: analysis.columns.order.index,
      stateSourceIndex: analysis.columns.state.index,
      orderIndex: outputIndexByKey.get("order"),
      stateIndex: outputIndexByKey.get("state"),
      priorityIndex: outputIndexByKey.has("priority") ? outputIndexByKey.get("priority") : -1,
      slaEndIndex: outputIndexByKey.has("slaEnd") ? outputIndexByKey.get("slaEnd") : -1,
      nttCreatedIndex: outputIndexByKey.has("nttCreated") ? outputIndexByKey.get("nttCreated") : -1
    };
  }

  function calculateColumnWidths(output) {
    if (!output.length) return [];
    const colCount = output[0].length;
    const sampleLimit = Math.min(output.length, 250);
    const widths = [];

    for (let c = 0; c < colCount; c++) {
      let maxLen = String(output[0][c] ?? "").length;
      for (let r = 1; r < sampleLimit; r++) {
        const value = output[r][c];
        const len = value instanceof Date ? 19 : String(value ?? "").length;
        if (len > maxLen) maxLen = len;
      }
      widths.push(Math.max(10, Math.min(40, maxLen + 2)));
    }

    return widths;
  }

  async function loadPreviousSnapshot(file) {
    const result = { statuses: new Map(), history: [], duplicates: 0, conflictKeys: new Set(), skippedStates: 0 };
    if (!file) return result;

    log(`Lendo planilha anterior: ${file.name} (${formatBytes(file.size)}).`);
    const wb = XLSX.read(await file.arrayBuffer(), { type: "array", raw: true, cellDates: true, dense: true });
    if (!wb.SheetNames.length) throw new Error("A planilha anterior não possui abas reconhecíveis.");
    const sheet = wb.Sheets[wb.SheetNames.includes(SHEET_NAME) ? SHEET_NAME : wb.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
    if (!matrix.length) throw new Error("A planilha anterior está vazia.");
    const headers = matrix[0].map(v => String(v ?? ""));
    const index = buildHeaderIndex(headers);
    const order = resolveColumn(index, COLUMN_SCHEMA.find(item => item.key === "order"));
    const status = resolveColumn(index, COLUMN_SCHEMA.find(item => item.key === "status"));
    const state = resolveColumn(index, COLUMN_SCHEMA.find(item => item.key === "state"));
    if (!order) throw new Error('A planilha anterior precisa possuir "Número de Ordem" ou "TSK".');
    if (!state) log('⚠ Planilha anterior sem coluna "Estado": verifique se possui a mesma abrangência da planilha atual.', "err");
    if (!status) log('Planilha anterior sem Status/Resumo; o cruzamento continuará sem reaproveitar resumos.');

    for (let r = 1; r < matrix.length; r++) {
      const row = matrix[r] || [];
      if (state && !normalizeText(row[state.index]).includes("nao iniciado")) {
        result.skippedStates++;
        continue;
      }
      const key = normalizeKey(row[order.index]);
      if (!key) continue;
      const summary = status ? String(row[status.index] ?? "").trim() : "";
      if (result.statuses.has(key)) {
        result.duplicates++;
        const existing = result.statuses.get(key);
        if (meaningful(existing.summary) && meaningful(summary) && existing.summary !== summary) {
          result.conflictKeys.add(key);
        } else if (!meaningful(existing.summary) && meaningful(summary)) {
          existing.summary = summary;
        }
        continue;
      }
      result.statuses.set(key, { summary, displayKey: String(row[order.index] ?? "").trim() });
    }
    if (result.duplicates) log(`⚠ ${result.duplicates} linhas de TSK duplicadas na planilha anterior (${result.conflictKeys.size} com resumo conflitante). Resumos conflitantes não serão importados.`, "err");
    if (result.skippedStates) log(`${result.skippedStates} linhas anteriores fora do Estado "Não iniciado" foram ignoradas na comparação.`);

    if (wb.Sheets[HISTORY_SHEET]) {
      const oldHistory = XLSX.utils.sheet_to_json(wb.Sheets[HISTORY_SHEET], { header: 1, raw: true, defval: "" });
      if (oldHistory.length && normalizeHeaderForLookup(oldHistory[0][1]) === "tsk") {
        result.history = oldHistory.slice(1).filter(row => normalizeKey(row[1])).map(row => HISTORY_HEADER.map((_, c) => row[c] ?? ""));
        log(`${result.history.length} movimentações históricas recuperadas.`, "ok");
      }
    }
    log(`${result.statuses.size} TSKs ativas na planilha anterior.`, "ok");
    return result;
  }

  function resetActivity() {
    $("activityPanel").hidden = true;
    $("scopeWarning").hidden = true;
    $("scopeWarning").textContent = "";
    $("missingList").replaceChildren();
    $("outageList").replaceChildren();
  }

  function renderActivityList(id, items, makeItem) {
    const container = $(id);
    container.replaceChildren();
    if (!items.length) {
      const blank = document.createElement("div");
      blank.className = "activity-empty";
      blank.textContent = "Nenhum registro nesta categoria.";
      container.appendChild(blank);
      return;
    }
    items.slice(0, 15).forEach(item => {
      const [title, detail] = makeItem(item);
      const el = document.createElement("div");
      el.className = "activity-item" + (id === "outageList" ? " outage" : "");
      const heading = document.createElement("strong");
      heading.textContent = title;
      const desc = document.createElement("span");
      desc.textContent = detail;
      el.append(heading, desc);
      container.appendChild(el);
    });
    if (items.length > 15) {
      const rest = document.createElement("div");
      rest.className = "activity-empty";
      rest.textContent = `Mais ${items.length - 15} ocorrência(s) na planilha exportada.`;
      container.appendChild(rest);
    }
  }

  function excelCellValue(value) {
    if (value === null || value === undefined) return "";
    return value;
  }

  function addFormattedSheet(wb, name, rows, options = {}) {
    const ws = wb.addWorksheet(name, { views: [{ state: "frozen", ySplit: 1 }] });
    rows.forEach((sourceRow, rowIndex) => {
      const row = ws.addRow(sourceRow.map(excelCellValue));
      row.eachCell({ includeEmpty: true }, (cell, col) => {
        const isHeader = rowIndex === 0;
        cell.border = {
          top: { style: "thin", color: { argb: "FFB7B7B7" } },
          bottom: { style: "thin", color: { argb: "FFB7B7B7" } },
          left: { style: "thin", color: { argb: "FFB7B7B7" } },
          right: { style: "thin", color: { argb: "FFB7B7B7" } }
        };
        cell.font = { name: "Calibri", size: 11, bold: isHeader, color: { argb: isHeader ? "FF000000" : "FF243449" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: isHeader ? "FFFFF200" : "FFFFFFFF" } };
        cell.alignment = { vertical: "middle", horizontal: isHeader ? "center" : "left", wrapText: isHeader };
        if (options.dateIndexes?.includes(col - 1) && sourceRow[col - 1] instanceof Date) cell.numFmt = DATE_FORMAT;
        if (!isHeader && options.outageCells?.get(rowIndex)?.has(col - 1)) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFE4E6" } };
          cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: "FF9F1239" } };
        }
      });
      if (rowIndex === 0) row.height = 30;
    });
    calculateColumnWidths(rows).forEach((wch, c) => { ws.getColumn(c + 1).width = wch; });
    if (rows.length > 1) ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: rows.length, column: rows[0].length } };
    return ws;
  }

  function resetCounters() {
    $("totalRows").textContent = "0";
    $("filteredRows").textContent = "0";
    $("dateRows").textContent = "0";
    $("currentCount").textContent = "0";
    $("existingCount").textContent = "0";
    $("newCount").textContent = "0";
    $("missingCount").textContent = "0";
    $("outageCount").textContent = "0";
    $("reappearedCount").textContent = "0";
    resetActivity();
    $("countP1").textContent = "0";
    $("countP2").textContent = "0";
    $("countP3").textContent = "0";
    $("countP4").textContent = "0";
    $("countP5").textContent = "0";
    $("countSemFaixa").textContent = "0";
  }

  async function processFile(file) {
    generatedWorkbook = null;
    outputFileName = getOutputFileName(file.name).replace(/\.xlsx$/i, "-SmartSync-1.5.2.xlsx");
    downloadBtn.disabled = true;

    if (!window.XLSX || !window.ExcelJS) {
      throw new Error("Bibliotecas não carregadas (SheetJS / ExcelJS). Conecte-se à internet e recarregue o NotaSync.");
    }
    if (file.size > 75 * 1024 * 1024 || (previousFile && previousFile.size > 75 * 1024 * 1024)) {
      throw new Error("Arquivo acima de 75 MB: tamanho não suportado nesta versão para evitar travamento do navegador.");
    }

    progress(3, "Lendo arquivo...");
    log(`Arquivo recebido: ${file.name} (${formatBytes(file.size)}).`);

    const previousSnapshot = await loadPreviousSnapshot(previousFile);
    const previousStatusMap = previousSnapshot.statuses;
    const buffer = chooseEncoding(await file.arrayBuffer());
    await yieldToBrowser();

    const inputType = getInputTypeLabel(file.name);
    progress(12, previousFile ? `Interpretando ${inputType} e preparando cruzamento...` : `Interpretando ${inputType}...`);

    const workbook = XLSX.read(buffer, {
      type: "array",
      raw: true,
      cellDates: true,
      dense: true
    });

    if (!workbook.SheetNames.length) {
      throw new Error("O arquivo não possui uma planilha/dados reconhecíveis.");
    }

    const sourceSheet = workbook.Sheets[workbook.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json(sourceSheet, {
      header: 1,
      raw: true,
      defval: ""
    });

    if (!matrix.length) throw new Error("O arquivo está vazio.");

    const originalHeader = matrix[0].map(v => String(v ?? "").trim());
    const columnAnalysis = analyzeCurrentColumns(originalHeader);
    logCurrentColumnDiagnostics(columnAnalysis, originalHeader);

    const layout = buildDynamicOutputLayout(originalHeader, columnAnalysis);
    renderOutputColumns(layout.header);

    $("totalRows").textContent = String(Math.max(0, matrix.length - 1));

    const output = [layout.header.slice()];
    let dateCount = 0;
    let existingCount = 0;
    let newCount = 0;
    const priorityCounts = { P1: 0, P2: 0, P3: 0, P4: 0, P5: 0, SEM_FAIXA: 0 };
    const currentKeys = new Set();
    const outageCells = new Map();
    const outageMatches = [];
    const reappeared = [];
    const earlierMissingKeys = new Set(previousSnapshot.history
      .filter(row => String(row[2]).includes("Não localizada"))
      .map(row => normalizeKey(row[1])));
    let currentDuplicates = 0;
    let unsafeNumericKeys = 0;

    const orderSource = originalHeader[layout.orderSourceIndex];
    const stateSource = originalHeader[layout.stateSourceIndex];
    log(`Chave do cruzamento: "${orderSource}".`, "ok");
    log(`Filtro aplicado pela coluna: "${stateSource}".`);
    log('Filtro: Estado contém "Não iniciado" (sem diferenciar maiúsculas/minúsculas e acentos).');

    const dataRows = matrix.length - 1;
    for (let start = 1; start < matrix.length; start += CHUNK_SIZE) {
      const end = Math.min(matrix.length, start + CHUNK_SIZE);

      for (let r = start; r < end; r++) {
        const sourceRow = matrix[r] || [];
        const estado = normalizeText(sourceRow[layout.stateSourceIndex]);
        if (!estado.includes("nao iniciado")) continue;

        // Copia somente as colunas operacionais reconhecidas.
        const values = layout.selectedColumns.map(item => sourceRow[item.sourceIndex] ?? "");
        while (values.length < layout.header.length) values.push("");

        if (typeof values[layout.orderIndex] === "number" &&
            !Number.isSafeInteger(values[layout.orderIndex])) unsafeNumericKeys++;
        const displayOrder = String(values[layout.orderIndex] ?? "").trim();
        const orderKey = normalizeKey(displayOrder);
        if (orderKey) {
          if (currentKeys.has(orderKey)) currentDuplicates++;
          currentKeys.add(orderKey);
        }

        if (previousFile) {
          if (orderKey && previousStatusMap.has(orderKey)) {
            const previous = previousStatusMap.get(orderKey);
            const usablePrevious = !previousSnapshot.conflictKeys.has(orderKey);
            values[layout.statusIndex] = mergeSummary(values[layout.statusIndex], previous.summary, usablePrevious);
            existingCount++;
          } else {
            values[layout.statusIndex] = mergeSummary(values[layout.statusIndex], "", true);
            newCount++;
            if (orderKey && earlierMissingKeys.has(orderKey)) reappeared.push(displayOrder);
          }
        }

        // Destaca somente as células visíveis da saída que contenham a expressão
        // "Indisponibilidade Total". Não colore a TSK por fallback.
        const redIndexes = outageCellIndexes(values);
        const foundFields = new Set(
          [...redIndexes].map(outputIndex => layout.header[outputIndex])
        );
        if (redIndexes.size) {
          outageCells.set(output.length, redIndexes);
          outageMatches.push({ key: displayOrder || `(linha ${r + 1})`, fields: [...foundFields].join(", ") });
        }

        for (const dateIndex of [layout.slaEndIndex, layout.nttCreatedIndex]) {
          if (dateIndex < 0) continue;
          const parsed = parseSupportedDate(values[dateIndex]);
          if (parsed) {
            values[dateIndex] = parsed;
            dateCount++;
          }
        }

        output.push(values);

        if (layout.priorityIndex >= 0) {
          const faixa = String(values[layout.priorityIndex] ?? "").trim().toUpperCase();
          if (/^P[1-5]$/.test(faixa)) priorityCounts[faixa]++;
          else priorityCounts.SEM_FAIXA++;
        } else {
          priorityCounts.SEM_FAIXA++;
        }
      }

      const pct = 20 + ((end - 1) / Math.max(1, dataRows)) * 48;
      progress(pct, `Filtrando linhas... ${end - 1}/${dataRows}`);
      await yieldToBrowser();
    }

    $("filteredRows").textContent = String(output.length - 1);
    $("dateRows").textContent = String(dateCount);
    $("currentCount").textContent = String(output.length - 1);
    $("existingCount").textContent = String(existingCount);
    $("newCount").textContent = String(newCount);
    $("countP1").textContent = String(priorityCounts.P1);
    $("countP2").textContent = String(priorityCounts.P2);
    $("countP3").textContent = String(priorityCounts.P3);
    $("countP4").textContent = String(priorityCounts.P4);
    $("countP5").textContent = String(priorityCounts.P5);
    $("countSemFaixa").textContent = String(priorityCounts.SEM_FAIXA);

    log(`${output.length - 1} linhas filtradas.`);

    if (previousFile) {
      const statusName = layout.header[layout.statusIndex];
      log(`Cruzamento concluído: ${existingCount} já existentes | ${newCount} novas (${statusName} = N/A nas novas).`, "ok");
    } else if (columnAnalysis.columns.status) {
      log(`A coluna "${layout.header[layout.statusIndex]}" da planilha atual foi preservada.`, "ok");
    } else {
      log('A coluna "Status/Resumo" foi adicionada ao final para permitir o tratamento futuro.');
    }

    if (layout.priorityIndex >= 0) {
      log(`Faixas: P1=${priorityCounts.P1} | P2=${priorityCounts.P2} | P3=${priorityCounts.P3} | P4=${priorityCounts.P4} | P5=${priorityCounts.P5} | Sem faixa=${priorityCounts.SEM_FAIXA}.`);
    }
    log(`${dateCount} valores de data reconhecidos e convertidos.`);
    if (currentDuplicates) log(`⚠ ${currentDuplicates} linha(s) de TSK duplicada(s) na planilha atual. Confira antes de usar os dados.`, "err");
    if (unsafeNumericKeys) log(`⚠ ${unsafeNumericKeys} chaves numéricas longas podem ter perdido precisão no Excel de origem. Confira esses identificadores.`, "err");

    const missing = previousFile ? compareActive(previousStatusMap, currentKeys) : [];
    const bulkReview = previousFile && needsReview(missing.length, previousStatusMap.size, currentKeys.size);
    const now = new Date();
    const comparisonTime = now.toLocaleString("pt-BR");
    const history = previousSnapshot.history.slice();
    for (const item of missing) {
      history.push([comparisonTime, item.displayKey || item.key, "Não localizada na consulta atual", "Possível normalização — verificar", item.summary || "", previousFile.name, file.name]);
    }
    for (const key of reappeared) {
      history.push([comparisonTime, key, "Reapareceu na consulta atual", "Ativa novamente", "", previousFile.name, file.name]);
    }
    $("missingCount").textContent = String(missing.length);
    $("outageCount").textContent = String(outageMatches.length);
    $("reappearedCount").textContent = String(reappeared.length);
    $("activityPanel").hidden = false;
    renderActivityList("missingList", missing, item => [item.displayKey || item.key, "Não localizada — possível normalização, verificar."]);
    renderActivityList("outageList", outageMatches, item => [item.key, `Indisponibilidade Total • ${item.fields}`]);
    if (bulkReview) {
      const alertBox = $("scopeWarning");
      alertBox.hidden = false;
      alertBox.textContent = `Atenção: ${missing.length} de ${previousStatusMap.size} TSKs anteriores não aparecem na lista atual. Confira se os arquivos têm a mesma abrangência, período e filtro. Nenhuma foi confirmada como encerrada automaticamente.`;
      log("⚠ Grande variação de TSKs. Confirme filtros e abrangência antes de considerar as ausências como normalização.", "err");
    }
    if (missing.length) log(`${missing.length} TSK(s) não localizada(s). Histórico atualizado; a lista ativa contém apenas registros atuais.`, "ok");
    if (reappeared.length) log(`${reappeared.length} TSK(s) reapareceram após registro anterior de ausência.`, "ok");
    if (outageMatches.length) log(`${outageMatches.length} TSK(s) com "Indisponibilidade Total" identificada(s) e destacada(s) em vermelho.`, "err");

    progress(72, "Criando planilha...");
    await yieldToBrowser();

    const wb = new ExcelJS.Workbook();
    wb.creator = "OrbyttaCore";
    wb.lastModifiedBy = "NotaSync SmartSync";
    wb.created = new Date();
    wb.title = "NotaSync - Notas na fila";
    wb.subject = 'Linhas com Estado "Não iniciado" e movimentações de TSK';
    addFormattedSheet(wb, SHEET_NAME, output, {
      outageCells,
      dateIndexes: [layout.slaEndIndex, layout.nttCreatedIndex].filter(c => c >= 0)
    });
    // A aba de histórico nunca é tratada como lista ativa no próximo cruzamento.
    if (previousFile || history.length) {
      addFormattedSheet(wb, HISTORY_SHEET, [HISTORY_HEADER, ...history]);
    }

    progress(88, "Validando resultados...");
    await yieldToBrowser();
    const dateTests = [
      ["02/09/26 14:37", 2026, 8, 2, 14, 37, 0],
      ["01/09/2026 10:37:33", 2026, 8, 1, 10, 37, 33]
    ];
    for (const [text, y, mo, d, h, mi, se] of dateTests) {
      const testDate = parseBrazilDate(text);
      if (!testDate || testDate.getFullYear() !== y || testDate.getMonth() !== mo ||
          testDate.getDate() !== d || testDate.getHours() !== h ||
          testDate.getMinutes() !== mi || testDate.getSeconds() !== se) {
        throw new Error(`Falha ao validar parser de datas: ${text}`);
      }
    }

    generatedWorkbook = wb;
    progress(100, "Concluído.");
    log(`Aba criada: "${SHEET_NAME}".`, "ok");
    log(`Arquivo pronto: ${outputFileName}.`, "ok");
    log(`Saída operacional gerada com ${layout.header.length} colunas reconhecidas.`, "ok");
    if (previousFile) log(`Aba "${HISTORY_SHEET}" com ${history.length} evento(s).`, "ok");

    downloadBtn.disabled = false;
    updateActionState("download");
  }

  dropZone.addEventListener("click", () => fileInput.click());
  dropZone.addEventListener("keydown", e => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      fileInput.click();
    }
  });

  fileInput.addEventListener("change", e => {
    if (e.target.files && e.target.files[0]) selectFile(e.target.files[0]);
  });

  ["dragenter", "dragover"].forEach(eventName => {
    dropZone.addEventListener(eventName, e => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add("over");
    });
  });

  ["dragleave", "drop"].forEach(eventName => {
    dropZone.addEventListener(eventName, e => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove("over");
    });
  });

  dropZone.addEventListener("drop", e => {
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if (file) selectFile(file);
  });

  function selectFile(file) {
    const name = file.name.toLowerCase();
    const supported = name.endsWith(".csv") || name.endsWith(".cvs") || name.endsWith(".xlsx") || name.endsWith(".xls");

    if (!supported) {
      alert("Selecione um arquivo .csv, .cvs, .xlsx ou .xls.");
      return;
    }

    selectedFile = file;
    generatedWorkbook = null;
    fileName.textContent = file.name;
    fileSize.textContent = formatBytes(file.size);
    fileInfo.classList.add("show");
    convertBtn.disabled = false;
    downloadBtn.disabled = true;
    progress(0, "Arquivo selecionado. Pronto para processar.");
    log(`Selecionado: ${file.name}.`);
    updateActionState("process");
  }

  previousDropZone.addEventListener("click", () => previousFileInput.click());
  previousDropZone.addEventListener("keydown", e => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      previousFileInput.click();
    }
  });

  previousFileInput.addEventListener("change", e => {
    if (e.target.files && e.target.files[0]) selectPreviousFile(e.target.files[0]);
  });

  ["dragenter", "dragover"].forEach(eventName => {
    previousDropZone.addEventListener(eventName, e => {
      e.preventDefault();
      e.stopPropagation();
      previousDropZone.classList.add("over");
    });
  });

  ["dragleave", "drop"].forEach(eventName => {
    previousDropZone.addEventListener(eventName, e => {
      e.preventDefault();
      e.stopPropagation();
      previousDropZone.classList.remove("over");
    });
  });

  previousDropZone.addEventListener("drop", e => {
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if (file) selectPreviousFile(file);
  });

  function selectPreviousFile(file) {
    const name = file.name.toLowerCase();
    if (!(name.endsWith(".xlsx") || name.endsWith(".xls"))) {
      alert("Selecione uma planilha .xlsx ou .xls.");
      return;
    }

    previousFile = file;
    generatedWorkbook = null;
    previousFileName.textContent = file.name;
    previousFileSize.textContent = formatBytes(file.size);
    previousFileInfo.classList.add("show");
    downloadBtn.disabled = true;
    $("existingCount").textContent = "0";
    $("newCount").textContent = "0";
    $("missingCount").textContent = "0";
    $("outageCount").textContent = "0";
    $("reappearedCount").textContent = "0";
    resetActivity();
    log(`Planilha anterior selecionada: ${file.name}.`);
    if (selectedFile) updateActionState("process");
  }

  convertBtn.addEventListener("click", async () => {
    if (!selectedFile) return;

    convertBtn.disabled = true;
    downloadBtn.disabled = true;
    updateActionState("processing");
    logEl.innerHTML = "";
    resetCounters();

    try {
      await processFile(selectedFile);
    } catch (error) {
      progress(0, "Erro no processamento.");
      log(error?.message || String(error), "err");
      console.error(error);
    } finally {
      convertBtn.disabled = false;
      if (!generatedWorkbook && selectedFile) {
        updateActionState("process");
      }
    }
  });

  downloadBtn.addEventListener("click", async () => {
    if (!generatedWorkbook || downloadBtn.disabled) return;
    downloadBtn.disabled = true;
    try {
      progress(95, "Preparando download do Excel...");
      const buffer = await generatedWorkbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = outputFileName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 60000);
      progress(100, "Concluído. Arquivo preparado para download.");
      log(`Download iniciado: ${outputFileName}.`, "ok");
      downloadBtn.classList.remove("btn-attention");
      setActionHint("XLSX preparado com destaque vermelho e histórico. Guarde o arquivo para o próximo cruzamento.");
    } catch (error) {
      log("Erro ao gerar o XLSX: " + (error?.message || error), "err");
      progress(0, "Erro ao preparar o download.");
    } finally {
      downloadBtn.disabled = false;
    }
  });
})();
