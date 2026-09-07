(() => {
  "use strict";

  let outputFileName = "Notas na fila.xlsx";
  const SHEET_NAME = "Filtrado";
  const DATE_FORMAT = "dd/mm/yyyy hh:mm:ss";
  const CHUNK_SIZE = 500;

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

  function normalizeText(value) {
    return String(value ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[_-]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

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
      log("Estrutura alternativa detectada. O NotaSync aplicará o mapa de compatibilidade sem alterar as colunas da origem.", "ok");
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

    log(`${originalHeader.length} colunas da planilha de origem serão preservadas na saída.`);

    if (analysis.columns.status) {
      const statusName = originalHeader[analysis.columns.status.index];
      log(`✓ "${statusName}" será utilizada como coluna de resumo/status.`, "ok");
    } else {
      log('ℹ A origem não possui "Status/Resumo" ou "OBS" — somente essa coluna será adicionada ao final da saída.');
    }

    if (!analysis.columns.priority) {
      log('ℹ Coluna de faixa/prioridade não encontrada — o cruzamento continuará funcionando normalmente.');
    }

    log("Estrutura válida para processamento flexível.", "ok");
  }

  function cellValue(row, analysis, key) {
    const resolved = analysis.columns[key];
    if (!resolved) return "";
    return row[resolved.index] ?? "";
  }

  function buildDynamicOutputLayout(originalHeader, analysis) {
    const header = originalHeader.slice();
    let statusIndex = analysis.columns.status ? analysis.columns.status.index : -1;

    if (statusIndex < 0) {
      header.push("Status/Resumo");
      statusIndex = header.length - 1;
    }

    return {
      header,
      statusIndex,
      orderIndex: analysis.columns.order.index,
      stateIndex: analysis.columns.state.index,
      priorityIndex: analysis.columns.priority ? analysis.columns.priority.index : -1,
      slaEndIndex: analysis.columns.slaEnd ? analysis.columns.slaEnd.index : -1,
      nttCreatedIndex: analysis.columns.nttCreated ? analysis.columns.nttCreated.index : -1
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

  async function loadPreviousStatusMap(file) {
    const statusMap = new Map();
    if (!file) return statusMap;

    log(`Lendo planilha anterior: ${file.name} (${formatBytes(file.size)}).`);
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array", raw: true, cellDates: true, dense: true });
    if (!wb.SheetNames.length) throw new Error("A planilha anterior não possui abas reconhecíveis.");

    const preferredName = wb.SheetNames.includes(SHEET_NAME) ? SHEET_NAME : wb.SheetNames[0];
    const ws = wb.Sheets[preferredName];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
    if (!matrix.length) throw new Error("A planilha anterior está vazia.");

    const originalHeader = matrix[0].map(v => String(v ?? ""));
    const indexByName = buildHeaderIndex(originalHeader);
    const orderDefinition = COLUMN_SCHEMA.find(item => item.key === "order");
    const statusDefinition = COLUMN_SCHEMA.find(item => item.key === "status");
    const order = resolveColumn(indexByName, orderDefinition);
    const status = resolveColumn(indexByName, statusDefinition);

    if (!order) {
      throw new Error('A planilha anterior precisa possuir "Número de Ordem" ou "TSK".');
    }

    if (order.matchType === "alias") {
      log(`✓ Planilha anterior: "${originalHeader[order.index]}" reconhecida como "Número de Ordem".`, "ok");
    }

    if (status && status.matchType === "alias") {
      log(`✓ Planilha anterior: "${originalHeader[status.index]}" reconhecida como "Status/Resumo".`, "ok");
    }

    if (!status) {
      log('ℹ Planilha anterior sem "Status/Resumo" ou "OBS". O cruzamento identificará notas existentes, mas o resumo anterior ficará vazio.');
    }

    for (let r = 1; r < matrix.length; r++) {
      const row = matrix[r] || [];
      const orderKey = String(row[order.index] ?? "").trim();
      if (!orderKey || statusMap.has(orderKey)) continue;
      const statusValue = status ? String(row[status.index] ?? "").trim() : "";
      statusMap.set(orderKey, statusValue);
    }

    log(`${statusMap.size} números de ordem carregados da planilha anterior.`, "ok");
    return statusMap;
  }

  function resetCounters() {
    $("totalRows").textContent = "0";
    $("filteredRows").textContent = "0";
    $("dateRows").textContent = "0";
    $("currentCount").textContent = "0";
    $("existingCount").textContent = "0";
    $("newCount").textContent = "0";
    $("countP1").textContent = "0";
    $("countP2").textContent = "0";
    $("countP3").textContent = "0";
    $("countP4").textContent = "0";
    $("countP5").textContent = "0";
    $("countSemFaixa").textContent = "0";
  }

  async function processFile(file) {
    generatedWorkbook = null;
    outputFileName = getOutputFileName(file.name);
    downloadBtn.disabled = true;

    if (!window.XLSX) {
      throw new Error("SheetJS não foi carregado. Para uso totalmente offline, a biblioteca precisa estar embutida no próprio HTML.");
    }

    progress(3, "Lendo arquivo...");
    log(`Arquivo recebido: ${file.name} (${formatBytes(file.size)}).`);

    const previousStatusMap = await loadPreviousStatusMap(previousFile);
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

    const orderSource = originalHeader[layout.orderIndex];
    const stateSource = originalHeader[layout.stateIndex];
    log(`Chave do cruzamento: "${orderSource}".`, "ok");
    log(`Filtro aplicado pela coluna: "${stateSource}".`);
    log('Filtro: Estado contém "Não iniciado" (sem diferenciar maiúsculas/minúsculas e acentos).');

    const dataRows = matrix.length - 1;
    for (let start = 1; start < matrix.length; start += CHUNK_SIZE) {
      const end = Math.min(matrix.length, start + CHUNK_SIZE);

      for (let r = start; r < end; r++) {
        const sourceRow = matrix[r] || [];
        const estado = normalizeText(sourceRow[layout.stateIndex]);
        if (!estado.includes("nao iniciado")) continue;

        // Preserva todas as colunas existentes na planilha atual.
        const values = originalHeader.map((_, index) => sourceRow[index] ?? "");
        while (values.length < layout.header.length) values.push("");

        const orderKey = String(values[layout.orderIndex] ?? "").trim();

        if (previousFile) {
          if (orderKey && previousStatusMap.has(orderKey)) {
            values[layout.statusIndex] = previousStatusMap.get(orderKey) || "";
            existingCount++;
          } else {
            values[layout.statusIndex] = "N/A";
            newCount++;
          }
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

    progress(72, "Criando planilha...");
    await yieldToBrowser();

    const ws = XLSX.utils.aoa_to_sheet(output, { cellDates: true });

    const dateIndexes = [layout.slaEndIndex, layout.nttCreatedIndex].filter(index => index >= 0);
    for (let r = 1; r < output.length; r++) {
      for (const c of dateIndexes) {
        const value = output[r][c];
        if (value instanceof Date && !Number.isNaN(value.getTime())) {
          const cellRef = XLSX.utils.encode_cell({ r, c });
          ws[cellRef] = { v: value, t: "d", z: DATE_FORMAT };
        }
      }
    }

    ws["!cols"] = calculateColumnWidths(output).map(wch => ({ wch }));

    const lastRow = Math.max(1, output.length);
    const lastCol = layout.header.length - 1;
    ws["!ref"] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: lastRow - 1, c: lastCol }
    });
    ws["!autofilter"] = {
      ref: `A1:${XLSX.utils.encode_col(lastCol)}${lastRow}`
    };

    const thinBorder = {
      top: { style: "thin", color: { rgb: "B7B7B7" } },
      bottom: { style: "thin", color: { rgb: "B7B7B7" } },
      left: { style: "thin", color: { rgb: "B7B7B7" } },
      right: { style: "thin", color: { rgb: "B7B7B7" } }
    };

    for (let r = 0; r < output.length; r++) {
      for (let c = 0; c < layout.header.length; c++) {
        const ref = XLSX.utils.encode_cell({ r, c });
        if (!ws[ref]) ws[ref] = { v: "", t: "s" };
        const isHeader = r === 0;

        ws[ref].s = {
          border: thinBorder,
          alignment: isHeader
            ? { vertical: "center", horizontal: "center", wrapText: true }
            : { vertical: "center", horizontal: "left", wrapText: false },
          font: isHeader
            ? { bold: true, color: { rgb: "000000" } }
            : { color: { rgb: "000000" } },
          fill: isHeader
            ? { patternType: "solid", fgColor: { rgb: "FFF200" } }
            : { patternType: "solid", fgColor: { rgb: "FFFFFF" } }
        };
      }
    }

    ws["!rows"] = [{ hpt: 30 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME);
    wb.Props = {
      Title: "Filtrado - Não iniciado",
      Subject: 'Linhas com Estado "Não iniciado"',
      Author: "Fábio Gonçalves",
      CreatedDate: new Date()
    };

    progress(88, "Validando células de data...");
    await yieldToBrowser();

    const dateTests = [
      ["02/09/26 14:37", 2026, 8, 2, 14, 37, 0],
      ["01/09/2026 10:37:33", 2026, 8, 1, 10, 37, 33]
    ];

    for (const [text, y, mo, d, h, mi, se] of dateTests) {
      const testDate = parseBrazilDate(text);
      if (!testDate ||
          testDate.getFullYear() !== y ||
          testDate.getMonth() !== mo ||
          testDate.getDate() !== d ||
          testDate.getHours() !== h ||
          testDate.getMinutes() !== mi ||
          testDate.getSeconds() !== se) {
        throw new Error(`Falha na validação do parser de datas: ${text}`);
      }
    }

    let realDateCells = 0;
    for (let r = 1; r < output.length; r++) {
      for (const c of dateIndexes) {
        const cellRef = XLSX.utils.encode_cell({ r, c });
        const cell = ws[cellRef];
        if (cell && cell.t === "d" && cell.z === DATE_FORMAT && cell.v instanceof Date) {
          realDateCells++;
        }
      }
    }

    if (dateCount > 0 && realDateCells !== dateCount) {
      throw new Error(`Falha na validação das células de data: ${realDateCells}/${dateCount}`);
    }

    generatedWorkbook = wb;
    progress(100, "Concluído.");
    log(`Aba criada: "${SHEET_NAME}".`, "ok");
    log(`Arquivo pronto: ${outputFileName}.`, "ok");
    log(`Estrutura preservada: ${layout.header.length} colunas na saída.`, "ok");

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

  downloadBtn.addEventListener("click", () => {
    if (!generatedWorkbook) return;

    try {
      XLSX.writeFile(generatedWorkbook, outputFileName, {
        bookType: "xlsx",
        cellDates: true
      });
      log(`Download iniciado: ${outputFileName}.`, "ok");
      downloadBtn.classList.remove("btn-attention");
      setActionHint("Arquivo gerado com sucesso. Se necessário, você pode processar um novo arquivo a qualquer momento.");
    } catch (error) {
      log("Erro ao gerar o XLSX: " + (error?.message || error), "err");
    }
  });
})();
