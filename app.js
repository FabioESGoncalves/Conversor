(() => {
  "use strict";

  let outputFileName = "Notas na fila.xlsx";
  const SHEET_NAME = "Filtrado";
  const DATE_FORMAT = "dd/mm/yyyy hh:mm:ss";
  const CHUNK_SIZE = 500;

  const OUTPUT_COLUMNS = [
    "Número de Ordem",
    "Estado",
    "Fim SLA",
    "CM",
    "Criação do NTT",
    "END_ID",
    "NE ID",
    "Regra usuário criador",
    "Tipo da Falha",
    "Título do Alarme",
    "Faixa",
    "Status/Resumo"
  ];

  const SOURCE_COLUMNS = [
    "Número de Ordem",
    "Estado",
    "Fim SLA",
    "CM",
    "Criação do NTT",
    "END_ID",
    "NE ID",
    "Regra usuário criador",
    "Tipo da Falha",
    "Título do Alarme",
    "Faixa Priorização Dispatching"
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

  OUTPUT_COLUMNS.forEach(c => {
    const el = document.createElement("span");
    el.className = "chip";
    el.textContent = c;
    $("chips").appendChild(el);
  });

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
      setActionHint('Processamento em andamento. Aguarde a conclusão da análise.');
      return;
    }

    setActionHint("");
  }

  updateActionState();

  function log(message, type = "") {
    const div = document.createElement("div");
    if (type) div.className = type;
    div.textContent = `[${new Date().toLocaleTimeString("pt-BR")}] ${message}`;
    if (logEl.firstElementChild && logEl.firstElementChild.classList.contains("muted")) logEl.innerHTML = "";
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

  function chooseEncoding(arrayBuffer) {
    return arrayBuffer;
  }

  function getOutputFileName(originalName) {
    const name = String(originalName || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
    if (name.includes("merielem")) return "Notas na fila ES.xlsx";
    if (name.includes("vinicius")) return "Notas na fila Bxd.xlsx";
    return "Notas na fila.xlsx";
  }

  async function loadPreviousStatusMap(file) {
    const statusMap = new Map();
    if (!file) return statusMap;

    log(`Lendo planilha anterior: ${file.name} (${formatBytes(file.size)}).`);
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array", raw: true, cellDates: false, dense: true });
    if (!wb.SheetNames.length) throw new Error("A planilha anterior não possui abas reconhecíveis.");

    const preferredName = wb.SheetNames.includes(SHEET_NAME) ? SHEET_NAME : wb.SheetNames[0];
    const ws = wb.Sheets[preferredName];
    const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
    if (!matrix.length) throw new Error("A planilha anterior está vazia.");

    const header = matrix[0].map(v => normalizeText(v));
    const orderIndex = header.indexOf(normalizeText("Número de Ordem"));
    const statusIndex = header.indexOf(normalizeText("Status/Resumo"));

    if (orderIndex < 0) throw new Error('A planilha anterior não possui a coluna "Número de Ordem".');
    if (statusIndex < 0) throw new Error('A planilha anterior não possui a coluna "Status/Resumo". Gere primeiro uma planilha nesta nova versão.');

    for (let r = 1; r < matrix.length; r++) {
      const row = matrix[r] || [];
      const orderKey = String(row[orderIndex] ?? "").trim();
      if (!orderKey || statusMap.has(orderKey)) continue;
      statusMap.set(orderKey, String(row[statusIndex] ?? "").trim());
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

    progress(12, previousFile ? "Interpretando CSV/CVS e preparando cruzamento..." : "Interpretando CSV/CVS com SheetJS...");
    const workbook = XLSX.read(buffer, {
      type: "array",
      raw: true,
      cellDates: false,
      dense: true
    });

    if (!workbook.SheetNames.length) throw new Error("O arquivo não possui uma planilha/dados reconhecíveis.");

    const sourceSheet = workbook.Sheets[workbook.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json(sourceSheet, {
      header: 1,
      raw: true,
      defval: ""
    });

    if (!matrix.length) throw new Error("O arquivo está vazio.");

    const originalHeader = matrix[0].map(v => String(v ?? ""));
    const normalizedHeader = originalHeader.map(normalizeText);

    const indexByName = new Map();
    normalizedHeader.forEach((name, i) => {
      if (name && !indexByName.has(name)) indexByName.set(name, i);
    });

    const requiredSource = SOURCE_COLUMNS.map(name => ({
      name,
      index: indexByName.get(normalizeText(name))
    }));

    const missing = requiredSource.filter(x => x.index === undefined && x.name !== "Faixa Priorização Dispatching");
    if (missing.length) {
      throw new Error("Colunas obrigatórias ausentes: " + missing.map(x => x.name).join(", "));
    }

    const faixaIndex = indexByName.get(normalizeText("Faixa Priorização Dispatching"));
    if (faixaIndex === undefined) {
      throw new Error('A coluna "Faixa Priorização Dispatching" não foi encontrada.');
    }

    const estadoIndex = indexByName.get(normalizeText("Estado"));
    const fimSlaIndex = indexByName.get(normalizeText("Fim SLA"));
    const criacaoIndex = indexByName.get(normalizeText("Criação do NTT"));

    $("totalRows").textContent = String(Math.max(0, matrix.length - 1));

    const output = [OUTPUT_COLUMNS.slice()];
    let dateCount = 0;
    let existingCount = 0;
    let newCount = 0;
    const priorityCounts = { P1: 0, P2: 0, P3: 0, P4: 0, P5: 0, SEM_FAIXA: 0 };

    log(`Cabeçalho encontrado com ${originalHeader.length} colunas.`);
    log('Filtro: coluna Estado contém "Não iniciado" (sem diferenciar maiúsculas/minúsculas e acentos).');

    const dataRows = matrix.length - 1;
    for (let start = 1; start < matrix.length; start += CHUNK_SIZE) {
      const end = Math.min(matrix.length, start + CHUNK_SIZE);

      for (let r = start; r < end; r++) {
        const row = matrix[r] || [];
        const estado = normalizeText(row[estadoIndex]);
        if (!estado.includes("nao iniciado")) continue;

        const values = [
          row[indexByName.get(normalizeText("Número de Ordem"))] ?? "",
          row[estadoIndex] ?? "",
          row[fimSlaIndex] ?? "",
          row[indexByName.get(normalizeText("CM"))] ?? "",
          row[criacaoIndex] ?? "",
          row[indexByName.get(normalizeText("END_ID"))] ?? "",
          row[indexByName.get(normalizeText("NE ID"))] ?? "",
          row[indexByName.get(normalizeText("Regra usuário criador"))] ?? "",
          row[indexByName.get(normalizeText("Tipo da Falha"))] ?? "",
          row[indexByName.get(normalizeText("Título do Alarme"))] ?? "",
          row[faixaIndex] ?? "",
          ""
        ];

        const orderKey = String(values[0] ?? "").trim();
        if (previousFile) {
          if (orderKey && previousStatusMap.has(orderKey)) {
            const previousStatus = previousStatusMap.get(orderKey);
            values[11] = previousStatus || "";
            existingCount++;
          } else {
            values[11] = "N/A";
            newCount++;
          }
        }

        const fim = parseBrazilDate(values[2]);
        if (fim) { values[2] = fim; dateCount++; }

        const criacao = parseBrazilDate(values[4]);
        if (criacao) { values[4] = criacao; dateCount++; }

        output.push(values);

        const faixa = String(values[10] ?? "").trim().toUpperCase();
        if (/^P[1-5]$/.test(faixa)) priorityCounts[faixa]++;
        else priorityCounts.SEM_FAIXA++;
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
    if (previousFile) log(`Cruzamento concluído: ${existingCount} já existentes | ${newCount} novas (Status/Resumo = N/A).`, "ok");
    else log("Sem planilha anterior: Status/Resumo será exportado em branco.");
    log(`Faixas: P1=${priorityCounts.P1} | P2=${priorityCounts.P2} | P3=${priorityCounts.P3} | P4=${priorityCounts.P4} | P5=${priorityCounts.P5} | Sem faixa=${priorityCounts.SEM_FAIXA}.`);
    log(`${dateCount} valores de data reconhecidos e convertidos.`);

    progress(72, "Criando planilha...");
    await yieldToBrowser();

    const ws = XLSX.utils.aoa_to_sheet(output, { cellDates: true });

    for (let r = 1; r < output.length; r++) {
      for (const c of [2, 4]) {
        const value = output[r][c];
        if (value instanceof Date && !Number.isNaN(value.getTime())) {
          const cellRef = XLSX.utils.encode_cell({ r, c });
          ws[cellRef] = { v: value, t: "d", z: DATE_FORMAT };
        }
      }
    }

    const widths = [22, 16, 21, 24, 22, 18, 18, 32, 24, 38, 12, 34];
    ws["!cols"] = widths.map(wch => ({ wch }));

    const lastRow = Math.max(1, output.length);
    const lastCol = OUTPUT_COLUMNS.length - 1;
    ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: lastRow - 1, c: lastCol } });
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
      for (let c = 0; c < OUTPUT_COLUMNS.length; c++) {
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
      for (const c of [2, 4]) {
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
    log(`Nome definido automaticamente pelo arquivo de origem: ${outputFileName}.`, "ok");
    log('Teste "02/09/26 14:37" → 02/09/2026 14:37:00 OK.', "ok");

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
    if (!(name.endsWith(".csv") || name.endsWith(".cvs"))) {
      alert("Selecione um arquivo .csv ou .cvs.");
      return;
    }
    selectedFile = file;
    generatedWorkbook = null;
    fileName.textContent = file.name;
    fileSize.textContent = formatBytes(file.size);
    fileInfo.classList.add("show");
    convertBtn.disabled = false;
    downloadBtn.disabled = true;
    progress(0, "Arquivo selecionado. Pronto para converter.");
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
      setActionHint('Arquivo gerado com sucesso. Se necessário, você pode processar um novo arquivo a qualquer momento.');
    } catch (error) {
      log("Erro ao gerar o XLSX: " + (error?.message || error), "err");
    }
  });
})();
