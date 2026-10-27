// Teste isolado do formatador ExcelJS usado no app, com Workbook simulado.
// Confere a criação dos objetos de hyperlink, sem precisar de rede/CDNs.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { serviceNowTaskUrl } = require('../smartsync-core.js');

const source = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const start = source.indexOf('  function addFormattedSheet(');
const end = source.indexOf('\n  function renderLinkPreview(', start);
assert.ok(start >= 0 && end > start, 'função addFormattedSheet não encontrada');
const addFormattedSheet = vm.runInNewContext(
  `(${source.slice(start, end).trim()})`,
  { serviceNowTaskUrl, excelCellValue: v => v ?? '', calculateColumnWidths: rows => rows[0].map(() => 20), DATE_FORMAT: 'dd/mm/yyyy hh:mm:ss', Date }
);

function mockWorkbook() {
  const sheets = [];
  return {
    sheets,
    addWorksheet(name) {
      const ws = {
        name,
        rows: [],
        addRow(vals) {
          const cells = vals.map(v => ({ value: v }));
          const row = {
            height: 0,
            cells,
            eachCell(_opt, cb) { cells.forEach((c, i) => cb(c, i + 1)); }
          };
          this.rows.push(row);
          return row;
        },
        getColumn() { return { set width(v) { void v; } }; }
      };
      sheets.push(ws);
      return ws;
    }
  };
}

test('exportador formata TSK ativa como hyperlink e não transforma demais células', () => {
  const wb = mockWorkbook();
  const ws = addFormattedSheet(wb, 'Filtrado', [
    ['Número de Ordem', 'Estado', 'Resumo'],
    ['TSK261000114196', 'Não iniciado', 'Tudo normal'],
    ['OUTRO123', 'Não iniciado', 'Indisponibilidade Total']
  ], { taskColumnIndex: 0, outageCells: new Map([[2, new Set([2])]]) });
  const row1 = ws.rows[1].cells;
  assert.equal(row1[0].value.text, 'TSK261000114196');
  assert.equal(row1[0].value.hyperlink, serviceNowTaskUrl('TSK261000114196'));
  assert.equal(row1[0].font.underline, false);
  assert.equal(row1[0].font.color.argb, 'FF243449');
  assert.equal(row1[0].font.name, 'Calibri');
  assert.equal(row1[0].font.size, 11);
  assert.equal(row1[1].value, 'Não iniciado');
  assert.equal(ws.rows[2].cells[0].value, 'OUTRO123');
  assert.equal(ws.rows[2].cells[0].font.underline, undefined);
  assert.equal(ws.rows[2].cells[2].fill.fgColor.argb, 'FFFFE4E6');
  assert.equal(ws.rows[2].cells[0].fill.fgColor.argb, 'FFFFFFFF');
});

test('exportador insere hyperlink na coluna TSK do histórico', () => {
  const wb = mockWorkbook();
  const ws = addFormattedSheet(wb, 'Histórico TSKs', [
    ['Data', 'TSK', 'Movimentação'],
    ['23/10/2026', 'TSK261000114196', 'Possível normalização']
  ], { taskColumnIndex: 1 });
  assert.equal(ws.rows[1].cells[1].value.hyperlink, serviceNowTaskUrl('TSK261000114196'));
  assert.equal(ws.rows[1].cells[1].font.underline, false);
  assert.equal(ws.rows[1].cells[0].value, '23/10/2026');
  assert.equal(ws.rows[1].cells[2].value, 'Possível normalização');
});


test('importação CSV: TSK em outra coluna ganha hyperlink no XLSX, sem depender do índice principal', () => {
  const wb = mockWorkbook();
  const ws = addFormattedSheet(wb, 'Filtrado', [
    ['Número de Ordem', 'TSK', 'Estado', 'Status/Resumo'],
    ['24195', 'TSK261000114196', 'Não iniciado', 'Técnico acionado'],
    ['24196', 'TSK261000114197', 'Não iniciado', 'Indisponibilidade Total']
  ], { taskColumnIndex: 0, outageCells: new Map([[2,new Set([3])]]) });
  assert.equal(ws.rows[1].cells[0].value, '24195');
  assert.equal(ws.rows[1].cells[1].value.hyperlink, serviceNowTaskUrl('TSK261000114196'));
  assert.equal(ws.rows[2].cells[1].value.hyperlink, serviceNowTaskUrl('TSK261000114197'));
  assert.equal(ws.rows[2].cells[3].fill.fgColor.argb, 'FFFFE4E6');
  assert.equal(ws.rows[2].cells[1].fill.fgColor.argb, 'FFFFFFFF');
});

test('não vincula palavras descritivas que contêm TSK em frase maior', () => {
  const wb=mockWorkbook();
  const ws=addFormattedSheet(wb,'Filtrado',[
    ['Resumo'], ['Análise da TSK261000114196 em andamento']
  ],{});
  assert.equal(ws.rows[1].cells[0].value, 'Análise da TSK261000114196 em andamento');
});

test('preserva coluna TSK separada ao reconhecer Número de Ordem', () => {
  const start = source.indexOf('  function buildDynamicOutputLayout(');
  const end = source.indexOf('\n  function calculateColumnWidths(',start);
  assert.ok(start>=0 && end>start);
  const buildDynamicOutputLayout=vm.runInNewContext(
    `(${source.slice(start,end).trim()})`,
    { COLUMN_SCHEMA:[{key:'order',output:'Número de Ordem'}, {key:'state',output:'Estado'}], normalizeHeaderForLookup:v=>String(v).toLowerCase().trim() }
  );
  const originalHeader=['Número de Ordem','TSK','Estado'];
  const analysis={columns:{order:{index:0},state:{index:2}}};
  const result=buildDynamicOutputLayout(originalHeader,analysis);
  assert.deepEqual(Array.from(result.header),['Número de Ordem','TSK','Estado','Status/Resumo']);
  assert.equal(result.orderIndex,0);
  assert.equal(result.selectedColumns[1].sourceIndex,1);
});
