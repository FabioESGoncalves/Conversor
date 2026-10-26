const test = require('node:test');
const assert = require('node:assert/strict');
const core = require('../smartsync-core.js');

test('detecta indisponibilidade total mesmo sem acento e caixa', () => {
  for (const s of ['INDISPONIBILIDADE TOTAL', 'Falha: indisponibilidade total', ' Indisponibilidade   Total — CE ']) {
    assert.equal(core.isTotalOutage(s), true);
  }
  assert.equal(core.isTotalOutage('Indisponibilidade parcial'), false);
});
test('preserva zeros à esquerda e normaliza caixa', () => {
  assert.equal(core.normalizeKey('  tsk001  '), 'TSK001');
  assert.notEqual(core.normalizeKey('000123'), core.normalizeKey('123'));
});
test('não apaga resumo atual quando anterior vem vazio', () => {
  assert.equal(core.mergeSummary('Técnico acionado', '', true), 'Técnico acionado');
  assert.equal(core.mergeSummary('Técnico acionado', 'N/A', true), 'Técnico acionado');
  assert.equal(core.mergeSummary('Técnico acionado', 'Novo parecer', true), 'Novo parecer');
  assert.equal(core.mergeSummary('', '', true), 'N/A');
});
test('identifica TSK ausente sem incluí-la na lista atual', () => {
  const prior = new Map([['TSK001', {displayKey:'TSK001',summary:'Teste'}], ['TSK002',{displayKey:'TSK002'}]]);
  const missing = core.compareActive(prior, new Set(['TSK002', 'TSK003']));
  assert.deepEqual(missing.map(x=>x.key), ['TSK001']);
});
test('alerta sobre queda brusca de volume', () => {
  assert.equal(core.needsReview(41, 100, 59), true);
  assert.equal(core.needsReview(2, 5, 3), false);
  assert.equal(core.needsReview(4, 4, 0), true);
});

test('destaca somente a célula que contém indisponibilidade total, jamais a TSK por associação', () => {
  const columns = ['TSK001', 'Não iniciado', 'Indisponibilidade Total de serviço', 'Em análise'];
  assert.deepEqual([...core.outageCellIndexes(columns)], [2]);
  assert.deepEqual([...core.outageCellIndexes(['TSK002', 'Não iniciado', 'Indisponibilidade parcial'])], []);
});

test('gera hiperlink ServiceNow com identificador TSK da célula', () => {
  const expected = 'https://timbrasil.service-now.com/now/nav/ui/classic/params/target/u_task_evento.do%3Fsys_id%3DTSK261000114196';
  assert.equal(core.serviceNowTaskUrl('TSK261000114196'), expected);
  assert.equal(core.serviceNowTaskUrl(' tsk261000114196 '), expected);
  assert.equal(core.serviceNowTaskUrl('TSK001'), 'https://timbrasil.service-now.com/now/nav/ui/classic/params/target/u_task_evento.do%3Fsys_id%3DTSK001');
});

test('não transforma identificadores estranhos ou textos comuns em links', () => {
  for (const key of ['N/A', '', '12345', 'TSK/123', 'TSK123?x=1', '=HYPERLINK("https://outro.site")', 'javascript:alert(1)', 'TSK-123']) {
    assert.equal(core.serviceNowTaskUrl(key), null, key);
  }
});
