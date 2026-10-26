const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const start = source.indexOf('  function validateSerializedHyperlinks(');
const end = source.indexOf('\n  function resetCounters()', start);
assert.ok(start >= 0 && end > start, 'validação serializada não encontrada');
const functionSrc = source.slice(start, end).trim();

function checker({ actualUrl, expectedUrl, count = 1 }) {
  const logs = [];
  const ctx = {
    linkAudit: { count, samples: [{sheet: 'Filtrado', address: 'A2', url: expectedUrl}] },
    log: (text, cls) => logs.push([text, cls]),
    formatBytes: bytes => String(bytes),
    XLSX: { read: () => ({ Sheets: { Filtrado: { A2: actualUrl === null ? {v:'TSK261000114196'} : { l: { Target: actualUrl } } } } }) }
  };
  const fn = vm.runInNewContext(`(${functionSrc})`, ctx);
  return { fn, logs };
}
const URL = 'https://timbrasil.service-now.com/now/nav/ui/classic/params/target/u_task_evento.do%3Fsys_id%3DTSK261000114196';

test('verifica hyperlink preservado no XLSX serializado', () => {
  const { fn, logs } = checker({actualUrl: URL,expectedUrl: URL});
  assert.doesNotThrow(() => fn(new Uint8Array([0,1,2]).buffer));
  assert.ok(logs.some(x => /Validação do XLSX concluída/.test(x[0])));
});
test('bloqueia exportação sem hyperlink no arquivo serializado', () => {
  const { fn } = checker({actualUrl: null,expectedUrl: URL});
  assert.throws(() => fn(new Uint8Array([0,1,2]).buffer), /Falha de verificação do hiperlink/);
});
test('bloqueia exportação quando link salvo aponta para TSK errada', () => {
  const { fn } = checker({actualUrl: URL+'999',expectedUrl: URL});
  assert.throws(() => fn(new Uint8Array([0,1,2]).buffer), /Falha de verificação do hiperlink/);
});
