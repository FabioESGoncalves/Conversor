/* NotaSync SmartSync — regras testáveis e independentes da interface. */
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.NotaSyncCore = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const normalizeText = value => String(value ?? '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[_-]+/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
  // Não converter para número: 000123 e 123 são chaves diferentes.
  const normalizeKey = value => String(value ?? '').trim().toUpperCase();
  // URL fixa de destino: nunca utilizar o conteúdo da célula como URL completa.
  // Aceita somente identificadores TSK alfanuméricos (ex.: TSK261000114196).
  const SERVICE_NOW_TASK_PREFIX = 'https://timbrasil.service-now.com/now/nav/ui/classic/params/target/u_task_evento.do%3Fsys_id%3D';
  function serviceNowTaskUrl(value) {
    const key = normalizeKey(value).replace(/[\u200B-\u200D\uFEFF]/g, "");
    if (!/^TSK[A-Z0-9]+$/.test(key)) return null;
    return SERVICE_NOW_TASK_PREFIX + encodeURIComponent(key);
  }
  const meaningful = value => {
    const text = String(value ?? '').trim();
    return text !== '' && !['N/A', 'NA', '-'].includes(text.toUpperCase());
  };
  const isTotalOutage = value => normalizeText(value).includes('indisponibilidade total');
  const outageCellIndexes = cells => new Set(
    cells.flatMap((value, index) => isTotalOutage(value) ? [index] : [])
  );
  function mergeSummary(current, previous, hasPrevious) {
    if (hasPrevious && meaningful(previous)) return previous;
    if (meaningful(current)) return current;
    // Não apagar resumos atuais com um valor anterior vazio.
    return hasPrevious ? (String(current ?? '').trim() || 'N/A') : (current ?? '');
  }
  function compareActive(previousMap, currentKeys) {
    const gone = [];
    for (const [key, item] of previousMap) {
      if (!currentKeys.has(key)) gone.push({ key, ...item });
    }
    return gone;
  }
  function needsReview(goneCount, previousCount, currentCount) {
    return previousCount > 0 && (currentCount === 0 || goneCount >= 3 && goneCount / previousCount >= 0.4);
  }
  return { normalizeText, normalizeKey, meaningful, isTotalOutage, outageCellIndexes, mergeSummary, compareActive, needsReview, serviceNowTaskUrl };
});
