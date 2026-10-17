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
  const meaningful = value => {
    const text = String(value ?? '').trim();
    return text !== '' && !['N/A', 'NA', '-'].includes(text.toUpperCase());
  };
  const isTotalOutage = value => normalizeText(value).includes('indisponibilidade total');
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
  return { normalizeText, normalizeKey, meaningful, isTotalOutage, mergeSummary, compareActive, needsReview };
});
