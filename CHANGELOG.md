# Histórico de versões — NotaSync

## 1.5.0 — SmartSync (2026-10-17)

- Detecção de **Indisponibilidade Total**, ignorando caixa e acentuação, em quaisquer colunas da origem. Contador na tela, relação das ocorrências e destaque vermelho no arquivo Excel (na célula correspondente ou na TSK, se a coluna da origem não é exportada).
- Comparação das TSKs ativas entre a planilha anterior (`Filtrado`) e a atual, preservando somente a lista atual na aba `Filtrado`.
- TSKs anteriores não encontradas na atual são registradas na aba `Histórico TSKs` com a indicação **Possível normalização — verificar**; não há inferência de encerramento confirmado.
- O histórico de movimentações é reaproveitado no próximo cruzamento desde que se use a última planilha exportada como arquivo anterior.
- Registro de reaparecimento de TSK previamente marcada como ausente.
- Aviso para variações grandes, TSKs duplicadas e divergências de resumo em registros anteriores.
- Proteção contra sobrescrever resumos úteis com valores vazios e contra tratar chaves com zeros à esquerda como números.
- Migração do leitor SheetJS 0.18.5 para o **SheetJS 0.20.3 oficial**; remoção de `xlsx-js-style` antigo, e geração XLSX formatada usando **ExcelJS 4.4.0**.
- Logo OrbyttaCore com base branca, substituindo a identificação pessoal da interface e dos metadados gerados.
- Testes de lógica de comparação em `tests/`.

### Atenção

- O NotaSync continua **sem backend**; não mantém histórico automático entre sessões. Para manter a trilha, use a última planilha exportada como *planilha anterior*.
- A expressão “não localizada” indica diferença na consulta, não encerramento confirmado. As planilhas devem ter o **mesmo escopo, região e filtros**.
- Os scripts de planilhas são carregados de CDNs com versões fixas. É preciso conexão para carregar as bibliotecas. Recomenda-se hospedar cópias locais oficiais auditadas posteriormente.
- O arquivo de saída usa XLSX. A leitura de entradas `.xls`, `.csv`, `.cvs`, `.xlsx` foi mantida.

## 1.1 — 2026-09-06

- Importação de CSV, CVS, XLSX e XLS.
- Compatibilidade com planilhas já tratadas e colunas alternativas.
- Reaproveitamento do resumo, datas formatadas e UI responsiva.

## 1.0 — 2026-09-06

- Lançamento inicial do NotaSync.
