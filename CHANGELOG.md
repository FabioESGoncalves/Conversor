## 1.5.6 — Links discretos (2026-10-26)

- TSKs continuam com hyperlinks no XLSX, mas passam a usar a mesma cor `#243449`, fonte Calibri 11 e nenhum sublinhado.
- Alteração visual restrita às células TSK exportadas; destino ServiceNow, histórico, cruzamento e destaque de Indisponibilidade Total permanecem inalterados.
- Versão visível e parâmetros de cache atualizados para evitar carregar o app.js antigo após a publicação no GitHub Pages.

## 1.5.5 — Correção de hyperlinks CSV → XLSX

- O hyperlink agora é inserido em qualquer célula exportada cujo conteúdo seja exatamente um código TSK válido, sem depender de coluna fixa.
- Se o CSV possui "Número de Ordem" e "TSK" separadas, a coluna TSK também é preservada.
- A exportação e o cruzamento continuam usando o ExcelJS e a mesma chave operacional.
- Sem mudanças na regra de Indisponibilidade Total ou de histórico das TSKs.

## 1.5.4 — ServiceNow LinkCheck

- Validação dos hyperlinks **após a serialização XLSX** usando a biblioteca de leitura já carregada, com amostra por aba (até 15 MB).
- Contador de links criados e link direto de teste na interface.
- Melhoria na limpeza de caracteres invisíveis nos códigos TSK antes de construir o endereço.
- Cache-busting de HTML/JS/CSS com versão 154.
- Mantidas as regras SmartSync, destaque somente na célula de indisponibilidade total e marca OrbyttaCore.

## 1.5.3 — ServiceNow Links

- Links clicáveis em identificadores TSK da aba **Filtrado** no XLSX exportado.
- Links também na coluna **TSK** da aba **Histórico TSKs**, inclusive eventos preservados entre exportações.
- Links clicáveis nas TSKs das listas de movimentações exibidas no painel web.
- URL segura com domínio fixo e validação do identificador (somente padrões TSK alfanuméricos).
- Mantidos cruzamento, regras do SmartSync e destaque vermelho restrito ao texto de Indisponibilidade Total.

## 1.5.2 — LogoFix

- Logotipo oficial enviado pelo usuário incorporado ao HTML em formato JPEG com fundo branco, sem depender de caminhos `/assets/`.
- Fallback de marca em texto caso o navegador não aceite a imagem.
- Marcador de build `v1.5.2` visível no cabeçalho para diferenciar uma página antiga em cache.
- Cache-busting de `styles.css`, `app.js` e `smartsync-core.js`.
- Mantida a regra 1.5.1: vermelho somente nas células com “Indisponibilidade Total”, nunca na TSK por fallback.

## 1.5.1 — Hotfix de validação visual

- Corrigido o carregamento da marca OrbyttaCore com incorporação direta da logo na interface.
- Ajustada a regra de destaque em vermelho: agora somente as células que contêm a expressão "Indisponibilidade Total" são coloridas.
- Removido o fallback que destacava a TSK quando a expressão aparecia em colunas não exibidas.

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
