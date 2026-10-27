# NotaSync 1.5.6 — SmartSync

Ferramenta web estática da **OrbyttaCore** para leitura de CSV/CVS/XLS/XLSX, filtragem por `Estado = Não iniciado`, cruzamento de TSKs e exportação XLSX. Todo o processamento das planilhas acontece **no navegador**; o app não contém backend, banco de dados ou envio dos arquivos a um servidor.

## Como usar

1. Abra a página GitHub Pages (ou `index.html` em navegador com acesso à internet).
2. Informe a **planilha atual** com o estado dos chamados.
3. Selecione a **planilha anterior**, idealmente o XLSX exportado pelo NotaSync na última consulta.
4. Clique em **Processar planilha**. Revise contadores, registros e eventuais avisos de diferença de escopo.
5. Clique em **Baixar arquivo XLSX** e salve essa versão para o próximo cruzamento.

## Regras do SmartSync

- A aba **Filtrado** contém **somente os registros atuais** em que a coluna `Estado` contém `Não iniciado`. Registros ausentes não são devolvidos à planilha ativa.
- A aba **Histórico TSKs** registra os chamados que apareciam anteriormente e não aparecem entre os ativos agora, com a classificação **Possível normalização — verificar**. A diferença *não comprova fechamento*. O histórico é carregado do XLSX anterior e acumulado nas exportações seguintes.
- Chamados presentes agora e com registro prévio de ausência recebem no histórico o evento **Reapareceu na consulta atual**.
- A planilha atual pode ter `Número de Ordem` ou `TSK` como chave; a comparação ignora letras maiúsculas/minúsculas, mas mantém zeros à esquerda. O preenchimento anterior de `Status/Resumo` / `OBS` é reaproveitado apenas quando válido. Resumos vazios ou conflitantes não apagam anotações úteis presentes na planilha atual.
- Se o texto **Indisponibilidade Total** estiver presente em uma célula exportada, apenas essa célula recebe destaque vermelho. A coluna TSK não é colorida por associação.
- Diferenças grandes entre planilhas mostram um aviso. **Confira região, filtro, abrangência e período** antes de classificar como normalização.
- Não há persistência automática no navegador. O histórico só evolui se a planilha anterior contiver a aba `Histórico TSKs`.


## TSK clicável — ServiceNow

- Na planilha exportada, clique (ou Ctrl+clique, dependendo da versão do Excel) no código TSK da aba **Filtrado** para abrir o chamado no ServiceNow.
- A coluna TSK do **Histórico TSKs** também contém hyperlinks para consultar movimentações antigas.
- Nos painéis de movimentações do próprio NotaSync, os códigos TSK com link podem ser abertos em nova aba.
- Modelo: `https://timbrasil.service-now.com/now/nav/ui/classic/params/target/u_task_evento.do%3Fsys_id%3DTSK261000114196`
- A navegação exige acesso/autenticação válidos no ServiceNow. Somente códigos que começam com `TSK` e seguem com letras/números recebem link; outros valores permanecem texto normal.
- O código visível da célula é preservado: o endereço é armazenado como hyperlink XLSX, não como fórmula ou texto substituto.

## Arquivos

- `index.html` — interface e bibliotecas de planilhas.
- `styles.css` — identidade visual e interface.
- `app.js` — fluxo de arquivos, processamento, formatação e exportação.
- `smartsync-core.js` — regras isoladas de comparação para testes.
- `assets/` — iconografia NotaSync e logo OrbyttaCore com fundo branco.
- `tests/` — testes das regras da sincronização.

## Dependências

- [SheetJS CE 0.20.3](https://docs.sheetjs.com/docs/getting-started/installation/standalone/) do CDN oficial, para leitura de XLS, XLSX e CSV.
- [ExcelJS 4.4.0](https://github.com/exceljs/exceljs) via jsDelivr, para XLSX com estilos, filtros, células de data e duas abas.

Ambas as dependências são carregadas por scripts externos quando o site é aberto. O app continua sem enviar as planilhas, mas o carregamento das bibliotecas exige conexão e confiança nos fornecedores. Para uma implantação offline, hospede cópias confiáveis das bibliotecas e ajuste os caminhos do `index.html`.

## Publicação

Faça backup da publicação atual. Envie os arquivos da raiz e a pasta `assets/` para a **branch publicada no GitHub Pages**, preservando nomes e estrutura. Aguarde o deployment e recarregue com Ctrl+F5 para descartar caches antigos. Faça um teste usando cópias de planilhas reais antes de adotar a versão para operação.

## Testes

No terminal, com Node.js 18+:

```bash
node --test tests/*.test.js
```

© 2026 OrbyttaCore — NotaSync.

### Publicação correta no GitHub Pages (1.5.5)

1. Extraia o ZIP e copie **o conteúdo interno** da pasta `NotaSync-1.5.6-Links-Discretos/` para a **raiz** da branch publicada no GitHub Pages. Não crie uma pasta extra.
2. Confirme que o `index.html` publicado mostra `v1.5.5` no canto superior direito da área OrbyttaCore. Sem isso, a versão acessada ainda é antiga.
3. A marca está incorporada ao `index.html`; o arquivo `logo-orbyttacore-v152.jpg` é fornecido como recurso extra de inspeção, mas não é necessário para a marca aparecer.
4. Se o badge não mudar, confira **Settings → Pages** para identificar branch e pasta da publicação e abra a URL em aba anônima.

## Diagnóstico de links (v1.5.5)

- **Links ServiceNow** conta os links efetivamente montados para as abas do Excel; se mostrar `0`, confirme se os valores são TSKs válidas na coluna reconhecida como Número de Ordem/TSK.
- **Teste direto** aparece após processar; se abrir o chamado no navegador, o destino funciona.
- Ao clicar em **Baixar arquivo XLSX**, o NotaSync confere amostras do conteúdo XLSX serializado, verificando se os vínculos externos sobrevivem à exportação (arquivos de até 15 MB). Caso estejam ausentes, o download é interrompido com mensagem de diagnóstico.
- O link só aparece no arquivo **XLSX gerado** (não modifica o CSV nem o Excel original). Pode exigir `Ctrl + clique` no Excel desktop.
- A planilha não é enviada ao ServiceNow; é necessário estar autenticado para visualizar os chamados.

### Correção 1.5.5 — CSV → XLSX com TSK clicável

- O CSV de origem **não precisa conter hyperlinks**. Eles são adicionados ao XLSX exportado.
- Toda célula exportada contendo somente um identificador TSK válido vira um hyperlink (independente do nome ou posição da coluna).
- Quando o CSV possui "Número de Ordem" e "TSK" como colunas separadas, ambas são preservadas; o cruzamento mantém a chave que já era usada.
- O contador "Links ServiceNow" mostra quantos hyperlinks foram preparados e a validação verifica uma amostra do arquivo serializado antes do download (até 15 MB).


### 1.5.6 — Estilo discreto de hyperlinks

- O XLSX tem links clicáveis nos códigos TSK, mas visualmente são texto comum: cor escura (#243449), sem sublinhado.
- No Excel para Windows, use **Ctrl + clique** caso o programa exija para abrir links.
- Ao passar o mouse, o Excel ainda pode indicar que existe um link: isso é normal e não impede o estilo discreto da célula.
