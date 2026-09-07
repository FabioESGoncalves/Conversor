# Changelog

Todas as alterações relevantes do NotaSync serão registradas neste arquivo.

## 1.1 - 2026-09-06

- Campo principal agora aceita CSV, CVS, XLSX e XLS.
- Planilhas XLSX já tratadas ou geradas pelo NotaSync podem ser reutilizadas como planilha atual.
- Compatibilidade com as colunas `Faixa Priorização Dispatching` e `Faixa`.
- Preservação de `Status/Resumo` ao reprocessar uma XLSX sem planilha anterior.
- Preservação do formato de data ao reutilizar XLSX geradas pelo sistema.
- Correção do logotipo cortado: cabeçalho passa a usar o ícone oficial e wordmark responsivo em HTML/CSS.
- Nomes de saída ES/Bxd são preservados ao reutilizar arquivos gerados anteriormente.

## 1.0 - 2026-09-06

- Nome oficial definido como **NotaSync**.
- Inclusão do logotipo oficial na interface.
- Inclusão de favicon para a aba do navegador.
- Inclusão de ícone para atalhos em dispositivos móveis.
- Interface dark corporativa em azul com destaques em laranja.
- Projeto separado em HTML, CSS e JavaScript.
- Área dedicada para CSV/CVS atual e XLSX anterior.
- Cruzamento por Número de Ordem.
- Inclusão e reaproveitamento da coluna Status/Resumo.
- Identificação e contagem de notas novas.
- Contadores de distribuição por faixa.
- Estados visuais para os botões Processar planilha e Baixar arquivo XLSX.
- Assinatura e contato do responsável na interface.
- Rodapé com ano atual.
- Formatação da planilha XLSX com bordas, autofiltro e cabeçalho.
