# NotaSync

Ferramenta web para conversão, tratamento e cruzamento de arquivos **CSV/CVS ou XLSX para XLSX padronizado**, desenvolvida para simplificar o fluxo operacional de análise de notas em fila.

O sistema filtra automaticamente os registros com **Estado = Não iniciado**, mantém uma estrutura padronizada de colunas, converte campos de data para datas reais do Excel e permite cruzar a planilha atual com uma planilha tratada anteriormente.

## Funcionalidades

- Importação de arquivos `.csv`, `.cvs`, `.xlsx` e `.xls`.
- Filtro automático por **Estado = Não iniciado**.
- Conversão para `.xlsx`.
- Conversão de **Fim SLA** e **Criação do NTT** para datas reais do Excel.
- Inclusão automática da coluna **Status/Resumo** após a coluna **Faixa**.
- Cruzamento opcional com uma planilha XLSX anterior.
- Comparação baseada em **Número de Ordem**.
- Reaproveitamento do conteúdo já preenchido em **Status/Resumo**.
- Identificação de notas novas com `N/A` em **Status/Resumo**.
- Contadores de:
  - notas atuais;
  - notas já existentes;
  - notas novas;
  - distribuição por faixa P1, P2, P3, P4, P5 e Sem faixa.
- Geração de XLSX com cabeçalho formatado, autofiltro, larguras de coluna e bordas.
- Definição automática do nome do arquivo de saída de acordo com o nome do arquivo de origem.

## Fluxo de uso

1. Abra o sistema no navegador.
2. Selecione ou arraste o **CSV/CVS ou XLSX atual**.
3. Opcionalmente, selecione a **XLSX tratada anterior**.
4. Clique em **Processar planilha**.
5. Confira os indicadores apresentados na tela.
6. Clique em **Baixar arquivo XLSX**.

### Como funciona o cruzamento

A comparação utiliza a coluna **Número de Ordem** como chave.

- Se a ordem existir na planilha anterior, o sistema reaproveita o valor de **Status/Resumo**.
- Se a ordem não existir na planilha anterior, o sistema considera a nota como nova e preenche **Status/Resumo** com `N/A`.

Isso elimina a necessidade de realizar manualmente um `PROCV` ou `PROCX` para identificar as novas notas.

A planilha atual também pode ser uma XLSX já tratada ou gerada anteriormente pelo NotaSync. Nesse caso, a aplicação reconhece tanto a coluna **Faixa Priorização Dispatching** quanto a coluna **Faixa** e preserva **Status/Resumo** quando não houver uma planilha anterior selecionada.

## Colunas de saída

A planilha final utiliza a seguinte ordem:

1. Número de Ordem
2. Estado
3. Fim SLA
4. CM
5. Criação do NTT
6. END_ID
7. NE ID
8. Regra usuário criador
9. Tipo da Falha
10. Título do Alarme
11. Faixa
12. Status/Resumo

## Estrutura do projeto

```text
conversor-planilha/
├── assets/
│   ├── icon-notasync.png
│   ├── favicon.ico
│   ├── favicon.png
│   └── apple-touch-icon.png
├── index.html
├── styles.css
├── app.js
├── README.md
└── CHANGELOG.md
```

### `index.html`

Contém a estrutura da interface e importa as bibliotecas necessárias.

### `styles.css`

Responsável pelo layout, identidade visual, responsividade, estados dos botões e animações da interface.

### `app.js`

Contém toda a lógica da aplicação:

- leitura dos arquivos;
- filtros;
- cruzamento das planilhas;
- contadores;
- tratamento das datas;
- geração e download do XLSX.

### `CHANGELOG.md`

Registra as mudanças importantes de cada versão publicada.

## Identidade visual

A identidade oficial utiliza o nome **NotaSync**, com azul escuro como base e laranja como destaque. O cabeçalho usa o ícone oficial acompanhado do nome renderizado em HTML/CSS para garantir nitidez e adaptação a diferentes tamanhos de tela.

Os arquivos de marca ficam em `assets/`:

- `icon-notasync.png`: ícone principal em alta resolução e símbolo exibido no cabeçalho;
- `favicon.ico` e `favicon.png`: ícones da aba do navegador;
- `apple-touch-icon.png`: ícone para atalhos em dispositivos móveis.

## Tecnologias utilizadas

- HTML5
- CSS3
- JavaScript
- [SheetJS](https://sheetjs.com/)
- [xlsx-js-style](https://github.com/gitbrent/xlsx-js-style)

As bibliotecas JavaScript são carregadas por CDN no `index.html`.

## Privacidade dos arquivos

O processamento dos arquivos é realizado no **navegador do usuário**. O código desta aplicação não possui rotina para enviar os arquivos selecionados a um servidor.

> Observação: as bibliotecas SheetJS e xlsx-js-style são carregadas externamente por CDN. Portanto, é necessária conexão com a internet para carregá-las quando ainda não estiverem disponíveis no cache do navegador.

## Execução local

Por ser uma aplicação web estática, não é necessário instalar Node.js, banco de dados ou servidor de aplicação.

Você pode abrir o `index.html` diretamente no navegador.

Para desenvolvimento, também pode utilizar qualquer servidor HTTP local, por exemplo a extensão **Live Server** do Visual Studio Code.

## Publicação no GitHub Pages

1. Crie um repositório no GitHub.
2. Envie `index.html`, `styles.css`, `app.js`, `README.md`, `CHANGELOG.md` e a pasta `assets` para a raiz do repositório.
3. Acesse **Settings > Pages**.
4. Em **Build and deployment**, selecione **Deploy from a branch**.
5. Escolha a branch principal, normalmente `main`, e a pasta `/ (root)`.
6. Salve as configurações.

Após a publicação, o GitHub disponibilizará o endereço da aplicação.

## Compatibilidade

Recomendado utilizar versões atuais de:

- Google Chrome
- Microsoft Edge
- Mozilla Firefox

O navegador precisa oferecer suporte às APIs modernas de arquivos utilizadas pelo JavaScript.

## Responsável

**Fábio Gonçalves**  
Contato: **fabiogoncalves.contato@gmail.com**

## Direitos

© 2026 Fábio Gonçalves. Todos os direitos reservados.  
**NotaSync**
