# BI PC's Nao Autorizados

Sistema interno de registro, monitoramento e analise de ocorrencias em pontos de controle nao autorizados, desenvolvido em Google Apps Script com interface web integrada ao Google Sheets.

O projeto combina captura operacional de dados, dashboard analitico, visualizacao geografica, modo apresentacao para TV e suporte a deteccao automatica de paradas irregulares a partir de arquivos CSV de rastreamento.

## Visao Geral

Este sistema foi criado para apoiar a operacao da Viacao Catedral no acompanhamento de visitas/paradas em locais nao autorizados. A solucao centraliza o registro das ocorrencias em uma planilha, transforma os dados em indicadores e disponibiliza uma experiencia visual para analise diaria, acompanhamento gerencial e apresentacoes.

Do ponto de vista tecnico, o projeto demonstra uma entrega completa em ambiente Google Workspace: backend serverless com Apps Script, frontend HTML/CSS/JavaScript, integracao com Google Sheets, graficos interativos, mapa com Leaflet e rotinas auxiliares para enriquecimento dos dados.

## Destaques Para Recrutamento

- Desenvolvimento de uma solucao de BI operacional de ponta a ponta, desde a entrada de dados ate a visualizacao executiva.
- Automacao com Google Apps Script e Google Sheets, reduzindo tarefas manuais e padronizando registros.
- Dashboard web responsivo com filtros dinamicos, KPIs, rankings, historico paginado e mapa interativo.
- Modo apresentacao para monitores/TVs, com rotacao automatica de slides, periodo configuravel e atualizacao periodica.
- Processamento local de CSV de rastreamento para identificar paradas irregulares acima de 5 minutos.
- Integracao com Looker Studio e rotina de backfill usando dados externos do Supabase.
- Codigo organizado por responsabilidades: menu da planilha, formulario operacional, backend de BI e interfaces HTML.

## Funcionalidades

- Registro de ocorrencias via sidebar no Google Sheets.
- Cadastro rapido de motoristas diretamente pelo formulario.
- Consulta de locais, motoristas e bases a partir de abas da planilha.
- API `doPost` para receber ocorrencias de fontes externas.
- Dashboard web com:
  - total de visitas;
  - locais unicos;
  - motoristas envolvidos;
  - base mais ativa;
  - ranking de locais;
  - ranking de motoristas;
  - distribuicao por regiao;
  - serie temporal por dia;
  - ocorrencias por base;
  - ranking de linhas;
  - historico paginado e filtravel.
- Mapa interativo com Leaflet exibindo pontos autorizados/visitados e intensidade de ocorrencias.
- Modo apresentacao com 7 slides para acompanhamento gerencial.
- Upload de CSV de rastreamento para detectar paradas irregulares e copiar os resultados.
- Atalho no menu da planilha para abrir o historico, registrar ocorrencias e acessar dashboard externo.

## Stack

- Google Apps Script V8
- Google Sheets
- HTML, CSS e JavaScript
- Chart.js
- Leaflet
- Google Looker Studio
- Supabase REST API
- clasp para versionamento/deploy local do Apps Script

## Arquitetura

```text
Google Sheets
  |-- abas de apoio: PC'S NÃO AUTORIZADO, Motorista, Histórico
  |
  |-- Apps Script
      |-- Codigo.js        -> menu principal da planilha
      |-- ocorrencias.js   -> formulario, cadastro e API POST
      |-- BI_PC.js         -> web app, dados do BI e helpers
      |
      |-- Form.html        -> sidebar de registro operacional
      |-- index.html       -> dashboard principal
      |-- apresentacao.html-> modo apresentacao
```

### Fluxo Principal

1. O usuario registra uma ocorrencia pela sidebar `Form.html`.
2. O backend em `ocorrencias.js` valida os campos e grava os dados na aba `Histórico`.
3. O dashboard em `index.html` chama `getDadosBI(dataIni, dataFim)` para carregar os registros do periodo.
4. `BI_PC.js` cruza o histórico com a base de PCs, normaliza campos e devolve JSON para o frontend.
5. A interface renderiza KPIs, graficos, rankings, mapa e tabelas.
6. O modo apresentacao usa a mesma base de dados para exibir slides rotativos.

## Estrutura Das Abas

### `PC'S NÃO AUTORIZADO`

Aba usada como base de locais/pontos de controle.

Colunas relevantes:

- `A`: codigo do local
- `C`: descricao resumida
- `D`: descricao completa
- `E`: UF
- `F`: regiao
- `H`: tipo
- `AE`: latitude
- `AF`: longitude

### `Motorista`

Aba usada para cadastro e consulta de motoristas.

Colunas relevantes:

- `C`: nome do motorista
- `D`: matricula/ID
- `T`: base operacional

### `Histórico`

Aba onde as ocorrencias sao gravadas.

Colunas esperadas:

- `A`: ID Local
- `B`: Local
- `C`: Carro
- `D`: ID Motorista
- `E`: Motorista
- `F`: Data do Relatorio
- `G`: Base
- `H`: UF
- `I`: Regiao
- `J`: Linha

## Principais Arquivos

| Arquivo | Responsabilidade |
| --- | --- |
| `Codigo.js` | Cria o menu customizado da planilha e atalhos de navegacao. |
| `ocorrencias.js` | Lista locais/motoristas/bases, registra ocorrencias, cadastra motoristas e recebe POST externo. |
| `BI_PC.js` | Publica o web app, carrega dados para o BI, normaliza bases/datas/numeros e executa backfill de linhas. |
| `Form.html` | Interface lateral para registro de ocorrencias e cadastro rapido de motoristas. |
| `index.html` | Dashboard analitico principal com graficos, filtros, mapa, rankings e processamento de CSV. |
| `apresentacao.html` | Modo TV/apresentacao com slides automaticos e configuracao de periodo. |
| `appsscript.json` | Manifesto do Apps Script. |
| `.clasp.json` | Configuracao do projeto para deploy via clasp. |

## Como Executar

### 1. Preparar o Google Sheets

Crie ou use uma planilha com as abas:

- `PC'S NÃO AUTORIZADO`
- `Motorista`
- `Histórico`

Mantenha os nomes das abas e colunas conforme descrito acima, pois o Apps Script usa esses nomes como contrato de dados.

### 2. Configurar o Apps Script

Instale e autentique o clasp, se ainda nao estiver configurado:

```bash
npm install -g @google/clasp
clasp login
```

Depois, envie os arquivos para o projeto Apps Script:

```bash
clasp push
```

### 3. Publicar como Web App

No editor do Apps Script:

1. Abra `Deploy > New deployment`.
2. Escolha `Web app`.
3. Configure a execucao como usuario que fez o deploy.
4. Defina o acesso conforme a necessidade da operacao.
5. Copie a URL gerada para acessar o dashboard.

O dashboard principal abre pela URL padrao do web app. O modo apresentacao pode ser acessado com:

```text
?view=apresentacao
```

## Endpoints E Funcoes Importantes

- `doGet(e)`: entrega o dashboard ou o modo apresentacao.
- `getDadosBI(dataIni, dataFim)`: retorna os registros do periodo em JSON.
- `getPCs()`: retorna a base de pontos de controle com coordenadas.
- `openSidebarOcorrencias()`: abre o formulario lateral na planilha.
- `addOcorrencia(payload)`: grava uma nova ocorrencia no historico.
- `addMotorista(payload)`: cadastra motorista na aba `Motorista`.
- `doPost(e)`: recebe ocorrencias externas via POST.
- `backfillLinhas()`: rotina manual para preencher linhas/trips no historico a partir do Supabase.

## Exemplo De POST

```json
{
  "token": "TOKEN_CONFIGURADO",
  "localId": "123",
  "localNome": "Nome do local",
  "carro": "24614",
  "motoristaId": "5028",
  "motoristaNome": "NOME DO MOTORISTA",
  "base": "BRASILIA",
  "dataRelatorio": "2026-04-30"
}
```

## Cuidados Antes De Publicar Como Portfolio

Antes de tornar este repositorio publico, recomenda-se:

- Remover ou rotacionar tokens, chaves e URLs sensiveis.
- Mover segredos para `PropertiesService` no Apps Script.
- Remover `scriptId` real do `.clasp.json` ou usar um projeto demonstrativo.
- Substituir nomes/dados internos por dados ficticios.
- Validar permissoes do web app, especialmente quando `access` estiver como `ANYONE_ANONYMOUS`.
- Adicionar capturas de tela anonimizadas do dashboard e do modo apresentacao.

## Melhorias Futuras

- Centralizar configuracoes em `PropertiesService`.
- Adicionar uma aba de configuracao para nomes de abas, URL do Looker Studio e parametros de integracao.
- Criar camada de validacao mais robusta para payloads recebidos via POST.
- Adicionar testes manuais documentados para deploy e validacao de dados.
- Tratar melhor erros de encoding e padronizar arquivos em UTF-8.
- Criar dataset demonstrativo para uso em entrevistas e portfolio publico.
- Separar CSS e JavaScript dos arquivos HTML em um processo de build, se o projeto crescer.

## Autor

Desenvolvido por Lucas Inacio.

Projeto com foco em automacao operacional, visualizacao de dados e construcao de ferramentas internas para tomada de decisao.
