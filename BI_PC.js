// ============================================================
//  BI PC's Não Autorizados — Viação Catedral
//  BI_PC.gs — Google Apps Script Backend
//  Dev: Lucas Inácio
// ============================================================

// ── 1. Web App ───────────────────────────────────────────────
function doGet(e) {
  const view =
    e && e.parameter && e.parameter.view ? e.parameter.view : "dashboard";

  if (view === "apresentacao") {
    return HtmlService.createHtmlOutputFromFile("apresentacao")
      .setTitle("Modo Apresentação · BI PC's Não Autorizados · Viação Catedral")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (view === "tempo-permanencia") {
    return HtmlService.createHtmlOutputFromFile("tempo_permanencia")
      .setTitle("Tempo de Permanência · Viação Catedral")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (view === "motoristas") {
    return HtmlService.createHtmlOutputFromFile("motoristas")
      .setTitle("Motoristas · Viação Catedral")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("BI PC's Não Autorizados · Viação Catedral")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ── 1.1 Drive — pasta de relatórios (Excesso de Permanência) ───
// Chamada via google.script.run pelo tempo_permanencia.html antes de
// enviar o PDF ao Drive. Evita depender de service account + env var
// no backend (Koyeb): usa a autorização do próprio Apps Script (roda
// como o usuário que fez o deploy — ver "executeAs" no appsscript.json)
// pra achar (ou criar, na 1ª vez) a pasta e emprestar um token OAuth
// válido só pra essa chamada. O token dura ~1h, por isso é buscado de
// novo a cada "Gerar Relatório" em vez de guardado.
var DRIVE_PASTA_RELATORIOS_NOME = "Relatórios - Excesso de Permanência";
var PROP_PASTA_DRIVE_RELATORIOS_ID = "PASTA_DRIVE_RELATORIOS_ID";

function getCredenciaisDriveRelatorios() {
  var props = PropertiesService.getScriptProperties();
  var pastaSalvaId = props.getProperty(PROP_PASTA_DRIVE_RELATORIOS_ID);
  var pasta = null;

  if (pastaSalvaId) {
    try {
      pasta = DriveApp.getFolderById(pastaSalvaId);
    } catch (e) {
      // Pasta escolhida manualmente foi excluída ou o script perdeu acesso
      // a ela — cai pro fallback por nome abaixo em vez de quebrar.
      pasta = null;
    }
  }

  if (!pasta) {
    var pastas = DriveApp.getFoldersByName(DRIVE_PASTA_RELATORIOS_NOME);
    pasta = pastas.hasNext() ? pastas.next() : DriveApp.createFolder(DRIVE_PASTA_RELATORIOS_NOME);
  }

  return {
    folderId: pasta.getId(),
    folderName: pasta.getName(),
    accessToken: ScriptApp.getOAuthToken(),
  };
}

// Chamada pelo modal "Configurar Drive" em tempo_permanencia.html quando o
// usuário cola o link (ou ID) da pasta de destino no Drive — o Google
// Picker foi tentado e descartado, ver comentário em
// tempo_permanencia.html perto de _extrairFolderIdDeLink. Fica salva em
// ScriptProperties (vale pra todo mundo que gerar relatório a partir
// daqui, até alguém trocar de novo).
function definirPastaDriveRelatorios(folderId, folderName) {
  if (!folderId) throw new Error("folderId obrigatório");
  // Confirma que a pasta existe e o script realmente tem acesso a ela
  // antes de salvar — evita guardar um ID inválido.
  var pasta = DriveApp.getFolderById(folderId);
  PropertiesService.getScriptProperties().setProperty(PROP_PASTA_DRIVE_RELATORIOS_ID, folderId);
  return { folderId: pasta.getId(), folderName: pasta.getName() };
}

// ── 1.2 Resumo do dia (PDF) — Tempo de Permanência ─────────────
// Documento de registro/consulta com tudo que já foi ANALISADO numa
// região (relatório gerado OU descartado por justificativa) — não cria
// ocorrência nenhuma na API externa, diferente do "Relatório em
// massa" (esse sim cria ocorrências reais, só pros pontos PENDENTES).
// Usa a mesma pasta do Drive configurada em "Configurar Drive" (ver
// getCredenciaisDriveRelatorios acima). Layout inspirado no relatório
// diário do Gerador de Relatórios Operacionais (mesma identidade
// visual: cabeçalho laranja + logo, cards de resumo, barras de
// distribuição por motivo, tabela listando tudo) — HTML/CSS convertido
// pra PDF via Utilities/Blob, sem precisar do DocumentApp (evita pedir
// escopo novo de permissão pro Google Docs).

// Logo Viação Catedral em base64 (mesmo arquivo do Gerador de
// Relatórios Operacionais, src/assets/catedral.png — na verdade um
// JPEG apesar da extensão .png, daí o mime type usado abaixo).
var CATEDRAL_LOGO_B64 =
  "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";

// Razão social/CNPJ pro rodapé do PDF — mesmo texto usado pelo Gerador
// de Relatórios Operacionais (a Viação Catedral é a marca; o CNPJ é
// registrado sob essa razão social). Sem numeração "Página X de Y"
// aqui: esse recurso, no sistema de referência, vem do Puppeteer
// (page.pdf com footerTemplate) — o conversor HTML→PDF do Apps Script
// não tem equivalente pra rodapé repetido/numeração por página.
var EMPRESA_RAZAO_SOCIAL = "KANDANGO TRANSPORTE E TURISMO LTDA";
var EMPRESA_CNPJ = "03.233.439/0001-52";

var MOTIVO_CORES_PADRAO = {
  ANTECIPADO: "#f0c040",
  INICIO_VIAGEM: "#3b82c4",
  EMBARQUE: "#22c97a",
  OCORRENCIA: "#8b5cf6",
  RELATORIO_GERADO: "#e05050",
};

function _escHtml(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// payload: { regiao, dataLabel, analisadoPor, totalItens,
//            itens: [{ veiculo, ponto, cidade, uf, entrada, saida,
//                       excedenteMin, motivo, motivoLabel, motivoCor,
//                       detalhe }] }
function _montarHtmlResumoAnaliseRegiao(payload) {
  var itens = payload.itens || [];
  var regiao = payload.regiao || "Região";
  var dataLabel = payload.dataLabel || "";

  var veiculos = {};
  var porMotivo = {}; // motivo -> { label, cor, count }
  var maiorExcedente = 0;
  itens.forEach(function (it) {
    veiculos[it.veiculo] = true;
    maiorExcedente = Math.max(maiorExcedente, Number(it.excedenteMin) || 0);
    var m = it.motivo || "—";
    if (!porMotivo[m]) {
      porMotivo[m] = { label: it.motivoLabel || m, cor: it.motivoCor || MOTIVO_CORES_PADRAO[m] || "#9ca3af", count: 0 };
    }
    porMotivo[m].count++;
  });
  var qtdVeiculos = Object.keys(veiculos).length;
  var culpaCount = (porMotivo.RELATORIO_GERADO || { count: 0 }).count;
  var descarteCount = itens.length - culpaCount;

  var motivosOrdenados = Object.keys(porMotivo)
    .map(function (k) { return porMotivo[k]; })
    .sort(function (a, b) { return b.count - a.count; });
  var maxMotivoCount = motivosOrdenados.length ? motivosOrdenados[0].count : 1;

  var distRows = motivosOrdenados.map(function (m) {
    var pct = Math.round((m.count / maxMotivoCount) * 100);
    return (
      "<tr>" +
      '<td class="dist-name">' + _escHtml(m.label) + "</td>" +
      '<td class="dist-bar-cell"><div class="dist-bar-bg"><div class="dist-bar-fill" style="width:' + pct + "%;background:" + m.cor + ';"></div></div></td>' +
      '<td class="dist-count">' + m.count + "</td>" +
      "</tr>"
    );
  }).join("");

  var ordenados = itens.slice().sort(function (a, b) {
    return String(a.veiculo).localeCompare(String(b.veiculo)) || String(a.entrada).localeCompare(String(b.entrada));
  });

  var linhasTabela = ordenados.map(function (it, idx) {
    var num = String(idx + 1).padStart(2, "0");
    var local = _escHtml(it.ponto) + (it.cidade ? ' <span class="occ-sub">(' + _escHtml(it.cidade) + (it.uf ? " - " + _escHtml(it.uf) : "") + ")</span>" : "");
    var cor = it.motivoCor || MOTIVO_CORES_PADRAO[it.motivo] || "#9ca3af";
    return (
      '<tr class="' + (idx % 2 === 0 ? "row-even" : "row-odd") + '">' +
      '<td class="occ-num occ-td">' + num + "</td>" +
      '<td class="occ-td col-nowrap" style="font-family:monospace;font-weight:700;">' + _escHtml(it.veiculo) + "</td>" +
      '<td class="occ-td">' + local + "</td>" +
      '<td class="occ-td col-nowrap col-center">' + _escHtml(it.entrada) + "</td>" +
      '<td class="occ-td col-nowrap col-center">' + _escHtml(it.saida) + "</td>" +
      '<td class="occ-td col-nowrap col-center" style="font-weight:700;color:#f47920;">+' + (it.excedenteMin != null ? it.excedenteMin : "—") + " min</td>" +
      '<td class="occ-td col-nowrap col-center"><span class="motivo-badge" style="background:' + cor + ';">' + _escHtml(it.motivoLabel) + "</span></td>" +
      '<td class="occ-td" style="font-size:7.5pt;color:#6b7280;">' + _escHtml(it.detalhe || "—") + "</td>" +
      "</tr>"
    );
  }).join("");

  var geradoEm = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy 'às' HH:mm");

  return (
    '<!doctype html><html><head><meta charset="utf-8" /><title>Resumo de Análise — ' + _escHtml(regiao) + "</title><style>" +
    "@page { margin-top: 18mm; margin-right: 14mm; margin-left: 14mm; margin-bottom: 18mm; }" +
    'body { font-family: "Segoe UI", Arial, Helvetica, sans-serif; font-size: 10.5pt; color: #111; margin: 0; }' +
    ".report-header { width:100%; border-collapse:collapse; border:1.5px solid #e2e8f0; border-radius:5px; overflow:hidden; margin-bottom:16px; }" +
    ".report-header td { vertical-align: middle; }" +
    ".report-header-logo { background:#fff; padding:10px 18px; width:110px; border-right:1.5px solid #e2e8f0; text-align:center; }" +
    ".report-header-logo img { height:42px; }" +
    ".report-header-body { padding:14px 20px; background:#f47920; }" +
    ".report-header-title { font-size:14.5pt; font-weight:800; color:#fff; letter-spacing:.5px; text-transform:uppercase; }" +
    ".report-header-sub { font-size:9pt; color:rgba(255,255,255,.9); margin-top:3px; }" +
    ".summary-table { width:100%; border-collapse:separate; border-spacing:6px 0; margin:0 0 16px -6px; }" +
    ".summary-cell { text-align:center; padding:10px 6px; border:1px solid #e5e7eb; border-radius:4px; background:#fafafa; }" +
    ".summary-number { font-size:19pt; font-weight:700; color:#1e293b; line-height:1; }" +
    ".summary-label { font-size:7.5pt; color:#6b7280; margin-top:4px; text-transform:uppercase; letter-spacing:.4px; }" +
    ".culpa-cell { background:#fef2f2; border-color:#fca5a5; } .culpa-cell .summary-number { color:#e05050; }" +
    ".descarte-cell { background:#f0fdf4; border-color:#86efac; } .descarte-cell .summary-number { color:#22c97a; }" +
    ".section-header { font-size:9.5pt; font-weight:700; text-transform:uppercase; letter-spacing:.7px; color:#f47920; border-bottom:2px solid #f47920; padding-bottom:3px; margin-bottom:8px; }" +
    ".dist-table { width:100%; border-collapse:collapse; margin-bottom:18px; }" +
    ".dist-name { font-size:9pt; padding:3px 6px 3px 0; width:32%; }" +
    ".dist-bar-cell { padding:4px 8px; }" +
    ".dist-bar-bg { background:#f3f4f6; border-radius:3px; height:9px; }" +
    ".dist-bar-fill { height:9px; border-radius:3px; }" +
    ".dist-count { font-size:9pt; font-weight:700; text-align:right; width:26px; padding:3px 0; }" +
    ".occ-table { width:100%; border-collapse:collapse; font-size:8pt; border:1px solid #d1d5db; }" +
    ".occ-th { background:#f47920; color:#fff; font-weight:700; padding:6px 7px; text-align:left; font-size:7.5pt; white-space:nowrap; letter-spacing:.3px; border-right:1px solid rgba(255,255,255,.25); }" +
    ".occ-th:last-child { border-right:none; }" +
    ".occ-num { padding:5px 6px; color:#9ca3af; font-weight:700; text-align:center; border-right:1px solid #e5e7eb; }" +
    ".occ-td { padding:5px 7px; vertical-align:middle; border-right:1px solid #e5e7eb; border-bottom:1px solid #e5e7eb; }" +
    ".occ-td:last-child { border-right:none; }" +
    ".occ-sub { color:#9ca3af; font-size:7.5pt; }" +
    ".col-nowrap { white-space:nowrap; } .col-center { text-align:center; }" +
    ".row-even td { background:#fff; } .row-odd td { background:#f9fafb; }" +
    ".motivo-badge { display:inline-block; border-radius:3px; font-size:7pt; font-weight:700; padding:2px 6px; color:#fff; white-space:nowrap; }" +
    ".report-footer-table { position:fixed; bottom:0; left:0; right:0; width:100%; border-collapse:collapse; background:#fff; }" +
    ".report-footer-table td { border-top:1px solid #e5e7eb; padding-top:8px; font-size:7.5pt; color:#9ca3af; vertical-align:top; }" +
    ".footer-left { text-align:left; width:33%; }" +
    ".footer-center { text-align:center; width:34%; line-height:1.4; color:#6b7280; }" +
    ".footer-right { text-align:right; width:33%; }" +
    "</style></head><body>" +
    '<table class="report-header"><tr>' +
    '<td class="report-header-logo"><img src="data:image/jpeg;base64,' + CATEDRAL_LOGO_B64 + '" alt="Viação Catedral" /></td>' +
    '<td class="report-header-body">' +
    '<div class="report-header-title">Resumo de Análise — Tempo de Permanência</div>' +
    '<div class="report-header-sub">' + _escHtml(regiao) + " &nbsp;·&nbsp; " + _escHtml(dataLabel) + " &nbsp;·&nbsp; " + itens.length + " excedência" + (itens.length !== 1 ? "s" : "") + " analisada" + (itens.length !== 1 ? "s" : "") + "</div>" +
    "</td></tr></table>" +
    '<table class="summary-table"><tr>' +
    '<td class="summary-cell"><div class="summary-number">' + itens.length + '</div><div class="summary-label">Analisadas</div></td>' +
    '<td class="summary-cell"><div class="summary-number">' + qtdVeiculos + '</div><div class="summary-label">Veículos</div></td>' +
    '<td class="summary-cell"><div class="summary-number">+' + maiorExcedente + '</div><div class="summary-label">Maior excedente (min)</div></td>' +
    '<td class="summary-cell culpa-cell"><div class="summary-number">' + culpaCount + '</div><div class="summary-label">Culpa do motorista</div></td>' +
    '<td class="summary-cell descarte-cell"><div class="summary-number">' + descarteCount + '</div><div class="summary-label">Descartadas</div></td>' +
    "</tr></table>" +
    (motivosOrdenados.length ? '<div class="section-header">Distribuição por motivo</div><table class="dist-table">' + distRows + "</table>" : "") +
    '<div class="section-header">Listagem completa</div>' +
    '<table class="occ-table"><thead><tr>' +
    '<th class="occ-th" style="width:22px;text-align:center;">#</th>' +
    '<th class="occ-th" style="width:56px;">Veículo</th>' +
    '<th class="occ-th">Ponto</th>' +
    '<th class="occ-th" style="width:52px;">Chegada</th>' +
    '<th class="occ-th" style="width:48px;">Saída</th>' +
    '<th class="occ-th" style="width:64px;">Excedente</th>' +
    '<th class="occ-th" style="width:112px;">Motivo</th>' +
    '<th class="occ-th">Detalhe</th>' +
    "</tr></thead><tbody>" + linhasTabela + "</tbody></table>" +
    '<table class="report-footer-table"><tr>' +
    '<td class="footer-left">' + (payload.analisadoPor ? "Por " + _escHtml(payload.analisadoPor) : "") + "</td>" +
    '<td class="footer-center">' + _escHtml(EMPRESA_RAZAO_SOCIAL) + "<br/>CNPJ: " + _escHtml(EMPRESA_CNPJ) + "</td>" +
    '<td class="footer-right">Gerado em ' + geradoEm + "</td>" +
    "</tr></table>" +
    "</body></html>"
  );
}

function gerarResumoAnaliseRegiaoPdf(payload) {
  payload = payload || {};
  var itens = payload.itens || [];
  if (!itens.length) throw new Error("Nenhuma excedência analisada pra resumir.");

  var regiao = payload.regiao || "Região";
  var dataLabel = payload.dataLabel || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");

  var pastaInfo = getCredenciaisDriveRelatorios();
  var pasta = DriveApp.getFolderById(pastaInfo.folderId);

  var html = _montarHtmlResumoAnaliseRegiao(payload);
  var htmlBlob = Utilities.newBlob(html, "text/html", "resumo.html");
  var pdfBlob = htmlBlob.getAs(MimeType.PDF);
  var nomeArquivo = ("Resumo Analise - " + regiao + " - " + dataLabel).replace(/\//g, ".") + ".pdf";
  pdfBlob.setName(nomeArquivo);
  var pdfFile = pasta.createFile(pdfBlob);

  return { url: pdfFile.getUrl(), fileName: pdfFile.getName() };
}

// ── 2. Estrutura da aba "PC'S NÃO AUTORIZADO" (índices 0-based) ─
// Col  0  → Código
// Col  1  → Cód. Emb.
// Col  2  → Desc. Resumida
// Col  3  → Descrição
// Col  4  → UF
// Col  5  → REGIÃO
// Col  6  → Unidade Empresarial
// Col  7  → Tipo
// Col  8  → Raio ... (não usados) ...
// Col 30  → Latitude_OK
// Col 31  → Longitude_OK
// Total de colunas lidas: 32

// ── 3. Estrutura da aba "Histórico" (índices 0-based) ────────
// Col 0 → ID Local
// Col 1 → Local
// Col 2 → Carro
// Col 3 → ID Motorista
// Col 4 → Motorista
// Col 5 → Data do Relatório
// Col 6 → Base
// Col 7 → UF       (pode ser fórmula na planilha)
// Col 8 → Região   (pode ser fórmula na planilha)
// Col 9 → Linha    (ex: "0126", "BRASÍLIA-RIO")
// Total de colunas lidas: 10

// ── 3b. Normalização de bases ────────────────────────────────
// _normalizeBase(raw) remove acentos, converte para maiúsculas
// e elimina sufixos operacionais (CATEDRAL, FILIAL, etc.)
// Não depende de aba de configurações.

// ── 4. getPCs() ──────────────────────────────────────────────
function getPCs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("PC'S NÃO AUTORIZADO");
  if (!aba) throw new Error('Aba "PC\'S NÃO AUTORIZADO" não encontrada.');

  const lastRow = aba.getLastRow();
  if (lastRow < 2) return JSON.stringify([]);

  const data = aba.getRange(2, 1, lastRow - 1, 32).getValues();

  const pcs = data
    .filter((r) => String(r[0] || "").trim() !== "")
    .map((r) => ({
      codigo: String(r[0] || "").trim(),
      descResumida: String(r[2] || "").trim(),
      descCompleta: String(r[3] || "").trim(),
      uf: String(r[4] || "").trim(),
      regiao: String(r[5] || "").trim(),
      tipo: String(r[7] || "").trim(),
      lat: _coordOuNull(r[30], 90),
      lng: _coordOuNull(r[31], 180),
    }));

  return JSON.stringify(pcs);
}

// ── 4b. getPontosControle() ──────────────────────────────────
// Base de referência usada só pelo Tempo de Permanência (excesso de
// parada) pra vincular o nome do ponto (do CSV de rastreamento) à
// lat/long e calcular a região no mapa/gráficos. Aba diferente de
// getPCs() (que é a "PC'S NÃO AUTORIZADO", usada por index.html e
// apresentacao.html) — não mexer em getPCs() pra não afetar as outras
// telas.
// Col D (índice 3)  → nome/descrição do ponto (usado pro match)
// Col Y (índice 24) → Latitude
// Col Z (índice 25) → Longitude
function getPontosControle() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("PONTOS_CONTROLE");
  if (!aba) throw new Error('Aba "PONTOS_CONTROLE" não encontrada.');

  const lastRow = aba.getLastRow();
  if (lastRow < 2) return JSON.stringify([]);

  const data = aba.getRange(2, 1, lastRow - 1, 26).getValues();

  const pontos = data
    .filter((r) => String(r[3] || "").trim() !== "")
    .map((r) => ({
      descCompleta: String(r[3] || "").trim(),
      lat: _coordOuNull(r[24], 90),
      lng: _coordOuNull(r[25], 180),
    }));

  return JSON.stringify(pontos);
}

// ── 5. getDadosBI(dataIni, dataFim) ──────────────────────────
function getDadosBI(dataIni, dataFim) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Monta mapa de PCs
  const pcMap = _buildPcMap(ss);

  const abaHist = ss.getSheetByName("Histórico");
  if (!abaHist) throw new Error('Aba "Histórico" não encontrada.');

  const lastRow = abaHist.getLastRow();
  if (lastRow < 2) return JSON.stringify([]);

  const data = abaHist.getRange(2, 1, lastRow - 1, 10).getValues();

  const dtIni = new Date(dataIni + "T00:00:00");
  const dtFim = new Date(dataFim + "T23:59:59");

  const registros = [];

  data.forEach(function (r) {
    const dataBruta = _normData(r[5]);
    if (!dataBruta) return;
    if (dataBruta < dtIni || dataBruta > dtFim) return;

    const y = dataBruta.getFullYear();
    const m = String(dataBruta.getMonth() + 1).padStart(2, "0");
    const d = String(dataBruta.getDate()).padStart(2, "0");

    const idLocal = String(r[0] || "").trim();
    const pc = pcMap[idLocal] || {};
    const base = String(r[6] || "").trim();

    registros.push({
      idLocal: idLocal,
      local: String(r[1] || "").trim(),
      carro: String(r[2] || "").trim(),
      idMotorista: String(r[3] || "").trim(),
      motorista: String(r[4] || "").trim(),
      data: y + "-" + m + "-" + d,
      base: base,
      baseNorm: _normalizeBase(base) || base,
      uf: String(r[7] || "").trim() || pc.uf || "",
      regiao: String(r[8] || "").trim() || pc.regiao || "",
      linha: String(r[9] || "").trim(),
      lat: pc.lat || null,
      lng: pc.lng || null,
      tipo: pc.tipo || "",
      descCompleta: pc.descCompleta || String(r[1] || "").trim(),
      descResumida: pc.descResumida || String(r[1] || "").trim(),
    });
  });

  return JSON.stringify(registros);
}

// ── 5b. Status de análise — Tempo de Permanência ──────────────
// Aba "TEMPO_PERMANENCIA_STATUS": CHAVE, VEICULO, PONTO, ENTRADA,
// STATUS, ATUALIZADO_EM — guarda quais excedências já foram
// marcadas como analisadas, pra sobreviver a reload/novo upload
// do mesmo CSV. CHAVE = "<veiculo>|<entrada>" (gerada no front-end).
var TEMPO_PERMANENCIA_STATUS_HEADER = [
  "CHAVE", "VEICULO", "PONTO", "ENTRADA", "STATUS", "ATUALIZADO_EM",
  "MOTIVO", "QTD_EMBARQUE", "QTD_DESEMBARQUE", "APOIO_RODOVIARIA", "OBSERVACAO",
  "LINHA_COD", "LINHA_NOME", "LINHA_TRIP_ID", "HORARIO_SESSAO", "PONTO_SESSAO",
];

// Antes disto, marcarStatusPermanencia() lançava erro se a aba não
// existisse — como ela nunca é criada manualmente, TODA marcação de
// "analisado" falhava silenciosamente (o erro só ia pro console via
// withFailureHandler) e nada era persistido, por isso sumia após F5.
// Agora a aba é criada sozinha aqui, igual _abaHistoricoExcesso().
//
// As colunas MOTIVO..OBSERVACAO (G-K) foram adicionadas depois — guardam
// a justificativa exigida pra marcar uma excedência como analisada sem
// gerar relatório (ver abrirModalJustificativa() em tempo_permanencia.html).
// LINHA_COD..LINHA_TRIP_ID (L-N), adicionadas depois ainda, guardam a
// linha (itinerário) vinculada àquela excedência assim que ela é
// carregada/selecionada (esquema vinculado ou busca manual no modal de
// Gerar Relatório) — fica salva junto da análise, sem precisar
// reselecionar a linha depois (ver justRelatorio em _onClickBtnGerar()).
// HORARIO_SESSAO/PONTO_SESSAO (O-P) guardam o horário previsto de saída
// (sessão) e o nome do ponto do roteiro casado pelo esquema, específicos
// DAQUELE ponto de parada (cada ponto do roteiro tem seu próprio horário
// de sessão) — sem isso, o campo "sessionTime" mandado pra API ao criar a
// ocorrência (_criarOcorrencia) ficava null depois de um reload, mesmo
// pra pontos que já tinham esquema vinculado antes.
// Migra sozinha o cabeçalho de abas criadas antes disso, sem mexer nas
// linhas de dados já gravadas.
function _abaStatusPermanencia(ss) {
  var aba = ss.getSheetByName("TEMPO_PERMANENCIA_STATUS");
  if (!aba) {
    aba = ss.insertSheet("TEMPO_PERMANENCIA_STATUS");
    aba.appendRow(TEMPO_PERMANENCIA_STATUS_HEADER);
    return aba;
  }
  if (aba.getLastColumn() < TEMPO_PERMANENCIA_STATUS_HEADER.length) {
    aba.getRange(1, 1, 1, TEMPO_PERMANENCIA_STATUS_HEADER.length).setValues([TEMPO_PERMANENCIA_STATUS_HEADER]);
  }
  return aba;
}

// Devolve as excedências já marcadas como "ANALISADO", com a justificativa
// registrada — front-end usa pra pintar como riscado E pra saber o motivo.
function getStatusPermanencia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("TEMPO_PERMANENCIA_STATUS");
  if (!aba) return JSON.stringify([]);

  const lastRow = aba.getLastRow();
  if (lastRow < 2) return JSON.stringify([]);

  const data = aba.getRange(2, 1, lastRow - 1, TEMPO_PERMANENCIA_STATUS_HEADER.length).getValues();
  const registros = data
    .filter(
      (r) =>
        String(r[0] || "").trim() !== "" &&
        String(r[4] || "").trim().toUpperCase() === "ANALISADO",
    )
    .map((r) => ({
      chave: String(r[0]).trim(),
      motivo: String(r[6] || "").trim(),
      qtdEmbarque: String(r[7] || "").trim(),
      qtdDesembarque: String(r[8] || "").trim(),
      apoioRodoviaria: String(r[9] || "").trim(),
      observacao: String(r[10] || "").trim(),
      linhaCod: String(r[11] || "").trim(),
      linhaNome: String(r[12] || "").trim(),
      linhaTripId: String(r[13] || "").trim(),
      horarioSessao: String(r[14] || "").trim(),
      pontoSessao: String(r[15] || "").trim(),
    }));

  return JSON.stringify(registros);
}

// payload: { chave, veiculo, ponto, entrada, analisado, motivo,
//            qtdEmbarque, qtdDesembarque, apoioRodoviaria, observacao,
//            linhaCod, linhaNome, linhaTripId, horarioSessao, pontoSessao }
function marcarStatusPermanencia(payload) {
  const chave = String((payload && payload.chave) || "").trim();
  if (!chave) throw new Error("chave é obrigatória.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = _abaStatusPermanencia(ss);

  const status = payload.analisado ? "ANALISADO" : "";
  const agora = new Date();
  // Reverter (analisado=false) limpa a justificativa (e a linha/sessão)
  // junto — não faz sentido guardar dado de uma análise desfeita.
  const valoresLinha = [
    status,
    agora,
    payload.analisado ? payload.motivo || "" : "",
    payload.analisado ? payload.qtdEmbarque || "" : "",
    payload.analisado ? payload.qtdDesembarque || "" : "",
    payload.analisado ? payload.apoioRodoviaria || "" : "",
    payload.analisado ? payload.observacao || "" : "",
    payload.analisado ? payload.linhaCod || "" : "",
    payload.analisado ? payload.linhaNome || "" : "",
    payload.analisado ? payload.linhaTripId || "" : "",
    payload.analisado ? payload.horarioSessao || "" : "",
    payload.analisado ? payload.pontoSessao || "" : "",
  ];

  const lastRow = aba.getLastRow();
  let linhaExistente = -1;
  if (lastRow >= 2) {
    const chavesExistentes = aba.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < chavesExistentes.length; i++) {
      if (String(chavesExistentes[i][0]).trim() === chave) {
        linhaExistente = i + 2;
        break;
      }
    }
  }

  if (linhaExistente > 0) {
    aba.getRange(linhaExistente, 5, 1, valoresLinha.length).setValues([valoresLinha]);
  } else {
    aba.appendRow([
      chave,
      payload.veiculo || "",
      payload.ponto || "",
      payload.entrada || "",
    ].concat(valoresLinha));
  }

  return { ok: true };
}

// Devolve as excedências marcadas como "ANALISADO" e com um motivo
// preenchido (ou seja: descartadas por justificativa, não por relatório
// gerado sem motivo antigo) — filtra pela data de ENTRADA no mesmo
// range usado por getHistoricoExcesso(). Alimenta o card "Motivos de
// análise" do dashboard em tempo_permanencia.html.
function getJustificativasPermanencia(dataIni, dataFim) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("TEMPO_PERMANENCIA_STATUS");
  if (!aba) return JSON.stringify([]);

  const lastRow = aba.getLastRow();
  if (lastRow < 2) return JSON.stringify([]);

  const data = aba.getRange(2, 1, lastRow - 1, TEMPO_PERMANENCIA_STATUS_HEADER.length).getValues();

  const dtIni = dataIni ? new Date(dataIni + "T00:00:00") : null;
  const dtFim = dataFim ? new Date(dataFim + "T23:59:59") : null;

  const registros = [];
  data.forEach((r) => {
    const chave = String(r[0] || "").trim();
    if (!chave || String(r[4] || "").trim().toUpperCase() !== "ANALISADO") return;

    const motivo = String(r[6] || "").trim();
    if (!motivo) return; // marcado analisado sem motivo (dado antigo, anterior à justificativa)

    const dataEntrada = _normData(r[3]);
    if (dtIni && dtFim) {
      if (!dataEntrada || dataEntrada < dtIni || dataEntrada > dtFim) return;
    }

    registros.push({
      chave: chave,
      veiculo: String(r[1] || "").trim(),
      ponto: String(r[2] || "").trim(),
      entrada: dataEntrada
        ? Utilities.formatDate(dataEntrada, Session.getScriptTimeZone(), "yyyy-MM-dd")
        : String(r[3] || ""),
      motivo: motivo,
      qtdEmbarque: Number(r[7]) || 0,
      qtdDesembarque: Number(r[8]) || 0,
      apoioRodoviaria: String(r[9] || "").trim(),
      observacao: String(r[10] || "").trim(),
    });
  });

  return JSON.stringify(registros);
}

// ── 5c. Histórico de excedências — Tempo de Permanência ────────
// Aba "HISTORICO_EXCESSO": um registro por relatório gerado a
// partir de uma excedência (não por excedência apenas detectada).
// Alimenta o dashboard da tela tempo_permanencia.html.
var HISTORICO_EXCESSO_HEADER = [
  "Chave", "Data", "Veiculo", "Linha", "Ponto", "Cidade", "UF", "Regiao",
  "Motorista", "MotoristaCodigo", "Entrada", "Saida", "PermanenciaMin",
  "PermitidoMin", "ExcedenteMin", "OccurrenceId", "Lat", "Lng", "SalvoEm",
];

function _abaHistoricoExcesso(ss) {
  var aba = ss.getSheetByName("HISTORICO_EXCESSO");
  if (!aba) {
    aba = ss.insertSheet("HISTORICO_EXCESSO");
    aba.appendRow(HISTORICO_EXCESSO_HEADER);
  }
  return aba;
}

// payload: { chave, data, veiculo, linha, ponto, cidade, uf, regiao,
//            motorista, motoristaCodigo, entrada, saida, permanenciaMin,
//            permitidoMin, excedenteMin, occurrenceId }
function salvarHistoricoExcesso(payload) {
  payload = payload || {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = _abaHistoricoExcesso(ss);

  aba.appendRow([
    payload.chave || "",
    payload.data || "",
    payload.veiculo || "",
    payload.linha || "",
    payload.ponto || "",
    payload.cidade || "",
    payload.uf || "",
    payload.regiao || "",
    payload.motorista || "",
    payload.motoristaCodigo || "",
    payload.entrada || "",
    payload.saida || "",
    payload.permanenciaMin || 0,
    payload.permitidoMin || 0,
    payload.excedenteMin || 0,
    payload.occurrenceId || "",
    _coordOuNull(payload.lat, 90) || "",
    _coordOuNull(payload.lng, 180) || "",
    new Date(),
  ]);

  return { ok: true };
}

function getHistoricoExcesso(dataIni, dataFim) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("HISTORICO_EXCESSO");
  if (!aba) return JSON.stringify([]);

  var lastRow = aba.getLastRow();
  if (lastRow < 2) return JSON.stringify([]);

  var data = aba.getRange(2, 1, lastRow - 1, HISTORICO_EXCESSO_HEADER.length).getValues();

  var dtIni = dataIni ? new Date(dataIni + "T00:00:00") : null;
  var dtFim = dataFim ? new Date(dataFim + "T23:59:59") : null;

  var registros = [];
  data.forEach(function (r) {
    var chave = String(r[0] || "").trim();
    if (!chave) return;

    var dataBruta = _normData(r[1]);
    if (dtIni && dtFim) {
      if (!dataBruta || dataBruta < dtIni || dataBruta > dtFim) return;
    }

    registros.push({
      chave: chave,
      data: dataBruta ? Utilities.formatDate(dataBruta, Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[1] || ""),
      veiculo: String(r[2] || "").trim(),
      linha: String(r[3] || "").trim(),
      ponto: String(r[4] || "").trim(),
      cidade: String(r[5] || "").trim(),
      uf: String(r[6] || "").trim(),
      regiao: String(r[7] || "").trim(),
      motorista: String(r[8] || "").trim(),
      motoristaCodigo: String(r[9] || "").trim(),
      entrada: String(r[10] || "").trim(),
      saida: String(r[11] || "").trim(),
      permanenciaMin: Number(r[12]) || 0,
      permitidoMin: Number(r[13]) || 0,
      excedenteMin: Number(r[14]) || 0,
      occurrenceId: String(r[15] || "").trim(),
      lat: _coordOuNull(r[16], 90),
      lng: _coordOuNull(r[17], 180),
    });
  });

  return JSON.stringify(registros);
}

// ── 6. Helpers privados ───────────────────────────────────────
// ── Normalização automática de nomes de bases ─────────────────
function _normalizeBase(raw) {
  if (!raw) return "";
  return (
    raw
      // 1. Remove acentos (NFD + strip combining marks)
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      // 2. Maiúsculas e trim
      .toUpperCase()
      .trim()
      // 3. Remove sufixos operacionais comuns
      .replace(
        /\b(CATEDRAL|FILIAL|RODOVIARIA|LTDA|S\.?A\.?|MATRIZ|AGENCIA|AGÊNCIA|TERMINAL|TRM|GARAGE|GARAGEM)\b/g,
        "",
      )
      // 4. Colapsa espaços extras gerados pela remoção
      .replace(/\s{2,}/g, " ")
      .trim()
  );
}

function _buildPcMap(ss) {
  const map = {};
  try {
    const aba = ss.getSheetByName("PC'S NÃO AUTORIZADO");
    if (!aba || aba.getLastRow() < 2) return map;
    aba
      .getRange(2, 1, aba.getLastRow() - 1, 32)
      .getValues()
      .forEach(function (r) {
        const codigo = String(r[0] || "").trim();
        if (!codigo) return;
        map[codigo] = {
          descResumida: String(r[2] || "").trim(),
          descCompleta: String(r[3] || "").trim(),
          uf: String(r[4] || "").trim(),
          regiao: String(r[5] || "").trim(),
          tipo: String(r[7] || "").trim(),
          lat: _coordOuNull(r[30], 90),
          lng: _coordOuNull(r[31], 180),
        };
      });
  } catch (e) {
    Logger.log("Erro ao carregar mapa de PCs: " + e);
  }
  return map;
}

function _normData(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  const s = String(val);
  const parts = s.split("/");
  if (parts.length === 3) {
    const d = new Date(
      parseInt(parts[2]),
      parseInt(parts[1]) - 1,
      parseInt(parts[0]),
    );
    return isNaN(d.getTime()) ? null : d;
  }
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

// ── _coordOuNull(val, limite) ─────────────────────────────────
// Parser de latitude/longitude tolerante a dois problemas encontrados
// na planilha:
// 1. Vírgula decimal (formato BR, ex.: "-15,78") — parseFloat() sozinho
//    para no primeiro caractere não numérico e devolve só "-15".
// 2. Ponto decimal perdido na importação/colagem (ex.: célula com
//    "-15835736" em vez de "-15.835736") — o valor bruto sempre vem
//    MUITO maior que o range geográfico válido (|lat|<=90, |lng|<=180),
//    então achamos a casa decimal certa dividindo por 10 até caber
//    dentro de "limite". Pra um valor já correto (ex.: -15.835736) o
//    loop não roda nenhuma vez.
function _coordOuNull(val, limite) {
  if (val === null || val === undefined || val === "") return null;
  const s = String(val).trim().replace(",", ".");
  let n = parseFloat(s);
  if (isNaN(n) || n === 0) return null;
  while (Math.abs(n) > limite) n = n / 10;
  return n;
}

// ── Backfill: preenche/corrige coluna "Linha" (col J) no Histórico ────
// Roda manualmente (menu Executar > backfillLinhas), quantas vezes
// precisar. Preenche linhas ainda vazias E corrige linhas que já tinham
// "Linha" gravada só com o sentido em branco (bug antigo da API que
// escrevia "codigo|rota|hora|" sem o sentido quando lineLabel vinha
// preenchido — ver occurrences.service.ts, corrigido, mas o que já tinha
// sido gravado no Histórico ficou incompleto pra sempre até rodar isso).
// Busca as ocorrências no Supabase e monta o mesmo formato usado pela
// API ao vivo: "codigo|rota|hora|sentido".
// Chave de cruzamento: data + número do carro (mesma limitação de sempre:
// se o mesmo carro tiver 2 paradas fora no mesmo dia, a última ocorrência
// buscada é quem preenche — o Histórico não guarda qual parada é qual).
function backfillLinhas() {
  var SUPABASE_URL = "https://rolcprjrwxmdibzajvri.supabase.co";
  var SUPABASE_KEY =
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJvbGNwcmpyd3htZGliemFqdnJpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzAyODY0ODIsImV4cCI6MjA4NTg2MjQ4Mn0.VSoKSDQOQHXR6bWHzxetfRwF50KG4I22K1hdcBpDP68";
  var TYPE_CODE = "DESCUMP_OP_PARADA_FORA";
  // A tabela "trips" no Supabase é bloqueada por RLS pra chave anon (só a
  // API, com service-role, lê) — por isso as viagens vêm de cá em vez de
  // direto do Supabase. Mesma URL que o tempo_permanencia.html usa
  // (REPORTS_API_URL). Rota pública, sem auth.
  var TRIPS_API_URL = "https://novel-trina-luccasinaacio-17a13a2a.koyeb.app/trips";

  function supaGet(path) {
    var res = UrlFetchApp.fetch(SUPABASE_URL + "/rest/v1/" + path, {
      headers: { apikey: SUPABASE_KEY, Authorization: "Bearer " + SUPABASE_KEY },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() >= 300) {
      throw new Error("Supabase " + res.getResponseCode() + ": " + res.getContentText());
    }
    return JSON.parse(res.getContentText());
  }

  // 1. Busca o type_id do tipo no Supabase
  var types = supaGet("occurrence_types?select=id&code=eq." + TYPE_CODE);
  if (!types.length) throw new Error("Tipo " + TYPE_CODE + " não encontrado.");
  var typeId = types[0].id;

  // 2. Busca todas as ocorrências desse tipo com trip_id preenchido.
  // trip_id aqui costuma ser o UUID da viagem canônica (tabela "trips") —
  // resolvido à parte no passo 3, via API (ver TRIPS_API_URL acima).
  var occs = supaGet(
    "occurrences?select=event_date,vehicle_number,trip_id" +
      "&type_id=eq." + typeId +
      "&trip_id=not.is.null" +
      "&limit=5000",
  );

  // 3. trip_id vem misturado: ocorrências antigas guardam o texto já
  // pronto direto em trip_id ("codigo|rota|hora|sentido" — modelo
  // anterior ao join com "trips"); as mais novas guardam o UUID da
  // viagem canônica, que resolvemos aqui via API (não dá pra ler "trips"
  // direto do Supabase com a chave anon — RLS bloqueia).
  var UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  var tripsRes = UrlFetchApp.fetch(TRIPS_API_URL, { muteHttpExceptions: true });
  if (tripsRes.getResponseCode() >= 300) {
    throw new Error("API de viagens " + tripsRes.getResponseCode() + ": " + tripsRes.getContentText());
  }
  var tripMap = {};
  (JSON.parse(tripsRes.getContentText()).data || []).forEach(function (t) {
    tripMap[t.id] = t;
  });

  // 4. Monta mapa: "YYYY-MM-DD|CARRO" → "codigo|rota|hora|sentido"
  // (mesmo formato montado ao vivo em occurrences.service.ts)
  var mapaTrip = {};
  occs.forEach(function (o) {
    var data = String(o.event_date || "").substring(0, 10); // YYYY-MM-DD
    var carro = String(o.vehicle_number || "").trim();
    if (!data || !carro || !o.trip_id) return;

    var linhaStr;
    if (UUID_RE.test(o.trip_id)) {
      var trip = tripMap[o.trip_id];
      if (!trip) return;
      linhaStr = [
        trip.lineCode || "",
        trip.lineName || "",
        (trip.departureTime || "").slice(0, 5),
        String(trip.direction || "").toUpperCase(),
      ].join("|");
    } else {
      // trip_id já é o texto pronto (ocorrência antiga)
      linhaStr = String(o.trip_id).trim();
    }
    mapaTrip[data + "|" + carro] = linhaStr;
  });

  Logger.log("Ocorrências Supabase carregadas: " + occs.length);
  Logger.log("Chaves no mapa: " + Object.keys(mapaTrip).length);

  // 5. Lê o Histórico
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hist = ss.getSheetByName("Histórico");
  if (!hist) throw new Error('Aba "Histórico" não encontrada.');

  var lastRow = hist.getLastRow();
  if (lastRow < 2) {
    Logger.log("Histórico vazio.");
    return;
  }

  var range = hist.getRange(2, 1, lastRow - 1, 10);
  var valores = range.getValues();

  var atualizados = 0;
  var corrigidos = 0;

  valores.forEach(function (r, i) {
    var atual = String(r[9] || "").trim();

    // Preenche linhas ainda vazias, e também corrige linhas já preenchidas
    // mas que ficaram com o sentido em branco (bug antigo da API: escrevia
    // "codigo|rota|hora|" sem o 4º campo) — desde que já tenhamos, agora,
    // um valor completo pra mesma data+carro.
    var partesAtual = atual.split("|");
    var faltaSentido = atual !== "" && partesAtual.length >= 4 && !partesAtual[3];
    if (atual !== "" && !faltaSentido) return;

    var dataVal = _normData(r[5]);
    if (!dataVal) return;

    var y = dataVal.getFullYear();
    var m = String(dataVal.getMonth() + 1).padStart(2, "0");
    var d = String(dataVal.getDate()).padStart(2, "0");
    var dataStr = y + "-" + m + "-" + d;
    var carro = String(r[2] || "").trim();

    var chave = dataStr + "|" + carro;
    var trip = mapaTrip[chave];
    if (!trip) return;

    var partesNovo = trip.split("|");
    var novoTemSentido = partesNovo.length >= 4 && !!partesNovo[3];
    if (faltaSentido && !novoTemSentido) return; // não tem nada melhor pra oferecer

    hist.getRange(i + 2, 10).setValue(trip); // col J
    if (faltaSentido) corrigidos++; else atualizados++;
  });

  Logger.log("Linhas preenchidas (vazias): " + atualizados + " | corrigidas (sentido que faltava): " + corrigidos + " de " + (lastRow - 1));
  SpreadsheetApp.getUi().alert(
    "Backfill concluído: " + atualizados + " preenchida(s) + " + corrigidos + " corrigida(s) (sentido).",
  );
}
