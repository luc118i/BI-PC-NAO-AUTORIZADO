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
];

// Antes disto, marcarStatusPermanencia() lançava erro se a aba não
// existisse — como ela nunca é criada manualmente, TODA marcação de
// "analisado" falhava silenciosamente (o erro só ia pro console via
// withFailureHandler) e nada era persistido, por isso sumia após F5.
// Agora a aba é criada sozinha aqui, igual _abaHistoricoExcesso().
function _abaStatusPermanencia(ss) {
  var aba = ss.getSheetByName("TEMPO_PERMANENCIA_STATUS");
  if (!aba) {
    aba = ss.insertSheet("TEMPO_PERMANENCIA_STATUS");
    aba.appendRow(TEMPO_PERMANENCIA_STATUS_HEADER);
  }
  return aba;
}

function getStatusPermanencia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName("TEMPO_PERMANENCIA_STATUS");
  if (!aba) return JSON.stringify([]);

  const lastRow = aba.getLastRow();
  if (lastRow < 2) return JSON.stringify([]);

  const data = aba.getRange(2, 1, lastRow - 1, 6).getValues();
  const chaves = data
    .filter(
      (r) =>
        String(r[0] || "").trim() !== "" &&
        String(r[4] || "").trim().toUpperCase() === "ANALISADO",
    )
    .map((r) => String(r[0]).trim());

  return JSON.stringify(chaves);
}

// payload: { chave, veiculo, ponto, entrada, analisado }
function marcarStatusPermanencia(payload) {
  const chave = String((payload && payload.chave) || "").trim();
  if (!chave) throw new Error("chave é obrigatória.");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const aba = _abaStatusPermanencia(ss);

  const status = payload.analisado ? "ANALISADO" : "";
  const agora = new Date();

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
    aba.getRange(linhaExistente, 5, 1, 2).setValues([[status, agora]]);
  } else {
    aba.appendRow([
      chave,
      payload.veiculo || "",
      payload.ponto || "",
      payload.entrada || "",
      status,
      agora,
    ]);
  }

  return { ok: true };
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

// ── Backfill: preenche coluna "Linha" (col J) no Histórico ────
// Roda manualmente (menu Executar > backfillLinhas), quantas vezes
// precisar — só toca linhas do Histórico ainda sem "Linha" preenchida.
// Busca as ocorrências no Supabase e monta o mesmo formato usado pela
// API ao vivo (occurrences.service.ts): "codigo|rota|hora|sentido".
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

  valores.forEach(function (r, i) {
    // Só preenche linhas que ainda não têm trip_id (col J = índice 9)
    if (String(r[9] || "").trim() !== "") return;

    var dataVal = _normData(r[5]);
    if (!dataVal) return;

    var y = dataVal.getFullYear();
    var m = String(dataVal.getMonth() + 1).padStart(2, "0");
    var d = String(dataVal.getDate()).padStart(2, "0");
    var dataStr = y + "-" + m + "-" + d;
    var carro = String(r[2] || "").trim();

    var chave = dataStr + "|" + carro;
    var trip = mapaTrip[chave];

    if (trip) {
      hist.getRange(i + 2, 10).setValue(trip); // col J
      atualizados++;
    }
  });

  Logger.log("Linhas atualizadas: " + atualizados + " de " + (lastRow - 1));
  SpreadsheetApp.getUi().alert(
    "Backfill concluído: " + atualizados + " linha(s) preenchida(s).",
  );
}
