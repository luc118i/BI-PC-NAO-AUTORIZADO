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

  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("BI PC's Não Autorizados · Viação Catedral")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
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
      lat: _numOuNull(r[30]),
      lng: _numOuNull(r[31]),
    }));

  return JSON.stringify(pcs);
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
          lat: _numOuNull(r[30]),
          lng: _numOuNull(r[31]),
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

function _numOuNull(val) {
  if (val === null || val === undefined || val === "") return null;
  const n = parseFloat(val);
  return isNaN(n) || n === 0 ? null : n;
}

// ── Backfill: preenche coluna "Linha" (col J) no Histórico ────
// Roda UMA vez manualmente. Busca trip_id do Supabase e cruza
// com as linhas do Histórico que ainda estão sem linha preenchida.
// Chave de cruzamento: data + número do carro.
function backfillLinhas() {
  var SUPABASE_URL = "https://rolcprjrwxmdibzajvri.supabase.co";
  var SUPABASE_KEY =
    "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJvbGNwcmpyd3htZGliemFqdnJpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzAyODY0ODIsImV4cCI6MjA4NTg2MjQ4Mn0.VSoKSDQOQHXR6bWHzxetfRwF50KG4I22K1hdcBpDP68";
  var TYPE_CODE = "DESCUMP_OP_PARADA_FORA";

  // 1. Busca o type_id do tipo no Supabase
  var typeRes = UrlFetchApp.fetch(
    SUPABASE_URL + "/rest/v1/occurrence_types?select=id&code=eq." + TYPE_CODE,
    {
      headers: {
        apikey: SUPABASE_KEY,
        Authorization: "Bearer " + SUPABASE_KEY,
      },
    },
  );
  var types = JSON.parse(typeRes.getContentText());
  if (!types.length) throw new Error("Tipo " + TYPE_CODE + " não encontrado.");
  var typeId = types[0].id;

  // 2. Busca todas as ocorrências desse tipo com trip_id preenchido
  var occRes = UrlFetchApp.fetch(
    SUPABASE_URL +
      "/rest/v1/occurrences" +
      "?select=event_date,vehicle_number,trip_id" +
      "&type_id=eq." +
      typeId +
      "&trip_id=not.is.null" +
      "&limit=5000",
    {
      headers: {
        apikey: SUPABASE_KEY,
        Authorization: "Bearer " + SUPABASE_KEY,
      },
    },
  );
  var occs = JSON.parse(occRes.getContentText());

  // 3. Monta mapa: "YYYY-MM-DD|CARRO" → trip_id
  var mapaTrip = {};
  occs.forEach(function (o) {
    var data = String(o.event_date || "").substring(0, 10); // YYYY-MM-DD
    var carro = String(o.vehicle_number || "").trim();
    var trip = String(o.trip_id || "").trim();
    if (data && carro && trip) {
      mapaTrip[data + "|" + carro] = trip;
    }
  });

  Logger.log("Ocorrências Supabase carregadas: " + occs.length);
  Logger.log("Chaves no mapa: " + Object.keys(mapaTrip).length);

  // 4. Lê o Histórico
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
