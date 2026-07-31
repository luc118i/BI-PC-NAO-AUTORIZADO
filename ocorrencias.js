/** ==============================
 *  FORMULÁRIO DE OCORRÊNCIAS
 *  ============================== */
const OCORRENCIA_CFG = Object.freeze({
  LOCAIS_SHEET: "PC'S NÃO AUTORIZADO",
  MOTORISTAS_SHEET: "Motorista",
  HISTORY_SHEET: "Histórico",
  HEADER_ROWS: 1,
});

/** Abre o formulário (sidebar) */
function openSidebarOcorrencias() {
  const html = HtmlService.createHtmlOutputFromFile("Form")
    .setTitle("Registrar Ocorrência")
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** ==============================
 *  LISTAR LOCAIS
 *  ============================== */
function getLocais() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(OCORRENCIA_CFG.LOCAIS_SHEET);
  if (!sh)
    throw new Error(`Aba "${OCORRENCIA_CFG.LOCAIS_SHEET}" não encontrada.`);

  const lastRow = sh.getLastRow();
  if (lastRow <= OCORRENCIA_CFG.HEADER_ROWS) return [];

  const data = sh
    .getRange(
      OCORRENCIA_CFG.HEADER_ROWS + 1,
      1,
      lastRow - OCORRENCIA_CFG.HEADER_ROWS,
      4,
    )
    .getValues(); // A:D

  return data
    .filter((r) => r[0] && r[3])
    .map((r) => ({ id: String(r[0]).trim(), nome: String(r[3]).trim() }));
}

/** Retorna a última linha com dado em uma coluna específica (ignorando linhas só formatadas). */
function getLastDataRow_(sheet, colIndex, headerRows) {
  const startRow = (headerRows || 1) + 1;
  const maxRow = sheet.getMaxRows();
  if (maxRow < startRow) return headerRows || 1;

  const numRows = Math.max(0, lastRow - OCORRENCIA_CFG.HEADER_ROWS);
  const values = sheet.getRange(startRow, colIndex, numRows, 1).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const v = values[i][0];
    if (v !== "" && v !== null) {
      return startRow + i; // linha real com conteúdo
    }
  }
  return headerRows || 1; // só tem cabeçalho
}

/** ==============================
 *  LISTAR MOTORISTAS
 *  ============================== */
function getMotoristas() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(OCORRENCIA_CFG.MOTORISTAS_SHEET);
  if (!sh)
    throw new Error(`Aba "${OCORRENCIA_CFG.MOTORISTAS_SHEET}" não encontrada.`);

  const lastRow = sh.getLastRow();
  if (lastRow <= OCORRENCIA_CFG.HEADER_ROWS) return [];

  // Lê colunas C (nome) e D (matrícula)
  const data = sh
    .getRange(
      OCORRENCIA_CFG.HEADER_ROWS + 1,
      3,
      lastRow - OCORRENCIA_CFG.HEADER_ROWS,
      2,
    )
    .getValues(); // [[nome, matricula], ...]

  return data
    .filter((r) => r[0] && r[1])
    .map((r) => ({
      nome: String(r[0]).trim(),
      id: String(r[1]).trim(),
    }));
}

/** ==============================
 *  LISTAR BASES DOS MOTORISTAS
 *  ============================== */
function getBasesMotorista() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(OCORRENCIA_CFG.MOTORISTAS_SHEET);
    if (!sh)
      throw new Error(
        `Aba "${OCORRENCIA_CFG.MOTORISTAS_SHEET}" não encontrada.`,
      );

    const colT = sh
      .getRange("T:T")
      .getValues()
      .flat()
      .map((b) => String(b).trim())
      .filter((b) => b);
    if (colT.length === 0) return [];

    const unicas = [...new Set(colT)].sort((a, b) => a.localeCompare(b));
    Logger.log(`Bases carregadas: ${JSON.stringify(unicas)}`);
    return unicas.map((base) => ({ nome: base }));
  } catch (e) {
    Logger.log("Erro em getBasesMotorista: " + e);
    throw new Error(
      "Erro ao carregar lista de bases. Verifique a aba Motorista.",
    );
  }
}

/** ==============================
 *  ADICIONAR NOVA OCORRÊNCIA
 *  ============================== */
function addOcorrencia(payload) {
  const {
    localId,
    localNome,
    carro,
    motoristaId,
    motoristaNome,
    base,
    dataRelatorio,
    linha,
  } = payload || {};
  if (!localNome || !carro || !motoristaId || !motoristaNome)
    throw new Error("Preencha todos os campos.");

  const ss = SpreadsheetApp.getActive();
  const shHist =
    ss.getSheetByName(OCORRENCIA_CFG.HISTORY_SHEET) ||
    ss.insertSheet(OCORRENCIA_CFG.HISTORY_SHEET);

  if (shHist.getLastRow() === 0) {
    shHist.appendRow([
      "ID Local",
      "Local",
      "Carro",
      "ID Motorista",
      "Motorista",
      "Data do Relatório",
      "Base",
    ]);
  }

  // Colunas A–G (H=UF e I=Região são fórmulas — não sobrescrever)
  shHist.appendRow([
    localId,
    localNome,
    carro,
    motoristaId,
    motoristaNome,
    dataRelatorio,
    base ?? "",
  ]);

  // Coluna J (10) = Linha — escrita separada para não deslocar H e I
  shHist.getRange(shHist.getLastRow(), 10).setValue(linha ?? "");

  return { ok: true };
}

/** ==============================
 *  ADICIONAR NOVO MOTORISTA
 *  ============================== */
function addMotorista(payload) {
  const { id, nome, base } = payload || {};
  if (!id || !nome || !base)
    throw new Error("Preencha ID, Nome e Base do motorista.");

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(OCORRENCIA_CFG.MOTORISTAS_SHEET);
  if (!sh)
    throw new Error(`Aba "${OCORRENCIA_CFG.MOTORISTAS_SHEET}" não encontrada.`);

  const lastRow = sh.getLastRow();
  const matriculas = sh
    .getRange(2, 4, Math.max(0, lastRow - 1))
    .getValues()
    .flat()
    .map(String);

  if (matriculas.includes(String(id))) {
    throw new Error("Essa matrícula (ID do motorista) já existe.");
  }

  const row = sh.getLastRow() + 1;
  sh.getRange(row, 3).setValue(nome); // Nome (C)
  sh.getRange(row, 4).setValue(id); // ID (D)
  sh.getRange(row, 20).setValue(base); // Base (T)

  return { ok: true };
}

/** ==============================
 *  ADICIONAR NOVOS PONTOS DE PARADA IRREGULAR (via CSV)
 *  ============================== */
function _normCodigoPonto_(v) {
  return String(v || "").trim().toUpperCase();
}

function addPontosParadaIrregular(rows) {
  if (!Array.isArray(rows) || !rows.length)
    throw new Error("Nenhum registro para importar.");

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(OCORRENCIA_CFG.LOCAIS_SHEET);
  if (!sh)
    throw new Error(`Aba "${OCORRENCIA_CFG.LOCAIS_SHEET}" não encontrada.`);

  const lastRow = sh.getLastRow();
  const existentes = {};
  if (lastRow >= OCORRENCIA_CFG.HEADER_ROWS + 1) {
    sh.getRange(OCORRENCIA_CFG.HEADER_ROWS + 1, 1, lastRow - OCORRENCIA_CFG.HEADER_ROWS, 1)
      .getValues()
      .forEach((r) => {
        const cod = _normCodigoPonto_(r[0]);
        if (cod) existentes[cod] = true;
      });
  }

  const added = [];
  const skipped = [];
  let nextRow = lastRow + 1;

  rows.forEach((payload) => {
    const codigo = _normCodigoPonto_(payload && payload.codigo);
    const descricao = String((payload && payload.descricao) || "").trim();
    const lat = Number(payload && payload.latitude);
    const lng = Number(payload && payload.longitude);

    if (!codigo || !descricao || !isFinite(lat) || !isFinite(lng)) {
      skipped.push({ codigo: codigo || "(vazio)", motivo: "Dados obrigatórios ausentes ou inválidos." });
      return;
    }
    if (existentes[codigo]) {
      skipped.push({ codigo: codigo, motivo: "Código já cadastrado na base." });
      return;
    }

    sh.getRange(nextRow, 1, 1, 4).setValues([[
      payload.codigo,
      payload.codEmb || "",
      payload.descResumida || "",
      descricao,
    ]]);
    sh.getRange(nextRow, 7, 1, 2).setValues([[
      payload.unidadeEmpresarial || "",
      payload.tipo || "",
    ]]);
    sh.getRange(nextRow, 31, 1, 2).setValues([[lat, lng]]);

    existentes[codigo] = true;
    added.push(codigo);
    nextRow++;
  });

  return { ok: true, added: added.length, skipped: skipped };
}

/** ==============================
 *  WEB APP — recebe POST externo
 *  ============================== */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    if (
      payload.token !==
      "CSg1u5eG8mqbAf3IPH4XIgRLahxh71etkACmh711kp8PIQL64rUiG3JmWrxihwwU"
    ) {
      return ContentService.createTextOutput(
        JSON.stringify({ error: "Unauthorized" }),
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (payload.evento === "excesso_permanencia") {
      salvarHistoricoExcesso(payload);
    } else {
      addOcorrencia({
        localId: payload.localId,
        localNome: payload.localNome,
        carro: payload.carro,
        motoristaId: payload.motoristaId,
        motoristaNome: payload.motoristaNome,
        base: payload.base,
        dataRelatorio: payload.dataRelatorio,
        linha: payload.linha,
      });
    }

    return ContentService.createTextOutput(
      JSON.stringify({ ok: true }),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: err.message }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
