const { Router } = require("express");
const { buscarImagensMaps } = require("../services/maps");
const { montarRelatorio } = require("../services/relatorio");
const store = require("../store");

const router = Router();

// POST /api/relatorio
// Body: { ocorrencia, matricula }
//   ocorrencia: objeto com veiculo, ponto, codigo, lat, lng, entrada, saida, parada, regiao, uf, base, motorista...
//   matricula:  string — busca escala em memória (capturada pelo Tampermonkey)
router.post("/", async (req, res) => {
  const { ocorrencia, matricula } = req.body || {};

  if (!ocorrencia) {
    return res.status(400).json({ error: "Campo 'ocorrencia' obrigatório." });
  }

  try {
    // 1. Imagens do local (satélite + street view) via Maps Static API
    const imagens = await buscarImagensMaps(ocorrencia.lat, ocorrencia.lng);

    // 2. Escala do motorista (capturada pelo Tampermonkey e guardada na memória)
    const escala = matricula ? store.getEscala(matricula) : null;

    // 3. Monta relatório HTML com todos os dados
    const html = montarRelatorio({ ocorrencia, imagens, escala });

    res.setHeader("Content-Type", "text/html; charset=utf-8");
    res.send(html);
  } catch (err) {
    console.error("[relatorio] Erro:", err);
    res.status(500).json({ error: err.message });
  }
});

module.exports = router;
