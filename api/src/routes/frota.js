const { Router } = require("express");
const store = require("../store");

const router = Router();

// POST /api/frota — recebe dados de veículo capturados do OPTZ Frotas
router.post("/", (req, res) => {
  const payload = req.body;

  if (!payload || !payload.prefixo) {
    return res.status(400).json({ error: "Campo 'prefixo' obrigatório." });
  }

  store.setFrota(payload.prefixo, payload);

  console.log(
    `[frota] Recebido prefixo ${payload.prefixo} (${payload.linha || "—"})` +
    ` · ${payload.servicos?.length ?? 0} blocos`
  );

  res.json({ ok: true, prefixo: payload.prefixo });
});

// GET /api/frota/:prefixo — consulta dados de veículo
router.get("/:prefixo", (req, res) => {
  const frota = store.getFrota(req.params.prefixo);
  if (!frota) {
    return res.status(404).json({ error: "Frota não encontrada para este prefixo." });
  }
  res.json(frota);
});

// GET /api/frota — lista todos os prefixos em memória
router.get("/", (_req, res) => {
  res.json(store.listFrotas());
});

module.exports = router;
