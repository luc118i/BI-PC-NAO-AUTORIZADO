const { Router } = require("express");
const store = require("../store");

const router = Router();

// POST /api/escala
// Recebe escala capturada pelo Tampermonkey no OPTZ e a guarda em memória.
// O endpoint GET permite que o BI consulte a escala mais recente por matrícula.
router.post("/", (req, res) => {
  const payload = req.body;

  if (!payload || !payload.matricula) {
    return res.status(400).json({ error: "Campo 'matricula' obrigatório." });
  }

  store.setEscala(payload.matricula, payload);

  console.log(
    `[escala] Recebida escala de ${payload.nome} (${payload.matricula}) ` +
    `· ${payload.escala?.length ?? 0} blocos`
  );

  res.json({ ok: true, matricula: payload.matricula });
});

// GET /api/escala/por-servico/:numero — cross-reference: qual motorista tem este serviço?
router.get("/por-servico/:numero", (req, res) => {
  const resultado = store.buscarMotoristaPorServico(req.params.numero);
  if (!resultado) {
    return res.status(404).json({ error: "Nenhum motorista com este serviço capturado." });
  }
  res.json(resultado);
});

// GET /api/escala/:matricula — consulta escala guardada
router.get("/:matricula", (req, res) => {
  const escala = store.getEscala(req.params.matricula);
  if (!escala) {
    return res.status(404).json({ error: "Escala não encontrada para esta matrícula." });
  }
  res.json(escala);
});

// GET /api/escala — lista todas as matrículas com escala em memória
router.get("/", (_req, res) => {
  res.json(store.listEscalas());
});

module.exports = router;
