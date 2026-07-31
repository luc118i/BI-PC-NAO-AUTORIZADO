const { Router } = require("express");
const store      = require("../store");

const router = Router();

// ── Rotas ────────────────────────────────────────────────────

// POST /api/optz/auth
router.post("/auth", (req, res) => {
  const { jwt } = req.body;
  if (!jwt) return res.status(400).json({ error: "jwt obrigatório" });
  store.setJwt(jwt);
  console.log("[optz] JWT atualizado");
  res.json({ ok: true });
});

// POST /api/optz/empresa — armazena UUID da empresa capturado de qualquer URL do OPTZ
router.post("/empresa", (req, res) => {
  const { uuid } = req.body;
  if (!uuid) return res.status(400).json({ error: "uuid obrigatório" });
  store.setEmpresaUuid(uuid);
  res.json({ ok: true });
});

// POST /api/optz/servicos
router.post("/servicos", (req, res) => {
  const { area, inicio, fim, trilhos, empresaUuid } = req.body;
  if (!Array.isArray(trilhos)) return res.status(400).json({ error: "trilhos[] obrigatório" });

  if (empresaUuid) store.setEmpresaUuid(empresaUuid);
  store.setServicosOptz(area, inicio, fim, trilhos);

  const label = area === "2" ? "frotas" : "tripulantes";
  console.log(`[optz] Servicos ${label}: ${trilhos.length} trilhos (${inicio} → ${fim})`);
  res.json({ ok: true, area, count: trilhos.length });
});

// POST /api/optz/grupos
router.post("/grupos", (req, res) => {
  const grupos = req.body;
  if (!Array.isArray(grupos)) return res.status(400).json({ error: "array obrigatório" });
  const updated = store.updateGruposMeta(grupos);
  console.log(`[optz] grupos meta: ${updated} atualizados`);
  res.json({ ok: true, updated });
});

// GET /api/optz/motorista?servico=X&data=YYYY-MM-DDTHH:MM
router.get("/motorista", (req, res) => {
  const { servico, data } = req.query;
  if (!servico) return res.status(400).json({ error: "servico obrigatório" });

  const resultado = store.buscarMotoristaPorServicoOptz(servico, data);
  if (!resultado) return res.status(404).json({ error: "not_cached" });
  res.json(resultado);
});

// GET /api/optz/frota?prefixo=X&data=YYYY-MM-DDTHH:MM
router.get("/frota", (req, res) => {
  const { prefixo, data } = req.query;
  if (!prefixo) return res.status(400).json({ error: "prefixo obrigatório" });

  const resultado = store.buscarServicoPorFrota(prefixo, data);
  if (!resultado) return res.status(404).json({ error: "not_cached" });
  res.json(resultado);
});

// GET /api/optz/status/dia?data=YYYY-MM-DD
router.get("/status/dia", (req, res) => {
  const { data } = req.query;
  if (!data) return res.status(400).json({ error: "data obrigatória" });
  const status = store.temServicosParaDia(data);
  res.json({ ...status, data: data.slice(0, 10) });
});

// GET /api/optz/status
router.get("/status", (_req, res) => {
  const s = store.statusOptz();
  res.json({ ...s, empresaUuid: store.getEmpresaUuid()?.slice(0, 8) + "..." });
});

// GET /api/optz/debug/frota?prefixo=X
router.get("/debug/frota", (req, res) => {
  const { prefixo } = req.query;
  const status = store.statusOptz();
  if (prefixo) {
    const matches  = status.frotas.filter(f =>
      String(f.prefixo || "").includes(prefixo) || String(f.uuid).includes(prefixo)
    );
    const detalhes = matches.map(f => store.getFrotaDetalhes(f.uuid));
    return res.json({ query: prefixo, matches: detalhes, total: status.frotas.length });
  }
  const comPrefixo = status.frotas.filter(f => f.prefixo).slice(0, 10);
  const semPrefixo = status.frotas.filter(f => !f.prefixo).length;
  res.json({ comPrefixo, semPrefixo, total: status.frotas.length });
});

module.exports = router;
