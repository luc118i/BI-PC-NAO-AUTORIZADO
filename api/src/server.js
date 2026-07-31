require("dotenv").config({ path: require("path").join(__dirname, "../.env") });

const express = require("express");
const cors = require("cors");
const escalaRouter    = require("./routes/escala");
const frotaRouter     = require("./routes/frota");
const relatorioRouter = require("./routes/relatorio");
const optzRouter      = require("./routes/optz");

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors({ origin: "*" })); // restrinja para o domínio GAS em produção
app.use(express.json({ limit: "2mb" }));

app.get("/health", (_req, res) => res.json({ ok: true }));

// POST /api/escala   — escala de motorista (Tampermonkey / Tripulantes)
app.use("/api/escala", escalaRouter);

// POST /api/frota    — dados de veículo (Tampermonkey / Frotas)
app.use("/api/frota", frotaRouter);

// POST /api/relatorio — monta relatório completo (imagens + escala)
app.use("/api/relatorio", relatorioRouter);

// /api/optz/* — interceptor OPTZ v3 (JWT + servicos)
app.use("/api/optz", optzRouter);

app.listen(PORT, "127.0.0.1", () => {
  console.log(`[API] Ocorrências Inteligentes rodando em http://localhost:${PORT}`);
});
