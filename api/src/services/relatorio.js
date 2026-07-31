// Gera relatório em HTML standalone (imprimível / salvável como PDF via Ctrl+P)

function montarRelatorio({ ocorrencia: oc, imagens, escala }) {
  const agora = new Date().toLocaleString("pt-BR", { timeZone: "America/Sao_Paulo" });
  const dataOc = oc.entrada
    ? new Date(oc.entrada).toLocaleDateString("pt-BR")
    : oc.dataRelatorio || "—";

  const imgSat = imagens?.satelite
    ? `<img src="${imagens.satelite}" alt="Satélite">`
    : `<div class="img-placeholder">Imagem de satélite indisponível</div>`;

  const imgSV = imagens?.streetView && imagens?.svDisponivel
    ? `<img src="${imagens.streetView}" alt="Street View">`
    : `<div class="img-placeholder">Street View indisponível para este local</div>`;

  const escalaRows = escala?.escala?.length
    ? escala.escala.map((s) => `
        <tr>
          <td>${s.dtInicio || "—"}</td>
          <td>${s.tipo || "—"}</td>
          <td>${s.numero || s.duracao || "—"}</td>
          <td>${s.rota || "—"}</td>
          <td>${s.horario || "—"}</td>
        </tr>`).join("")
    : `<tr><td colspan="5" style="text-align:center;color:#888">Escala não capturada</td></tr>`;

  return /* html */ `<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Ocorrência · ${oc.veiculo} · ${dataOc}</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: Arial, sans-serif; font-size: 12px; color: #1a1a1a; background: #fff; padding: 24px; }
  h1 { font-size: 18px; font-weight: 700; color: #c0392b; margin-bottom: 4px; }
  .sub { font-size: 11px; color: #666; margin-bottom: 20px; }
  section { margin-bottom: 20px; }
  h2 { font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; color: #555; border-bottom: 1px solid #ddd; padding-bottom: 4px; margin-bottom: 10px; }
  .grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 8px 16px; }
  .field label { font-size: 10px; font-weight: 700; text-transform: uppercase; color: #888; }
  .field span { display: block; font-size: 13px; font-weight: 600; color: #111; margin-top: 1px; }
  .dur { color: #c0392b; }
  .imgs { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
  .imgs img { width: 100%; border-radius: 4px; border: 1px solid #ddd; display: block; }
  .img-placeholder { background: #f4f4f4; border: 1px solid #ddd; border-radius: 4px; height: 160px; display: flex; align-items: center; justify-content: center; color: #999; font-size: 11px; }
  table { width: 100%; border-collapse: collapse; font-size: 11px; }
  th { text-align: left; font-size: 10px; text-transform: uppercase; letter-spacing: 0.06em; color: #888; padding: 5px 8px; border-bottom: 2px solid #ddd; }
  td { padding: 5px 8px; border-bottom: 1px solid #eee; color: #333; }
  tr:last-child td { border-bottom: none; }
  .footer { margin-top: 24px; font-size: 10px; color: #aaa; text-align: right; }
  @media print {
    body { padding: 0; }
    @page { margin: 15mm; }
  }
</style>
</head>
<body>

<h1>⚡ Parada Não Autorizada</h1>
<p class="sub">Gerado em ${agora} · BI PC's Não Autorizados · Viação Catedral</p>

<section>
  <h2>Dados da Ocorrência</h2>
  <div class="grid2">
    <div class="field"><label>Veículo</label><span>${oc.veiculo || "—"}</span></div>
    <div class="field"><label>Ponto de controle</label><span>${oc.descResumida || oc.ponto || "—"}</span></div>
    <div class="field"><label>Código PC</label><span>${oc.codigo || "—"}</span></div>
    <div class="field"><label>Região / UF</label><span>${oc.regiao || "—"}${oc.uf ? " · " + oc.uf : ""}</span></div>
    <div class="field"><label>Entrada</label><span>${oc.entrada || "—"}</span></div>
    <div class="field"><label>Saída</label><span>${oc.saida || "—"}</span></div>
    <div class="field"><label>Duração</label><span class="dur">${oc.parada || "—"}</span></div>
    <div class="field"><label>Data do relatório</label><span>${dataOc}</span></div>
  </div>
</section>

<section>
  <h2>Motorista / Base</h2>
  <div class="grid2">
    <div class="field"><label>Motorista</label><span>${oc.motoristaNome || "—"}</span></div>
    <div class="field"><label>Matrícula</label><span>${oc.motoristaId || "—"}</span></div>
    <div class="field"><label>Base</label><span>${oc.base || "—"}</span></div>
    <div class="field"><label>Linha / Serviço</label><span>${oc.linha || "—"}</span></div>
  </div>
</section>

<section>
  <h2>Local — Satélite &amp; Street View</h2>
  <div class="imgs">
    ${imgSat}
    ${imgSV}
  </div>
  ${oc.lat ? `<p style="margin-top:6px;font-size:10px;color:#888">Coordenadas: ${oc.lat}, ${oc.lng}</p>` : ""}
</section>

<section>
  <h2>Escala do motorista${escala ? ` — ${escala.nome || escala.matricula}` : ""}</h2>
  <table>
    <thead>
      <tr><th>Horário</th><th>Tipo</th><th>Serviço / Duração</th><th>Rota</th><th>Saída</th></tr>
    </thead>
    <tbody>${escalaRows}</tbody>
  </table>
</section>

<div class="footer">BI PC's Não Autorizados · Viação Catedral · ${agora}</div>

</body>
</html>`;
}

module.exports = { montarRelatorio };
