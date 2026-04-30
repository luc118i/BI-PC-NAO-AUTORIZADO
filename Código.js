/** ==============================
 *  MENU PRINCIPAL
 *  ============================== */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Painel 📋")
    .addSubMenu(
      ui
        .createMenu("Ocorrências 🚐")
        .addItem("Registrar ocorrência", "openSidebarOcorrencias")
        .addItem("Abrir histórico", "abrirHistorico"),
    )
    .addSeparator()
    .addItem("📊 Abrir Dashboard", "abrirDashboard")
    .addToUi();
}

/** Abre a aba "Histórico" direto */
function abrirHistorico() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Histórico");
  if (sh) ss.setActiveSheet(sh);
  else SpreadsheetApp.getUi().alert('A aba "Histórico" ainda não existe.');
}

/** Abre a dashboard direto no navegador, sem alerta */
function abrirDashboard() {
  const url = "https://lookerstudio.google.com/s/u5nH3lsAuXg";
  const html = HtmlService.createHtmlOutput(
    `
    <script>
      window.open('${url}', '_blank');
      google.script.host.close();
    </script>
  `,
  )
    .setWidth(1)
    .setHeight(1);

  SpreadsheetApp.getUi().showModelessDialog(html, " ");
}
