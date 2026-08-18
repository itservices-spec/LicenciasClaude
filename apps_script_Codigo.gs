// Apps Script — sirve el dashboard para incrustarlo en Google Sites.
// 1) Crea un proyecto en https://script.google.com
// 2) Pega este código en "Codigo.gs".
// 3) Crea un archivo HTML llamado "dashboard" (Archivo > Nuevo > HTML) y pega
//    ahí TODO el contenido de dashboard_google_site.html.
// 4) Implementar > Nueva implementación > Aplicación web:
//       - Ejecutar como: Yo
//       - Quién tiene acceso: Cualquiera / Cualquiera de tu organización
//    Copia la URL que termina en /exec
// 5) En Google Sites: Insertar > Insertar > Por URL > pega la URL /exec
function doGet() {
  return HtmlService.createHtmlOutputFromFile('dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('Dashboard de Usabilidad · Licencias Claude')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
