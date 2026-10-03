// SiM-Guru — modular web app / HTML loader
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
function getPageContent(pageName) {
  var normalized = pageName === "Page_StudentAttendanceInput" ? "Page_StudentAttendance" : pageName;
  try {
    var html = include(normalized);
    return html
      .replace(/__SIM_GURU_LAYOUT_SIDEBAR__/g, include("UI_Sidebar"))
      .replace(/__SIM_GURU_LAYOUT_NAVBAR__/g, include("UI_Navbar"))
      .replace(/__SIM_GURU_LAYOUT_FOOTER__/g, include("UI_Footer"))
      .replace(/__SIM_GURU_LAYOUT_MODAL__/g, include("UI_Modal"));
  } catch (e) {
    return "<div class='p-4 text-red-500 font-bold'>Error: " + e.toString() + "</div>";
  }
}
function doGet(e) {
  var template = HtmlService.createTemplateFromFile("Index");
  try {
    var configRaw = getData("Config"), config = {};
    (configRaw || []).forEach(function(c){ config[c.key] = c.value; });
    template.serverConfig = JSON.stringify(config);
  } catch(err) { template.serverConfig = "{}"; }
  return template.evaluate().setTitle("SiM-Guru")
    .addMetaTag("viewport","width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
