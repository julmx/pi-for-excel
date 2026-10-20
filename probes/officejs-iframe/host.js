const PARENT = "https://127.0.0.1:3151";
const RUNNER = "https://localhost:3152";
const SHEET = "__PiIframeProbe";
const CANARY = "__PiIframeSyntheticCanary";
const status = document.querySelector("#status");
const prepare = document.querySelector("#prepare");
const execute = document.querySelector("#execute");
const frame = document.querySelector("#runner");
let childReady = false;
let prepared = false;
let pending = null;

function report(message) {
  status.textContent += message + "\n";
}

function makeCanary() {
  return new Promise((resolve, reject) => {
    const open = indexedDB.open(CANARY, 1);
    open.onupgradeneeded = () => {
      if (!open.result.objectStoreNames.contains("values")) {
        open.result.createObjectStore("values");
      }
    };
    open.onerror = () => reject(open.error);
    open.onsuccess = () => {
      const db = open.result;
      const tx = db.transaction("values", "readwrite");
      tx.objectStore("values").put("SYNTHETIC_NOT_A_SECRET", "marker");
      tx.oncomplete = () => { db.close(); resolve(); };
      tx.onerror = () => { db.close(); reject(tx.error); };
    };
  });
}

function clearCanary() {
  indexedDB.deleteDatabase(CANARY);
}

const code = [
  "const sheet = context.workbook.worksheets.getItem('__PiIframeProbe');",
  "const data = sheet.getRange('A1:C4');",
  "data.load('values');",
  "await context.sync();",
  "if (data.values[0][0] !== 'Item' || data.values[0][2] !== 'Score' || data.values[3][2] !== 30) {",
  "  throw new Error('Synthetic input does not match expected probe data');",
  "}",
  "sheet.autoFilter.apply(data);",
  "const f = sheet.getRange('C2:C4').conditionalFormats.add(Excel.ConditionalFormatType.colorScale);",
  "f.colorScale.criteria = {",
  " minimum: {type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: '#F8696B'},",
  " midpoint: {type: Excel.ConditionalFormatColorCriterionType.percent, formula: '50', color: '#FFEB84'},",
  " maximum: {type: Excel.ConditionalFormatColorCriterionType.highestValue, color: '#63BE7B'}",
  "};",
  "await context.sync();",
  "sheet.autoFilter.load('enabled');",
  "const count = sheet.getRange('C2:C4').conditionalFormats.getCount();",
  "await context.sync();",
  "return {filterEnabled: sheet.autoFilter.enabled, conditionalFormatCount: count.value};"
].join("\n");

Office.onReady(async (info) => {
  report("Host Office.onReady: " + JSON.stringify(info));
  if (info.host !== Office.HostType.Excel) {
    report("Échec : ouvrir ce manifeste dans Excel.");
    return;
  }
  report("Host ExcelApi 1.9: " + Office.context.requirements.isSetSupported("ExcelApi", "1.9"));
  prepare.disabled = false;
});

prepare.addEventListener("click", async () => {
  prepare.disabled = true;
  try {
    await Excel.run(async (context) => {
      const existing = context.workbook.worksheets.getItemOrNullObject(SHEET);
      existing.load("isNullObject");
      await context.sync();
      if (!existing.isNullObject) {
        throw new Error("La feuille synthétique existe déjà. Utiliser un nouveau classeur jetable.");
      }
      const sheet = context.workbook.worksheets.add(SHEET);
      sheet.getRange("A1:C4").values = [
        ["Item", "City", "Score"],
        ["alpha", "A", 10],
        ["beta", "B", 20],
        ["gamma", "A", 30]
      ];
      await context.sync();
    });
    await makeCanary();
    prepared = true;
    execute.disabled = !childReady;
    report("Feuille synthétique et canari IndexedDB préparés.");
  } catch (error) {
    report("Échec préparation: " + String(error));
  }
});

execute.addEventListener("click", () => {
  if (!prepared || !childReady || pending) return;
  execute.disabled = true;
  const id = crypto.randomUUID();
  pending = id;
  report("Envoi d'un script Office.js FIXE à l'iframe isolée...");
  frame.contentWindow.postMessage({type: "probe-run", id, code}, RUNNER);
  setTimeout(() => {
    if (pending !== id) return;
    pending = null;
    report("ÉCHEC : aucune réponse de l'iframe après 10 secondes. Vérifier sa console WebView2 et les erreurs CSP.");
    clearCanary();
  }, 10000);
});

window.addEventListener("message", async (event) => {
  if (event.origin !== RUNNER || event.source !== frame.contentWindow) return;
  const message = event.data;
  if (!message || typeof message !== "object") return;
  if (message.type === "probe-ready") {
    childReady = true;
    report("Iframe initialisée sur " + message.origin + " ; Office ready=" + message.officeReady
      + " ; host=" + String(message.host ?? "(absent)")
      + " ; Excel disponible=" + String(message.excelAvailable ?? "(inconnu)"));
    execute.disabled = !prepared;
    return;
  }
  if (message.type === "probe-started") {
    report("Iframe : commande reçue, exécution en cours.");
    return;
  }
  if (message.type !== "probe-result" || message.id !== pending) return;
  pending = null;
  report("Retour iframe: " + JSON.stringify(message));
  if (message.ok) {
    try {
      const check = await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(SHEET);
        sheet.autoFilter.load("enabled");
        const count = sheet.getRange("C2:C4").conditionalFormats.getCount();
        await context.sync();
        return { filterEnabled: sheet.autoFilter.enabled, colorScaleRules: count.value };
      });
      report("Vérification INDEPENDANTE depuis l'hôte: " + JSON.stringify(check));
      report(check.filterEnabled && check.colorScaleRules === 1 && message.parentAccessBlocked
        && message.canaryVisible === false ? "TEST FONCTIONNEL + ISOLATION : PASS" : "TEST INCOMPLET OU ÉCHEC");
    } catch (error) {
      report("Échec vérification hôte: " + String(error));
    }
  }
  if (!message.ok) report("TEST OFFICE.JS : ÉCHEC (" + message.error + ")");
  clearCanary();
});
window.addEventListener("pagehide", clearCanary);
