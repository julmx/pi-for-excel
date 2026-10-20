const PARENT = "https://127.0.0.1:3151";
const CANARY = "__PiIframeSyntheticCanary";
let officeReady = false;
let ran = false;

function respond(payload) {
  window.parent.postMessage(payload, PARENT);
}

async function canReadParentCanary() {
  return new Promise((resolve, reject) => {
    let fresh = false;
    const open = indexedDB.open(CANARY, 1);
    open.onupgradeneeded = () => {
      fresh = true;
      open.result.createObjectStore("values");
    };
    open.onerror = () => reject(open.error);
    open.onsuccess = () => {
      const db = open.result;
      const tx = db.transaction("values", "readonly");
      const get = tx.objectStore("values").get("marker");
      get.onsuccess = () => {
        const visible = get.result === "SYNTHETIC_NOT_A_SECRET";
        db.close();
        if (fresh) indexedDB.deleteDatabase(CANARY);
        resolve(visible);
      };
      get.onerror = () => {
        db.close();
        reject(get.error);
      };
    };
  });
}

Office.onReady((info) => {
  officeReady = info.host === Office.HostType.Excel;
  respond({type: "probe-ready", origin: window.location.origin, officeReady,
    host: info?.host ?? null, platform: info?.platform ?? null,
    excelAvailable: typeof Excel !== "undefined"});
});

window.addEventListener("message", async (event) => {
  if (event.origin !== PARENT || event.source !== window.parent) return;
  const request = event.data;
  if (!request || request.type !== "probe-run" || ran
      || typeof request.id !== "string" || typeof request.code !== "string"
      || request.code.length > 10000) return;
  ran = true;
  respond({type: "probe-started", id: request.id});

  let parentAccessBlocked = false;
  try {
    // Access to a cross-origin parent's DOM must throw SecurityError.
    void window.parent.document.body;
  } catch {
    parentAccessBlocked = true;
  }
  try {
    if (!officeReady) throw new Error("Office.onReady did not identify the Excel host in the iframe");
    const canaryVisible = await canReadParentCanary();
    const moduleSource = "export default async function execute(context) {\n"
      + request.code + "\n}";
    const url = URL.createObjectURL(new Blob([moduleSource], {type: "text/javascript"}));
    let outcome;
    try {
      // Deliberately tests dynamic Office.js code in the CHILD origin only.
      const module = await import(url);
      outcome = await Excel.run((context) => module.default(context));
    } finally {
      URL.revokeObjectURL(url);
    }
    respond({type: "probe-result", id: request.id, ok: true,
      origin: window.location.origin, parentAccessBlocked, canaryVisible, outcome});
  } catch (error) {
    respond({type: "probe-result", id: request.id, ok: false,
      origin: window.location.origin, parentAccessBlocked,
      error: String(error)});
  }
});
