/* commands.js — EnerXL robust version with in-sheet logging
   Works with manifest ExecuteFunction: createSolarTemplate
   Requires: ExcelApi 1.13+
*/
(function () {
  const TEMPLATE_URL = "https://orhanelturk.github.io/enerxl/Book1.xlsx"; // keep in sync with manifest
  const TARGET_SHEET_NAME = "System Summary";       // preferred sheet inside template
  const RENAMED_SHEET_NAME = "Solar Template";      // desired final name after insert
  const LOG_SHEET = "_EnerXL_Log";

  // ---------- Utils ----------
  function arrayBufferToBase64(buffer) {
    let binary = "", bytes = new Uint8Array(buffer), chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    return btoa(binary);
  }

  async function logToSheet(lines) {
    try {
      await Excel.run(async (context) => {
        const wb = context.workbook;
        const sheets = wb.worksheets;
        let ws;
        try { ws = sheets.getItem(LOG_SHEET); } catch { /* ignored */ }
        if (!ws) ws = sheets.add(LOG_SHEET);
        ws.activate();
        const rng = ws.getRange("A1");
        rng.load("values");
        await context.sync();
        const existing = (rng.values && rng.values[0] && rng.values[0][0]) ? String(rng.values[0][0]) : "";
        const stamp = new Date().toISOString();
        const add = Array.isArray(lines) ? lines.join("\n") : String(lines);
        const combined = (existing ? existing + "\n" : "") + `[${stamp}] ${add}`;
        ws.getUsedRange(true).clear();
        ws.getRange("A1").values = [[combined]];
        await context.sync();
      });
    } catch { /* logging must never throw */ }
  }

  async function notify(message) {
    await logToSheet(message);
    const html = `<html><body style="font:14px system-ui; padding:16px; line-height:1.4">
      <div>${(message || "").replace(/</g, "&lt;")}</div>
      <button onclick="Office.context.ui.messageParent('ok')" style="margin-top:12px;padding:6px 10px;border:1px solid #ccc;border-radius:6px;background:#f6f6f6">Close</button>
      <script>Office.onReady(()=>{});</script></body></html>`;
    try {
      await OfficeRuntime.displayWebDialog(`data:text/html,${encodeURIComponent(html)}`, { height: 30, width: 40, displayInIframe: false });
    } catch (_) { /* dialogs may be blocked; sheet log already has it */ }
  }

  function ensureSupportOrThrow() {
    const ok = Office?.context?.requirements?.isSetSupported?.("ExcelApi", "1.13");
    if (!ok) throw new Error("ExcelApi 1.13+ is required (insertWorksheetsFromBase64). Try Excel on the web or update Desktop.");
  }

  async function fetchTemplateBase64(url) {
    const resp = await fetch(url, { cache: "no-store" });
    if (!resp.ok) throw new Error(`Template fetch failed: ${resp.status} ${resp.statusText} @ ${url}`);
    const buf = await resp.arrayBuffer();
    return arrayBufferToBase64(buf);
  }

  async function getUniqueSheetName(context, desired) {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();
    const existing = new Set(sheets.items.map(s => s.name));
    if (!existing.has(desired)) return desired;
    let i = 2;
    while (existing.has(`${desired} (${i})`)) i++;
    return `${desired} (${i})`;
  }

  async function insertNamedSheetFromBase64(base64, sheetName, renameTo) {
    return await Excel.run(async (context) => {
      const wb = context.workbook;
      wb.insertWorksheetsFromBase64(base64, {
        sheetNamesToInsert: [sheetName],
        positionType: Excel.WorksheetPositionType.end
      });
      await context.sync();

      const ws = wb.worksheets.getItem(sheetName);
      const finalName = renameTo ? await getUniqueSheetName(context, renameTo) : sheetName;
      if (renameTo) ws.name = finalName;
      wb.worksheets.getItem(finalName).activate();
      await context.sync();
      return finalName;
    });
  }

  async function insertAllSheetsFromBase64(base64, renameTo) {
    return await Excel.run(async (context) => {
      const wb = context.workbook;
      const before = wb.worksheets;
      before.load("items/name");
      await context.sync();
      const countBefore = before.items.length;

      wb.insertWorksheetsFromBase64(base64, { positionType: Excel.WorksheetPositionType.end });
      await context.sync();

      const after = wb.worksheets;
      after.load("items/name");
      await context.sync();

      if (after.items.length <= countBefore) throw new Error("No sheets were inserted (unexpected).");

      const newFirst = after.items[countBefore];
      let finalName = newFirst.name;
      if (renameTo) {
        finalName = await getUniqueSheetName(context, renameTo);
        newFirst.name = finalName;
      }
      after.getItem(finalName).activate();
      await context.sync();
      return finalName;
    });
  }

  // ---------- Commands ----------
  async function createSolarTemplate(event) {
    try {
      ensureSupportOrThrow();
      await notify("⏳ Fetching template…");
      const base64 = await fetchTemplateBase64(TEMPLATE_URL);

      let finalName;
      try {
        finalName = await insertNamedSheetFromBase64(base64, TARGET_SHEET_NAME, RENAMED_SHEET_NAME);
        await notify(`✅ Inserted “${TARGET_SHEET_NAME}” as “${finalName}”.`);
      } catch (namedErr) {
        await notify(`ℹ️ Named sheet insert failed (“${TARGET_SHEET_NAME}”). Falling back to all sheets…\nReason: ${namedErr?.message || namedErr}`);
        finalName = await insertAllSheetsFromBase64(base64, RENAMED_SHEET_NAME);
        await notify(`✅ Inserted all sheets. Activated “${finalName}”.`);
      }
    } catch (err) {
      console.error(err);
      await notify("❌ Create Template failed: " + (err?.message || String(err)));
    } finally {
      try { event.completed(); } catch (_) {}
    }
  }

  async function createSolarTemplateAll(event) {
    try {
      ensureSupportOrThrow();
      await notify("⏳ Fetching template (all sheets) …");
      const base64 = await fetchTemplateBase64(TEMPLATE_URL);
      const finalName = await insertAllSheetsFromBase64(base64, RENAMED_SHEET_NAME);
      await notify(`✅ Inserted all sheets. Activated “${finalName}”.`);
    } catch (err) {
      console.error(err);
      await notify("❌ Create Template (All) failed: " + (err?.message || String(err)));
    } finally {
      try { event.completed(); } catch (_) {}
    }
  }

  // Quick sanity: proves the button is firing and ExcelApi is working
  async function addBlankSheet(event) {
    try {
      await Excel.run(async (context) => {
        const name = "EnerXL_Test";
        const ws = context.workbook.worksheets.add(name);
        ws.getRange("A1").values = [["✅ Command fired."]];
        ws.activate();
        await context.sync();
      });
      await notify("✅ Added a blank sheet named “EnerXL_Test”.");
    } catch (err) {
      await notify("❌ addBlankSheet failed: " + (err?.message || String(err)));
    } finally { try { event.completed(); } catch (_) {} }
  }

  // ---------- Registration ----------
  Office.onReady(() => {
    Office.actions.associate("createSolarTemplate", createSolarTemplate);
    Office.actions.associate("createSolarTemplateAll", createSolarTemplateAll);
    Office.actions.associate("addBlankSheet", addBlankSheet);
  });
})();
