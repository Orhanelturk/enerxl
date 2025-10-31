/* commands.js — ExecuteFunction commands for EnerXL
   - Requires: ExcelApi 1.13+
   - Manifest: <Action xsi:type="ExecuteFunction"><FunctionName>createSolarTemplate</FunctionName></Action>
*/
(function () {
  // =========================
  // Config (edit these)
  // =========================
  const TEMPLATE_URL = "https://orhanelturk.github.io/enerxl/Book1.xlsx?v=2025.10.31-005"; // keep in sync with manifest
  const TARGET_SHEET_NAME = "System Summary";     // preferred sheet in the template
  const RENAMED_SHEET_NAME = "Solar Template";    // desired name after insert

  // =========================
  // Utilities
  // =========================
  function arrayBufferToBase64(buffer) {
    let binary = "";
    const bytes = new Uint8Array(buffer);
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return btoa(binary);
  }

  async function notify(message) {
    const html = `
      <html><body style="font:14px system-ui; padding:16px; line-height:1.4">
        <div>${(message || "").replace(/</g, "&lt;")}</div>
        <button onclick="Office.context.ui.messageParent('ok')"
                style="margin-top:12px;padding:6px 10px;border:1px solid #ccc;border-radius:6px;background:#f6f6f6">
          Close
        </button>
        <script>Office.onReady(()=>{});</script>
      </body></html>`;
    try {
      await OfficeRuntime.displayWebDialog(`data:text/html,${encodeURIComponent(html)}`, {
        height: 30, width: 40, displayInIframe: false
      });
    } catch (_) { /* dialogs may be blocked; ignore */ }
  }

  function ensureSupportOrThrow() {
    const ok = Office?.context?.requirements?.isSetSupported?.("ExcelApi", "1.13");
    if (!ok) {
      throw new Error("This command requires ExcelApi 1.13+. Please update Office (Win/Mac/Web).");
    }
  }

  async function fetchTemplateBase64(url) {
    const resp = await fetch(url, { cache: "no-store" });
    if (!resp.ok) throw new Error(`Template fetch failed: ${resp.status} ${resp.statusText}`);
    const buf = await resp.arrayBuffer();
    return arrayBufferToBase64(buf);
  }

  // Generate a unique, non-conflicting sheet name: "Name", "Name (2)", ...
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

  // =========================
  // Core insert helpers
  // =========================
  async function insertNamedSheetFromBase64(base64, sheetName, renameTo) {
    await Excel.run(async (context) => {
      context.workbook.insertWorksheetsFromBase64(base64, {
        sheetNamesToInsert: [sheetName],
        positionType: Excel.WorksheetPositionType.end
      });
      await context.sync();

      const ws = context.workbook.worksheets.getItem(sheetName);
      const uniqueName = renameTo ? await getUniqueSheetName(context, renameTo) : sheetName;
      if (renameTo) ws.name = uniqueName;
      const finalName = renameTo ? uniqueName : sheetName;
      context.workbook.worksheets.getItem(finalName).activate();
      await context.sync();
      return finalName;
    });
  }

  async function insertAllSheetsFromBase64(base64, renameTo) {
    return await Excel.run(async (context) => {
      const before = context.workbook.worksheets;
      before.load("items/name");
      await context.sync();
      const countBefore = before.items.length;

      context.workbook.insertWorksheetsFromBase64(base64, {
        positionType: Excel.WorksheetPositionType.end
      });
      await context.sync();

      const after = context.workbook.worksheets;
      after.load("items/name");
      await context.sync();

      // Identify the first newly appended sheet (the one at index countBefore)
      const newBlockFirst = after.items[countBefore];
      let finalName = newBlockFirst.name;

      if (renameTo) {
        const uniqueName = await getUniqueSheetName(context, renameTo);
        newBlockFirst.name = uniqueName;
        finalName = uniqueName;
      }

      after.getItem(finalName).activate();
      await context.sync();
      return finalName;
    });
  }

  // =========================
  // Commands
  // =========================
  async function createSolarTemplate(event) {
    try {
      ensureSupportOrThrow();

      const base64 = await fetchTemplateBase64(TEMPLATE_URL);

      // Try named insert first; if TARGET_SHEET_NAME not present in template, fall back.
      let finalName;
      try {
        finalName = await insertNamedSheetFromBase64(base64, TARGET_SHEET_NAME, RENAMED_SHEET_NAME);
        await notify(`✅ Inserted “${TARGET_SHEET_NAME}” as “${finalName}”.`);
      } catch (namedErr) {
        // Most common: named sheet not found in template
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

  async function pingCommand(event) {
    try { await notify("🔔 EnerXL command is wired correctly."); }
    catch (err) { console.error(err); }
    finally { try { event.completed(); } catch (_) {} }
  }

  // =========================
  // Registration
  // =========================
  Office.onReady(() => {
    Office.actions.associate("createSolarTemplate", createSolarTemplate);
    Office.actions.associate("createSolarTemplateAll", createSolarTemplateAll);
    Office.actions.associate("pingCommand", pingCommand);
  });
})();
