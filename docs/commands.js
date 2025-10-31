/* commands.js — ExecuteFunction commands for EnerXL
   - Requires: ExcelApi 1.13+
   - Wire these in manifest with <Action xsi:type="ExecuteFunction"><FunctionName>createSolarTemplate</FunctionName></Action>
*/

(function () {
  // =========================
  // Config (edit these)
  // =========================
  const TEMPLATE_URL = "https://orhanelturk.github.io/enerxl/Book1.xlsx?v=2025.10.31-005"; // <-- update when you republish
  const TARGET_SHEET_NAME = "System Summary";            // <-- sheet to pull from the template
  const RENAMED_SHEET_NAME = "Solar Template";           // <-- how it should appear after insert

  // =========================
  // Utilities
  // =========================
  function arrayBufferToBase64(buffer) {
    let binary = "";
    const bytes = new Uint8Array(buffer);
    const chunk = 0x8000; // chunking avoids call stack issues
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return btoa(binary);
  }

  async function notify(message) {
    // Small dialog so ExecuteFunction gives user-visible feedback.
    const html = `
      <html>
        <body style="font:14px system-ui; padding:16px; line-height:1.4">
          <div>${(message || "").replace(/</g, "&lt;")}</div>
          <button onclick="Office.context.ui.messageParent('ok')"
                  style="margin-top:12px;padding:6px 10px;border:1px solid #ccc;border-radius:6px;background:#f6f6f6">
            Close
          </button>
          <script>Office.onReady(()=>{});</script>
        </body>
      </html>`;
    try {
      await OfficeRuntime.displayWebDialog(`data:text/html,${encodeURIComponent(html)}`, {
        height: 30, width: 40, displayInIframe: true
      });
    } catch (_) {
      // If dialogs are blocked, ignore silently.
    }
  }

  function ensureSupportOrThrow() {
    const ok = Office?.context?.requirements?.isSetSupported?.("ExcelApi", "1.13");
    if (!ok) {
      throw new Error("This command requires ExcelApi 1.13+. Please update Office (Win/Mac/Web).");
    }
  }

  async function fetchTemplateBase64(url) {
    const resp = await fetch(url, { cache: "no-store" });
    if (!resp.ok) {
      throw new Error(`Template fetch failed: ${resp.status} ${resp.statusText}`);
    }
    const buf = await resp.arrayBuffer();
    return arrayBufferToBase64(buf);
  }

  // =========================
  // Core insert helpers
  // =========================
  async function insertNamedSheetFromBase64(base64, sheetName, renamedTo) {
    await Excel.run(async (context) => {
      context.workbook.insertWorksheetsFromBase64(base64, {
        sheetNamesToInsert: [sheetName],
        positionType: Excel.WorksheetPositionType.end
      });
      await context.sync();

      const ws = context.workbook.worksheets.getItem(sheetName);
      if (renamedTo && renamedTo !== sheetName) ws.name = renamedTo;
      (renamedTo ? context.workbook.worksheets.getItem(renamedTo) : ws).activate();
      await context.sync();
    });
  }

  async function insertAllSheetsFromBase64(base64, renamedTo) {
    await Excel.run(async (context) => {
      // Insert all sheets at the end
      context.workbook.insertWorksheetsFromBase64(base64, {
        positionType: Excel.WorksheetPositionType.end
      });
      await context.sync();

      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      // Newly appended block appears after existing ones; activate the first of the newly appended block.
      // A simple way: activate the last sheet and rename it.
      if (sheets.items.length > 0) {
        const last = sheets.items[sheets.items.length - 1];
        if (renamedTo) last.name = renamedTo;
        last.activate();
        await context.sync();
      }
    });
  }

  // =========================
  // Commands
  // =========================
  async function createSolarTemplate(event) {
    let message = "";
    try {
      ensureSupportOrThrow();

      const base64 = await fetchTemplateBase64(TEMPLATE_URL);

      // Try named-sheet insert first
      let usedNamedInsert = true;
      try {
        await insertNamedSheetFromBase64(base64, TARGET_SHEET_NAME, RENAMED_SHEET_NAME);
      } catch (namedErr) {
        // Fallback: insert all sheets
        usedNamedInsert = false;
        await insertAllSheetsFromBase64(base64, RENAMED_SHEET_NAME);
      }

      message = usedNamedInsert
        ? `✅ Inserted “${TARGET_SHEET_NAME}” as “${RENAMED_SHEET_NAME}”.`
        : `✅ Inserted all sheets. Activated & renamed the last appended sheet to “${RENAMED_SHEET_NAME}”.`;
      await notify(message);

    } catch (err) {
      await notify("❌ Create Template failed: " + (err?.message || String(err)));
      console.error(err);
    } finally {
      try { event.completed(); } catch (_) {}
    }
  }

  // Force “insert all” version, in case you want a separate ribbon button for this behavior.
  async function createSolarTemplateAll(event) {
    try {
      ensureSupportOrThrow();
      const base64 = await fetchTemplateBase64(TEMPLATE_URL);
      await insertAllSheetsFromBase64(base64, RENAMED_SHEET_NAME);
      await notify(`✅ Inserted all sheets and activated “${RENAMED_SHEET_NAME}”.`);
    } catch (err) {
      await notify("❌ Create Template (All) failed: " + (err?.message || String(err)));
      console.error(err);
    } finally {
      try { event.completed(); } catch (_) {}
    }
  }

  // Tiny wiring test—useful to confirm the ribbon button is calling an ExecuteFunction at all.
  async function pingCommand(event) {
    try {
      await notify("🔔 EnerXL command is wired correctly.");
    } catch (err) {
      console.error(err);
    } finally {
      try { event.completed(); } catch (_) {}
    }
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
