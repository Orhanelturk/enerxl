/* commands.js — ExecuteFunction logic only (no task pane) */
(function () {
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
    // Lightweight UI so you see errors/success even without a task pane.
    try {
      await OfficeRuntime.displayWebDialog(`data:text/html,
        <html><body style="font:14px system-ui;padding:16px">
        <div>${message.replace(/</g,"&lt;")}</div>
        <button onclick="Office.context.ui.messageParent('ok')"
          style="margin-top:12px;padding:6px 10px">Close</button>
        <script>Office.onReady(()=>{});</script>
        </body></html>`, { height: 30, width: 40, displayInIframe: true });
    } catch (_) { /* ignore if blocked */ }
  }

  async function createSolarTemplate(event) {
    let msg = "";
    try {
      // 0) API support guard
      if (!Office.context.requirements.isSetSupported("ExcelApi", "1.13")) {
        msg = "This command needs ExcelApi 1.13+. Please use latest Excel (Win/Mac/Web) or update Office.";
        await notify(msg);
        return;
      }

      // 1) Template URL (keep this in sync with your published GitHub Pages asset)
      const templateUrl = "https://orhanelturk.github.io/enerxl/Book1.xlsx?v=2025.10.31-004"; // publish this exact file

      // 2) Fetch & convert to base64
      const resp = await fetch(templateUrl, { cache: "no-store" });
      if (!resp.ok) throw new Error(`Template fetch failed: ${resp.status} ${resp.statusText}`);
      const buf = await resp.arrayBuffer();
      const base64 = arrayBufferToBase64(buf);

      // 3) Try inserting a specific sheet first
      let insertedByName = true;
      try {
        await Excel.run(async (context) => {
          context.workbook.insertWorksheetsFromBase64(base64, {
            sheetNamesToInsert: ["System Summary"],    // ensure this exists in Book1.xlsx
            positionType: Excel.WorksheetPositionType.end
          });
          await context.sync();

          const ws = context.workbook.worksheets.getItem("System Summary");
          ws.name = "Solar Template";
          ws.activate();
          await context.sync();
        });
      } catch (e) {
        // Fallback: insert ALL sheets, then activate the first
        insertedByName = false;
        await Excel.run(async (context) => {
          context.workbook.insertWorksheetsFromBase64(base64, {
            positionType: Excel.WorksheetPositionType.end
          });
          await context.sync();

          const sheets = context.workbook.worksheets;
          sheets.load("items/name");
          await context.sync();
          if (sheets.items.length > 0) {
            const first = sheets.items[sheets.items.length - 1]; // newly appended block’s first sheet is last range
            first.name = "Solar Template";
            first.activate();
            await context.sync();
          }
        });
      }

      msg = insertedByName
        ? "✅ Inserted sheet “System Summary” as “Solar Template”."
        : "✅ Inserted all sheets from template. Activated and renamed the last appended sheet to “Solar Template”.";
      await notify(msg);

    } catch (err) {
      console.error("Create Template failed:", err);
      await notify("❌ Create Template failed: " + (err && err.message ? err.message : String(err)));
    } finally {
      try { event.completed(); } catch (_) {}
    }
  }

  Office.onReady(() => {
    Office.actions.associate("createSolarTemplate", createSolarTemplate);
  });
})();
