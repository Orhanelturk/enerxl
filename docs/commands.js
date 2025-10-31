/* commands.js — no taskpane, just command logic */
(function () {
  // Utility: arrayBuffer -> base64 (works in Office WebView)
  function arrayBufferToBase64(buffer) {
    let binary = "";
    const bytes = new Uint8Array(buffer);
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return btoa(binary);
  }

  async function createSolarTemplate(event) {
    try {
      // 1) Decide where to fetch the template from
      // Prefer a hard-coded URL (simple), or read from the manifest via Office.context (advanced).
      // Hard-coded, with cache-bust:
      const templateUrl = "https://orhanelturk.github.io/enerxl/Book1.xlsx?v=2025.10.31-004";

      // 2) Fetch template workbook and convert to base64
      const resp = await fetch(templateUrl, { cache: "no-store" });
      if (!resp.ok) throw new Error(`Template fetch failed: ${resp.status} ${resp.statusText}`);
      const buf = await resp.arrayBuffer();
      const base64 = arrayBufferToBase64(buf);

      // 3) Insert the specific sheet from the template
      await Excel.run(async (context) => {
        context.workbook.insertWorksheetsFromBase64(base64, {
          sheetNamesToInsert: ["System Summary"],           // <-- sheet name inside Book1.xlsx
          positionType: Excel.WorksheetPositionType.end
        });
        await context.sync();

        // 4) Rename and activate
        const ws = context.workbook.worksheets.getItem("System Summary");
        ws.name = "Solar Template";
        ws.activate();
        await context.sync();
      });
    } catch (err) {
      // Optional: lightweight error UI via notification messages
      try {
        Office.addin && Office.addin.showAsTaskpane && console.warn("Tip: no task pane open to show UI.");
      } catch (_) {}
      console.error("Create Template failed:", err);
    } finally {
      // Always complete so the ribbon unfreezes
      try { event.completed(); } catch (_) {}
    }
  }

  // Hook up the command when Office is ready
  Office.onReady(() => {
    Office.actions.associate("createSolarTemplate", createSolarTemplate);
  });
})();
