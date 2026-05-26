/* global document, Office */

console.log("TASKPANE JS LOADED");

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";

    document.getElementById("run").onclick = runProcess;
  }
});



/* =========================
   GET FILE CONTEXT
========================= */
async function getFileContext() {
  const token = await OfficeRuntime.auth.getAccessToken();

  const res = await fetch("https://graph.microsoft.com/v1.0/me/drive/recent", {
    headers: { Authorization: `Bearer ${token}` }
  });

  const data = await res.json();
  return data.value;
}

/* =========================
   FIND EXCEL FILES
========================= */
async function getTargetExcelFiles(pptItemId) {
  const token = await OfficeRuntime.auth.getAccessToken();

  const itemRes = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/items/${pptItemId}`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  const item = await itemRes.json();
  const parent = item.parentReference;

  const filesRes = await fetch(
    `https://graph.microsoft.com/v1.0/drives/${parent.driveId}/items/${parent.id}/children`,
    {
      headers: { Authorization: `Bearer ${token}` }
    }
  );

  const files = await filesRes.json();

  let financialFile = null;
  let engineeringFile = null;

  for (const f of files.value) {
    if (!f.name || !f.name.endsWith(".xlsx")) continue;

    if (f.name.toLowerCase().startsWith("financial analysis")) {
      financialFile = f;
    }

    if (f.name.toLowerCase().startsWith("engineering analysis")) {
      engineeringFile = f;
    }
  }

  if (!financialFile) throw new Error("Financial Analysis file not found");
  if (!engineeringFile) throw new Error("Engineering Analysis file not found");

  return { financialFile, engineeringFile };
}

/* =========================
   READ NAMED RANGE
========================= */
async function readNamedRange(fileId, name) {
  const token = await OfficeRuntime.auth.getAccessToken();

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/names('${name}')/range`,
    {
      headers: { Authorization: `Bearer ${token}` }
    }
  );

  if (!res.ok) {
    throw new Error(`Failed to read named range: ${name}`);
  }

  const data = await res.json();
  return data.values[0][0];
}

/* =========================
   GET MULTIPLE VALUES
========================= */
async function getExcelMapping(fileId, rangeNames) {
  const mapping = {};

  for (const name of rangeNames) {
    mapping[name] = await readNamedRange(fileId, name);
  }

  return mapping;
}

/* =========================
   REPLACE PLACEHOLDERS
========================= */
async function replacePlaceholders(mapping: any) {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");

    await context.sync();

    for (const slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load("items/textFrame/textRange/text");

      await context.sync();

      for (const shape of shapes.items) {
        if (!shape.textFrame) continue;

        let text = shape.textFrame.textRange.text;

        for (const key in mapping) {
          const placeholder = `{{${key}}}`;
          text = text.replace(placeholder, mapping[key]);
        }

        shape.textFrame.textRange.text = text;
      }
    }

    await context.sync();
  });
}

/* =========================
   MAIN FLOW
========================= */
async function runProcess() {
  try {
    const files = await getFileContext();

    const ppt = files.find((f) => f.name.endsWith(".pptx"));
    if (!ppt) throw new Error("PowerPoint file not found");

    const { financialFile, engineeringFile } =
      await getTargetExcelFiles(ppt.id);

    const financialMapping = await getExcelMapping(
      financialFile.id,
      ["Revenue", "Cost", "Profit"]
    );

    const engineeringMapping = await getExcelMapping(
      engineeringFile.id,
      ["ProjectLoad", "Efficiency"]
    );

    const finalMapping = {
      ...financialMapping,
      ...engineeringMapping
    };

    await replacePlaceholders(finalMapping);

  } catch (err) {
    console.error(err);
    alert(err.message);
  }
}
