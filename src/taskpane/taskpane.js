// ============================================================
// CONFIGURATION — fill these in after Azure App Registration
// ============================================================
const MSAL_CONFIG = {
  auth: {
    clientId: "7e82e477-f359-476c-b75e-17fc4c2a979b",        // From App Registration
    authority: "https://login.microsoftonline.com/9eaa5bb5-299e-4d3b-a108-c84d5c8349e7",
    redirectUri: "https://officialfred.github.io/pa-slidedeck-addin/taskpane.html"
  },
  cache: { cacheLocation: "sessionStorage" }
};

import * as msal from "@azure/msal-browser";

// Graph API scopes needed
const SCOPES = ["Files.Read", "Sites.Read.All"];

// ============================================================
// CELL MAPPING — define which cells map to which placeholders
// Each key is a placeholder string in your PowerPoint (e.g. {{REVENUE}})
// Each value is the Excel cell address to read from
// ============================================================
const CELL_MAPPING = {
  "{{PropertyName}}​":      null,
  "{{breakUpFee}}":  "OBF scenario!F22",
  "{{ecm1Name}}":   "Measures!B4",
  "{{ecm2Name}}":   "Measures!B5",
  "{{ecm3Name}}":   "Measures!B6",
  "{{ecm4Name}}":   "Measures!B7",
  "{{ecm5Name}}":   "Measures!B8",
  "{{ecm6Name}}":   "Measures!B9",
  "{{ecm7Name}}":   "Measures!B10",
  "{{ecm8Name}}":   "Measures!B11",
  "{{ecm9Name}}":   "Measures!B12",
  "{{ecm1Description}}":     "Measure descriptions!C2",
  "{{ecm2Description}}":     "Measure descriptions!C3",
  "{{ecm3Description}}":     "Measure descriptions!C4",
  "{{ecm4Description}}":     "Measure descriptions!C5",
  "{{ecm5Description}}":     "Measure descriptions!C6",
  "{{ecm6Description}}":     "Measure descriptions!C7",
  "{{ecm7Description}}":     "Measure descriptions!C8",
  "{{ecm8Description}}":     "Measure descriptions!C9",
  "{{ecm9Description}}":     "Measure descriptions!C10",

  // Add more mappings here as needed
};

// ============================================================
// GLOBALS
// ============================================================
let msalInstance;
let accessToken = null;
let foundExcelFiles = [];

// ============================================================
// INIT
// ============================================================
Office.onReady(() => {
  console.log("hello world)

  msalInstance = new msal.PublicClientApplication(MSAL_CONFIG);

  // Handle redirect response (for redirect flow)
  msalInstance.handleRedirectPromise().then(handleAuthResponse).catch(e => log(e, "err"));

  document.getElementById("sign-in-btn").addEventListener("click", signIn);
  document.getElementById("sign-out-btn").addEventListener("click", signOut);
  document.getElementById("find-sheets-btn").addEventListener("click", findExcelSheets);
  document.getElementById("run-btn").addEventListener("click", runReplacement);

  // Check if already signed in
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    showMainSection(accounts[0]);
    acquireToken();
  }
});

// ============================================================
// AUTH
// ============================================================
async function signIn() {
  try {
    await msalInstance.loginRedirect({ scopes: SCOPES });
    // Page will redirect away and come back — handleRedirectPromise() catches the return
  } catch (e) {
    log("Sign-in failed: " + e.message, "err");
  }
}

function handleAuthResponse(resp) {
  if (resp && resp.account) {
    showMainSection(resp.account);
    acquireToken();
  }
}

function showMainSection(account) {
  document.getElementById("sign-in-section").style.display = "none";
  document.getElementById("main-section").style.display = "block";
  document.getElementById("user-info").textContent = "Signed in as: " + account.username;
}

async function acquireToken() {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) return;
  try {
    const resp = await msalInstance.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
    accessToken = resp.accessToken;
    log("Token acquired ✓", "ok");
  } catch (e) {
    // Silent failed, fall back to redirect instead of popup
    await msalInstance.acquireTokenRedirect({ scopes: SCOPES });
  }
}

function signOut() {
  msalInstance.logoutRedirect({
    postLogoutRedirectUri: window.location.href
  }).then(() => {
    document.getElementById("sign-in-section").style.display = "block";
    document.getElementById("main-section").style.display = "none";
    document.getElementById("sheets-list").style.display = "none";
    accessToken = null;
    foundExcelFiles = [];
  });
}

// ============================================================
// STEP 1: FIND THE PRESENTATION'S FOLDER, THEN FIND EXCEL FILES
// ============================================================
async function findExcelSheets() {
  if (!accessToken) { log("Not signed in.", "err"); return; }
  log("Looking up current presentation...", "inf");

  try {
    // Get the current file's path from Office context
    const fileUrl = Office.context.document.url;
    log("Presentation URL: " + fileUrl, "inf");

    // Resolve the Drive item for this file via Graph
    const encodedUrl = encodeURIComponent(fileUrl);
    const itemResp = await graphFetch(`/v1.0/shares/u!${btoa(fileUrl).replace(/=/g, "").replace(/\+/g, "-").replace(/\//g, "_")}/driveItem`);

    // Fallback: search using the file name from URL
    let parentId;
    if (itemResp && itemResp.parentReference) {
      parentId = itemResp.parentReference.id;
      log("Found parent folder ID: " + parentId, "ok");
    } else {
      // Alternative: use /me/drive/root and search by filename
      log("Could not resolve via share link, searching by filename...", "inf");
      const fileName = decodeURIComponent(fileUrl.split("/").pop().split("?")[0]);
      const searchResp = await graphFetch(`/v1.0/me/drive/root/search(q='${encodeURIComponent(fileName)}')?$select=id,name,parentReference`);
      if (!searchResp || !searchResp.value || searchResp.value.length === 0) {
        log("Could not locate presentation in OneDrive.", "err");
        return;
      }
      parentId = searchResp.value[0].parentReference.id;
    }

    // List Excel files in the same folder
    const childrenResp = await graphFetch(`/v1.0/me/drive/items/${parentId}/children?$filter=file ne null&$select=id,name,file`);
    if (!childrenResp || !childrenResp.value) {
      log("Failed to list folder contents.", "err");
      return;
    }

    foundExcelFiles = childrenResp.value.filter(f =>
      f.name.endsWith(".xlsx") || f.name.endsWith(".xls")
    );

    if (foundExcelFiles.length === 0) {
      log("No Excel files found in the same folder.", "err");
      return;
    }

    // Display found files
    const ul = document.getElementById("sheets-ul");
    ul.innerHTML = "";
    foundExcelFiles.forEach(f => {
      const li = document.createElement("li");
      li.textContent = f.name;
      ul.appendChild(li);
    });
    document.getElementById("sheets-list").style.display = "block";
    log(`Found ${foundExcelFiles.length} Excel file(s) ✓`, "ok");

  } catch (e) {
    log("Error finding sheets: " + e.message, "err");
    console.error(e);
  }
}

// ============================================================
// STEP 2: READ CELLS FROM EXCEL FILES
// The add-in reads from each Excel file and builds a replacement map.
// If multiple Excel files are found, it merges all their data.
// Duplicate placeholders: last file wins.
// ============================================================
async function readExcelData(fileId) {
  const data = {};
  for (const [placeholder, cellAddress] of Object.entries(CELL_MAPPING)) {
    if (cellAddress === null) continue;   // ← skip filename-derived placeholders
    try {
      const resp = await graphFetch(
        `/v1.0/me/drive/items/${fileId}/workbook/worksheets/${encodeSheetName(cellAddress)}/range(address='${encodeURIComponent(cellAddress)}')`
      );
      if (resp && resp.values && resp.values[0] && resp.values[0][0] !== undefined) {
        data[placeholder] = String(resp.values[0][0]);
        log(`  ${placeholder} = "${data[placeholder]}" (from ${cellAddress})`, "ok");
      }
    } catch (e) {
      log(`  Could not read ${cellAddress}: ${e.message}`, "err");
    }
  }
  return data;
}

function encodeSheetName(cellAddress) {
  // Extract sheet name from "Sheet1!B2" → "Sheet1"
  return encodeURIComponent(cellAddress.split("!")[0]);
}


// ============================================================
// STEP 3: REPLACE PLACEHOLDERS IN POWERPOINT
// ============================================================
async function runReplacement() {
  if (foundExcelFiles.length === 0) { log("No Excel files loaded. Run search first.", "err"); return; }
  log("Reading Excel data...", "inf");

  // Build combined replacement map from all found Excel files
  let replacements = {};
  for (const file of foundExcelFiles) {
    log(`Reading from: ${file.name}`, "inf");
    const fileData = await readExcelData(file.id);
    replacements = { ...replacements, ...fileData };
  }

  // ── NEW: inject filename-derived values ──────────────────
  const projectName = extractProjectName(foundExcelFiles);
  if (projectName) {
    replacements["{{PROJECT_NAME}}"] = projectName;
    log(`  {{PROJECT_NAME}} = "${projectName}" (from filename)`, "ok");
  } else {
    log(`  Warning: no file matching "Engineering Analysis - ..." found`, "err");
  }
  // ─────────────────────────────────────────────────────────

  log("Replacing placeholders in presentation...", "inf");

  try {
    await PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      for (const slide of slides.items) {
        const shapes = slide.shapes;
        shapes.load("items");
        await context.sync();

        for (const shape of shapes.items) {
          // Only process shapes with text
          if (!shape.textFrame) continue;

          try {
            const textFrame = shape.textFrame;
            textFrame.load("textRange");
            await context.sync();

            const textRange = textFrame.textRange;
            textRange.load("text");
            await context.sync();

            let currentText = textRange.text;
            let modified = false;

            for (const [placeholder, value] of Object.entries(replacements)) {
              if (currentText.includes(placeholder)) {
                // Use find/replace to preserve formatting where possible
                const found = textRange.find(placeholder, { matchCase: true });
                found.load("text");
                await context.sync();
                if (found.text) {
                  found.text = value;
                  modified = true;
                  log(`  Replaced "${placeholder}" → "${value}" on slide`, "ok");
                }
              }
            }

            if (modified) await context.sync();
          } catch (shapeErr) {
            // Some shapes don't have editable text — skip silently
          }
        }
      }
    });

    log("✅ All replacements complete!", "ok");
  } catch (e) {
    log("PowerPoint error: " + e.message, "err");
    console.error(e);
  }
}

// ============================================================
// GRAPH API HELPER
// ============================================================
async function graphFetch(path) {
  if (!accessToken) throw new Error("No access token");
  const resp = await fetch("https://graph.microsoft.com" + path, {
    headers: { Authorization: "Bearer " + accessToken }
  });
  if (!resp.ok) {
    const errText = await resp.text();
    throw new Error(`Graph ${path} → ${resp.status}: ${errText}`);
  }
  return resp.json();
}

// ============================================================
// LOGGING
// ============================================================
function log(msg, type = "inf") {
  const div = document.getElementById("log");
  const line = document.createElement("div");
  line.className = "log-" + type;
  line.textContent = new Date().toLocaleTimeString() + " " + msg;
  div.appendChild(line);
  div.scrollTop = div.scrollHeight;
}

// ============================================================
// FILENAME-DERIVED VALUES
// Extracts project name from "Engineering Analysis - PROJECT NAME.xlsx"
// ============================================================
function extractProjectName(files) {
  const prefix = "Engineering Analysis - ";
  for (const file of files) {
    if (file.name.startsWith(prefix)) {
      // Strip the prefix and the .xlsx extension
      return file.name.slice(prefix.length).replace(/\.xlsx?$/i, "").trim();
    }
  }
  return null;
}
