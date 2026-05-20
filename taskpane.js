Office.onReady(() => {
  document.getElementById("copyBtn").addEventListener("click", copyFromOtherWorkbook);
});

function setStatus(msg) {
  document.getElementById("status").textContent = msg;
}

/**
 * 1) WSTAW swój Client ID z Entra App Registration (SPA).
 */
const msalConfig = {
  auth: {
    clientId: "PASTE_YOUR_CLIENT_ID_HERE",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: window.location.origin + "/taskpane.html"
  },
  cache: { cacheLocation: "sessionStorage" }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

/**
 * Minimal scope do odczytu plików.
 * (W razie problemów z SharePointem czasem trzeba Sites.Read.All – zależnie od polityk.)
 */
const graphScopes = ["Files.Read.All"];

async function getGraphToken() {
  const accounts = msalInstance.getAllAccounts();
  const request = { scopes: graphScopes, account: accounts[0] };

  try {
    const silent = await msalInstance.acquireTokenSilent(request);
    return silent.accessToken;
  } catch {
    const interactive = await msalInstance.acquireTokenPopup({ scopes: graphScopes });
    return interactive.accessToken;
  }
}

/**
 * /shares wymaga tzw. encoded sharing url:
 * base64(url) -> base64url (bez '='; '/'->'_' ; '+'->'-') oraz prefix "u!"
 */
function toBase64Url(str) {
  const base64 = btoa(unescape(encodeURIComponent(str)));
  return base64.replace(/=+$/g, "").replace(/\+/g, "-").replace(/\//g, "_");
}

async function graphGet(url, token) {
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!r.ok) {
    const text = await r.text();
    throw new Error(`Graph error ${r.status}: ${text}`);
  }
  return r.json();
}

async function copyFromOtherWorkbook() {
  try {
    const fileUrl = document.getElementById("fileUrl").value.trim();
    const sheetName = document.getElementById("sheetName").value.trim();

    if (!fileUrl || !sheetName) {
      setStatus("Podaj sharing link i nazwę arkusza.");
      return;
    }

    setStatus("1/4 Token do Microsoft Graph...");
    const token = await getGraphToken();

    // 1) Resolve sharing link -> driveItem  (GET /shares/{encoded}/driveItem)
    setStatus("2/4 Rozpoznaję plik (shares -> driveItem)...");
    const encoded = "u!" + toBase64Url(fileUrl);
    const driveItem = await graphGet(
      `https://graph.microsoft.com/v1.0/shares/${encoded}/driveItem`,
      token
    );

    const itemId = driveItem.id;
    const driveId = driveItem.parentReference?.driveId;

    if (!driveId || !itemId) {
      throw new Error("Nie udało się ustalić driveId/itemId z driveItem.");
    }

    // 2) Read used range from a worksheet (GET .../usedRange(valuesOnly=true))
    setStatus("3/4 Pobieram dane z arkusza (usedRange)...");
    const usedRange = await graphGet(
      `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/workbook/worksheets/${encodeURIComponent(sheetName)}/usedRange(valuesOnly=true)`,
      token
    );

    const values = usedRange.values;
    if (!Array.isArray(values) || values.length === 0) {
      throw new Error("Brak danych w usedRange (albo zła nazwa arkusza).");
    }

    // 3) Paste into current workbook using Office.js
    setStatus("4/4 Wklejam do aktualnego workbooka...");
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getItemOrNullObject("Imported");
      sheet.load("name");
      await context.sync();

      if (sheet.isNullObject) {
        sheet = context.workbook.worksheets.add("Imported");
      } else {
        sheet.getUsedRangeOrNullObject().clear();
      }

      const start = sheet.getRange("A1");
      const outRange = start.getResizedRange(values.length - 1, values[0].length - 1);
      outRange.values = values;

      sheet.activate();
      await context.sync();
    });

    setStatus(`Gotowe ✅ Wklejono ${values.length} wierszy do arkusza "Imported".`);
  } catch (e) {
    console.error(e);
    setStatus("Błąd ❌ " + (e?.message || e));
  }
}
``