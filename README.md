# RFP Extractor v2 — Outlook Web Add-in
**RADGov / Radiant Systems**

Hourly auto-scans the **SalesRFP** Outlook folder → Claude AI extracts RFP fields → new rows appended directly to your **SharePoint/OneDrive Excel file** with duplicate detection.

---

## What it does

| Step | What happens |
|---|---|
| ⏱ Every hour (or on demand) | Scans all emails in the `SalesRFP` Outlook folder |
| 🤖 Claude AI | Extracts OppName, State, OwnerOrganization, DueDate from each email |
| 🔍 Dedup check | Reads existing rows from your SharePoint Excel; skips already-added entries |
| ✅ Auto-append | Writes only NEW rows directly into your `RFP_List.xlsx` on SharePoint |

---

## File structure

```
RFPExtractor/
├── manifest.xml          ← Outlook add-in manifest
├── taskpane.html         ← Main UI (Dashboard / Log / Settings tabs)
├── commands.html         ← Required stub
├── auth-callback.html    ← Microsoft OAuth redirect handler
├── src/
│   └── taskpane.js       ← All logic: timer, EWS, Claude AI, Graph API
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    └── icon-80.png
```

---

## Setup: Step-by-step

### Step 1 — Register an Azure App

This is required to call the Microsoft Graph API (to write to SharePoint Excel).

1. Go to [portal.azure.com](https://portal.azure.com) → **Azure Active Directory** → **App registrations** → **New registration**
2. Name: `RFP Extractor`
3. Supported account types: **Accounts in this organizational directory only**
4. Redirect URI: `Single-page application (SPA)` → `https://YOUR_HOSTED_URL/auth-callback.html`
5. Click **Register** — copy the **Application (client) ID** and **Directory (tenant) ID**
6. Go to **API permissions** → Add:
   - `Microsoft Graph → Delegated → Files.ReadWrite`
   - `Microsoft Graph → Delegated → Sites.ReadWrite.All`
7. Click **Grant admin consent**

---

### Step 2 — Host the files

Upload all files to an **HTTPS host** (GitHub Pages, Azure Static Web Apps, Vercel, etc.)

Your URL will be something like: `https://yourorg.github.io/rfp-extractor/`

---

### Step 3 — Update manifest.xml

Replace all 5 instances of `https://YOUR_HOSTED_URL` with your actual hosted URL.

---

### Step 4 — Sideload in Outlook

1. Go to [outlook.office.com](https://outlook.office.com), open any email
2. Click `...` → **Get Add-ins** → **My add-ins** → **Add from file** → upload `manifest.xml`

---

### Step 5 — Configure the add-in

Open the add-in in Outlook → click **Settings** tab:

| Setting | What to enter |
|---|---|
| **Anthropic API Key** | Your `sk-ant-...` key from console.anthropic.com |
| **SharePoint Site URL** | `https://yourorg.sharepoint.com/sites/RADGov` |
| **Excel File Path** | `/Shared Documents/RFP_List.xlsx` |
| **Sheet Name** | The exact tab name in your Excel file (e.g. `Sheet1`) |
| **Scan Interval** | Minutes between auto-scans (default: 60) |

Click **Save Settings**, then click **Sign in with Microsoft** and authenticate with your Microsoft 365 account.

> **Note:** You also need to paste your Azure **Client ID** and **Tenant ID** into the browser console once:
> ```js
> localStorage.setItem('rfpClientId', 'YOUR_CLIENT_ID_HERE');
> localStorage.setItem('rfpTenantId', 'YOUR_TENANT_ID_HERE');
> ```

---

### Step 6 — Use it

- **Dashboard tab**: Toggle **Auto-scan** ON to enable hourly runs. The countdown ring shows time until next scan.
- **Run Now**: Triggers an immediate scan regardless of the timer.
- **Activity Log tab**: Full timestamped log of every action.
- Results table shows each row as **NEW** (written to Excel), **DUP** (already exists, skipped), or **ERR** (extraction failed).

---

## How duplicate detection works

Before appending, the add-in reads all existing rows from your SharePoint Excel and builds a set of `OppName + DueDate` keys. Any extracted row whose key already exists is marked **DUP** and skipped. Only genuinely new rows are written.

---

## How the Excel append works

Uses the **Microsoft Graph workbook API**:
1. Reads the used range to find the next empty row
2. PATCHes the range `A{n}:D{n+k}` with the new rows

No formulas or formatting are changed — rows are appended cleanly at the bottom.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| "No API key set" | Go to Settings tab, enter your Anthropic key, Save |
| "Folder SalesRFP not found" | Check exact folder name in Outlook — must be `SalesRFP` |
| "Graph PATCH failed 401" | Re-authenticate via Settings → Sign in with Microsoft |
| "Graph PATCH failed 403" | Check Azure app has `Sites.ReadWrite.All` permission with admin consent |
| "Could not read existing rows" | Verify the Sheet Name matches exactly (case-sensitive) |
| Auth dialog doesn't open | Make sure `auth-callback.html` is hosted at the correct URL |

---

## Security notes

- Your Anthropic API key is stored in **browser localStorage** — only accessible within the same browser/origin.
- The Microsoft Graph token is also stored in localStorage and scoped to `Files.ReadWrite` + `Sites.ReadWrite.All` for your org only.
- For higher security, proxy both API calls through a backend server.
