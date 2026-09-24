# IPS Ops Hub

Live operational dashboard for the IPS team.  
Hosted on **GitHub Pages** · refreshed daily at **07:30 IST** via GitHub Actions.

---

## Repository layout

```
ips-ops-hub/
├── .github/
│   └── workflows/
│       └── refresh.yml        # Daily GitHub Actions job
├── public/
│   ├── template.html          # Dashboard HTML (static shell — commit this)
│   ├── index.html             # Built output (auto-generated — commit initially, then CI owns it)
│   └── data.json              # Live data payload (auto-generated — gitignore or commit)
└── scripts/
    ├── package.json
    ├── fetch-data.js          # Pulls SharePoint + Jira data → data.json
    └── build-dashboard.js     # Injects data.json into template.html → index.html
```

---

## One-time setup

### 1. Create the GitHub repo

Under the **Material-Dev** org, create a new **private** repo called `ips-ops-hub`.  
Push this folder as the initial commit.

### 2. Enable GitHub Pages

Repo → **Settings → Pages → Source**: select **GitHub Actions**.

### 3. Add repository secrets

Repo → **Settings → Secrets and variables → Actions → New repository secret**

| Secret name | Where to get it |
|---|---|
| `AZURE_TENANT_ID` | Azure Portal → Azure Active Directory → Overview → Tenant ID |
| `AZURE_CLIENT_ID` | Azure Portal → App registrations → your app → Application (client) ID |
| `AZURE_CLIENT_SECRET` | Azure Portal → App registrations → your app → Certificates & secrets |
| `SHAREPOINT_DRIVE_ID` | Graph Explorer: `GET /v1.0/sites/{siteId}/drives` — copy the `id` of the drive containing the IPS files |
| `JIRA_EMAIL` | `arindam.chowdhury@materialplus.io` |
| `JIRA_API_TOKEN` | https://id.atlassian.com/manage-profile/security/api-tokens |

> `JIRA_BASE_URL` is already hardcoded in the workflow as `materialplus.atlassian.net`.

### 4. Azure app registration permissions

The Azure app needs these **Application** (not delegated) Microsoft Graph permissions:

| Permission | Why |
|---|---|
| `Files.Read.All` | Read SharePoint Excel files |
| `Sites.Read.All` | Access SharePoint site |

Grant admin consent after adding permissions.

### 5. Find the SharePoint Drive ID

Run this once in Graph Explorer (sign in as yourself):

```
GET https://graph.microsoft.com/v1.0/sites/materialworkplace.sharepoint.com:/sites/IPS:/drives
```

Find the drive whose `name` contains the IPS files and copy its `id`.

### 6. Trigger the first run

Repo → **Actions → Refresh IPS Ops Hub → Run workflow**.  
The dashboard will appear at `https://material-dev.github.io/ips-ops-hub/`.

---

## How it works

```
GitHub Actions (07:30 IST daily)
    │
    ├─ fetch-data.js
    │     ├─ Microsoft Graph → Mat SE Weekly Availability.xlsx  → availBlanks[]
    │     ├─ Microsoft Graph → IPS Vacation Calendar.xlsx       → vac{}
    │     └─ Jira REST API  → IPSIPS open issues               → jira[]
    │     └─ writes scripts/data.json
    │
    └─ build-dashboard.js
          ├─ reads  public/template.html
          ├─ replaces AVAIL_BLANKS, VAC, JIRA constants with live data
          ├─ stamps <!--BUILD_TS--> with IST timestamp
          └─ writes public/index.html + public/data.json
               └─ deployed to GitHub Pages
```

---

## Manual refresh

Anyone with repo access can trigger a fresh build from the **Actions** tab → **Run workflow**.

---

## Updating the dashboard design

Edit `public/template.html` (the design shell) and commit.  
The next Actions run will inject live data into the new template.

The three data constants that get replaced at build time:
- `const AVAIL_BLANKS=new Set([...]);`
- `const VAC={...};`
- `const JIRA=[...];`

Keep these as single-statement declarations so the regex replacement works correctly.
