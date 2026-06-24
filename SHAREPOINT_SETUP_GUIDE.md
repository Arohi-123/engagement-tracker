# Setup Guide — Neovation Client Engagement Tracker (SharePoint / Microsoft 365 version)

This version stores your data in SharePoint Lists inside your own Microsoft 365 tenant, and everyone signs in with their actual work Microsoft account — no shared passwords. About 30–40 minutes total, most of it the one-time admin registration step.

You'll do four things, in order:

1. Create four SharePoint Lists (your database) from the Excel files provided
2. Register the dashboard as an app in your Microsoft 365 admin center
3. Fill in five values inside `index.html`
4. Put the file online

---

## 1. Create the SharePoint Lists

You need a SharePoint site for this — an existing team site works fine, or create a new one (in SharePoint admin center: **Create site** → Team site). You'll create four lists on it, one per data type, using the Excel files provided alongside this guide: `Clients_for_SharePoint.xlsx`, `Opportunities_for_SharePoint.xlsx`, `Engagements_for_SharePoint.xlsx`, `Companies_for_SharePoint.xlsx`. Each already contains your existing data from the workbook, formatted as a table ready to import.

Repeat these steps four times, once per file:

1. On your SharePoint site, go to **Site contents** → **New** → **List**.
2. Choose **From Excel**.
3. Click **Upload file**, select the matching `..._for_SharePoint.xlsx` file.
4. SharePoint will detect the table inside and show you a column mapping/preview screen. Leave the column names and types as detected (they're already set up correctly: text, choice, number, and date columns). If anything shows as "Do not import," make sure it's *included* instead.
5. Name the list exactly as below (this isn't strictly required, but keeps your setup matching this guide):
   - `Clients_for_SharePoint.xlsx` → name the list **Clients**
   - `Opportunities_for_SharePoint.xlsx` → name the list **Opportunities**
   - `Engagements_for_SharePoint.xlsx` → name the list **Engagements**
   - `Companies_for_SharePoint.xlsx` → name the list **Companies**
6. Click **Create**. 

After all four are created, open each one briefly and confirm the row count matches: Clients 202, Opportunities 39, Engagements 34, Companies 18.

> If SharePoint creates an extra blank "Title" column alongside an already-populated one (a rare quirk), open **List settings → Columns**, delete the empty duplicate, and make sure the populated column is the one named exactly `Title`.

## 2. Register the app in Azure AD / Microsoft Entra ID 

This is the one step that needs admin rights on your Microsoft 365 tenant.

1. Go to **entra.microsoft.com** (or **portal.azure.com** → Microsoft Entra ID) and sign in as an admin.
2. Go to **App registrations** → **New registration**.
3. Name it something like "Neovation Client Engagement Tracker".
4. Under **Supported account types**, choose **Accounts in this organizational directory only** (single tenant).
5. Under **Redirect URI**, select platform **Single-page application (SPA)** and enter the web address where you'll host `index.html` (e.g. `https://yourcompany.com/tracker/index.html`). If you don't know the exact final URL yet, you can add it now and edit/add more later under **Authentication**.
6. Click **Register**.
7. On the app's **Overview** page, copy the **Application (client) ID** and the **Directory (tenant) ID** — you'll need both shortly.
8. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**. Search for and add:
   - `User.Read`
   - `Sites.ReadWrite.All`
9. Click **Grant admin consent for [Your Organization]** and confirm. This is what lets the dashboard read and write SharePoint data on behalf of whoever signs in, without each person approving it individually.

> **What this permission means:** `Sites.ReadWrite.All` lets the signed-in user's session read/write any SharePoint site they personally already have access to — through this app. It does not give the app standing access independent of who's logged in. If your IT team prefers a tighter scope limited to only this one site, ask them about `Sites.Selected`, which requires one extra Graph API call to grant the app access to just this site — a bit more setup, happy to walk through it if you'd rather go that route.

## 3. Find your Site ID and List IDs

1. Go to **developer.microsoft.com/graph/graph-explorer** and sign in with your work account.
2. Run this query (replace with your actual SharePoint hostname and site path — e.g. if your site URL is `https://contoso.sharepoint.com/sites/Neovation-BD`, use `contoso.sharepoint.com` and `Neovation-BD`):
   ```
   GET https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/Neovation-BD
   ```
   Copy the `"id"` value from the response — that's your **Site ID**.
3. Then run:
   ```
   GET https://graph.microsoft.com/v1.0/sites/{the-site-id-you-just-copied}/lists
   ```
   In the response, find the four lists by their `"displayName"` (Clients, Opportunities, Engagements, Companies) and copy each one's `"id"` value — these are your four **List IDs**.

## 4. Fill in the configuration

Open `index.html` in a text editor. Near the top of the `<script>` section you'll find:

```js
const MSAL_CLIENT_ID = 'YOUR_APP_CLIENT_ID';
const MSAL_TENANT_ID = 'YOUR_TENANT_ID';
const SHAREPOINT_SITE_ID = 'YOUR_SITE_ID';
const LIST_IDS = {
  clients:       'YOUR_CLIENTS_LIST_ID',
  opportunities: 'YOUR_OPPORTUNITIES_LIST_ID',
  engagements:   'YOUR_ENGAGEMENTS_LIST_ID',
  companies:     'YOUR_COMPANIES_LIST_ID'
};
const ADMIN_EMAILS = [
  'admin@yourcompany.com'
];
```

Replace each placeholder:
- `MSAL_CLIENT_ID` → the Application (client) ID from step 2.7
- `MSAL_TENANT_ID` → the Directory (tenant) ID from step 2.7
- `SHAREPOINT_SITE_ID` → from step 3.2
- The four `LIST_IDS` → from step 3.3
- `ADMIN_EMAILS` → the work email address(es) of whoever should be able to add records. Add as many as you like, one per line, e.g. `'priya@yourcompany.com', 'imran@yourcompany.com'`. Everyone else who signs in gets read-only dashboards — no separate "viewer password" needed anymore, since real sign-in handles that.

Save the file.

## 5. Put the file online

Upload `index.html` to wherever your domain is hosted — the exact same address you registered as the redirect URI in step 2.5. (If you registered a placeholder URL earlier, go back to **Authentication** in your app registration and update/add the real one now — it has to match exactly, including `https://` and the trailing path.)

---

## Using the tool day to day

- **Sign in** with the "Sign in with Microsoft" button — your normal work account, no separate password to remember.
- Anyone in your organization can sign in and view the dashboards. Only the email addresses listed in `ADMIN_EMAILS` see the **+ Add** buttons to create new records.
- Both Overview and each tab pull live from your four SharePoint Lists — anything anyone adds shows up for everyone else immediately.
- The **Companies** tab is calculated automatically by matching company names across the other three lists.

## If something looks wrong

- **Yellow "Setup needed" banner**: one of the five config values is still a placeholder — recheck step 4.
- **Sign-in works but nothing loads, with an error about SharePoint**: double check the Site ID and List IDs are exactly right, and that admin consent was granted in step 2.9.
- **"Could not save" when adding a record**: usually a permissions or List ID mismatch — confirm the signed-in account has edit access to the SharePoint site itself.
- **A colleague can't add records even though you expected them to**: check their exact sign-in email is listed in `ADMIN_EMAILS`, spelled exactly as their Microsoft 365 sign-in address.
- **Editing or deleting existing rows**: this first version supports adding new records through the form. To correct an existing row, open the list directly in SharePoint (it behaves like an editable spreadsheet grid). Let me know if you'd like in-app edit/delete added next.

## Files in this delivery

| File | Purpose |
|---|---|
| `index.html` | The dashboard itself. Upload this to your hosting once configured. |
| `Clients_for_SharePoint.xlsx` | Import this to create your Clients list (202 rows, pre-filled). |
| `Opportunities_for_SharePoint.xlsx` | Import this to create your Opportunities list (39 rows, pre-filled). |
| `Engagements_for_SharePoint.xlsx` | Import this to create your Engagements list (34 rows, pre-filled). |
| `Companies_for_SharePoint.xlsx` | Import this to create your Companies list (18 rows, pre-filled). |
| `SHAREPOINT_SETUP_GUIDE.md` | This file. |
