# CheckFire Marketing Hub — Azure Deployment Guide

This guide takes you from zero to a live internal website in about 30 minutes.
You need: an Azure subscription, access to Azure Active Directory (Entra ID), and your SharePoint site URL.

---

## Overview of what you're building

```
CheckFire staff browser
        │  signs in with Microsoft account
        ▼
Azure Static Web App  ──►  Microsoft Graph API  ──►  SharePoint
  (hosts the hub)           (reads your data)         (your lists & files)
```

Staff visit a URL, sign in once with their existing CheckFire Microsoft account, and see live data from SharePoint. No passwords to manage, no separate login — it just uses what they already have.

---

## Step 1 — Create an Azure Static Web App

1. Go to [portal.azure.com](https://portal.azure.com) and sign in
2. Click **Create a resource** → search for **Static Web App** → click **Create**
3. Fill in:
   - **Subscription**: your CheckFire subscription
   - **Resource group**: create new → `rg-marketing-hub`
   - **Name**: `checkfire-marketing-hub`
   - **Region**: UK South (or West Europe)
   - **Hosting plan**: Free
   - **Deployment source**: **Other** (we'll upload manually)
4. Click **Review + create** → **Create**
5. Once created, open the resource and note the **URL** (e.g. `https://checkfire-marketing-hub.azurestaticapps.net`)

---

## Step 2 — Register the app in Azure Active Directory

This is the one-time IT step. It tells Microsoft "this website is allowed to read SharePoint data on behalf of signed-in users."

1. In the Azure Portal, go to **Azure Active Directory** (search for it in the top bar)
2. Click **App registrations** → **New registration**
3. Fill in:
   - **Name**: `CheckFire Marketing Hub`
   - **Supported account types**: **Accounts in this organizational directory only (CheckFire only)**
   - **Redirect URI**: select **Single-page application (SPA)** and enter your Static Web App URL + `/` (e.g. `https://checkfire-marketing-hub.azurestaticapps.net/`)
4. Click **Register**
5. On the app overview page, copy:
   - **Application (client) ID** — you'll need this shortly
   - **Directory (tenant) ID** — you'll need this too

### Add API permissions

1. Click **API permissions** → **Add a permission** → **Microsoft Graph** → **Delegated permissions**
2. Add these permissions:
   - `User.Read`
   - `Sites.Read.All`
   - `Files.Read.All`
3. Click **Grant admin consent for CheckFire** → **Yes**

That's it for IT. The app registration is done.

---

## Step 3 — Configure the hub

Open `js/config.js` in this folder and fill in the values:

```js
const HUB_CONFIG = {
  tenantId:   'paste-your-tenant-id-here',
  clientId:   'paste-your-client-id-here',
  sharepointSite: 'https://checkfire.sharepoint.com/sites/Marketing',
  lists: {
    launches:  'Product Launches',   // must match exact SharePoint list name
    campaigns: 'Campaigns',
    events:    'Events',
  },
  documentsLibrary: 'Marketing Assets',
};
```

**SharePoint list columns required:**

| Hub section | List name | Required columns |
|---|---|---|
| Launches | Product Launches | Title, LaunchDate, SKU, Status, Description, RRP, LinkURL |
| Campaigns | Campaigns | Title, CampaignType, Status, StartDate, EndDate, Budget, Channels, Region, LinkURL |
| Events | Events | Title, EventDate, Location, EventType, Status, LinkURL |
| Resources | *(document library)* | Any document library — shows files automatically |

Column names are case-sensitive. The simplest approach is to create a SharePoint list with these exact column names.

---

## Step 4 — Deploy the files

### Option A: Drag and drop (quickest)

1. In the Azure Portal, open your Static Web App resource
2. Under **Deployment** → click **Manage deployment token** and copy it
3. Install the Azure Static Web Apps CLI:
   ```
   npm install -g @azure/static-web-apps-cli
   ```
4. From this folder, run:
   ```
   swa deploy . --deployment-token YOUR_TOKEN
   ```

### Option B: GitHub Actions (recommended for ongoing updates)

1. Push this `hub-app` folder to a GitHub repository
2. In the Azure Portal, open your Static Web App → **Deployment** → **GitHub**
3. Connect your repo — Azure will create a GitHub Action automatically
4. Every push to `main` automatically redeploys the hub

---

## Step 5 — Test it

1. Open your Static Web App URL in a browser
2. You should see a Microsoft sign-in screen
3. Sign in with your CheckFire account
4. The hub loads and pulls live data from SharePoint

If data doesn't appear, open the browser developer tools (F12) → Console tab — error messages will tell you exactly what's missing (wrong list name, missing column, permissions issue).

---

## Restricting access

By default, any Microsoft account can sign in. To restrict to CheckFire staff only:

1. In Azure AD, go to your app registration → **Enterprise applications** → find `CheckFire Marketing Hub`
2. Click **Properties** → set **Assignment required** to **Yes**
3. Click **Users and groups** → add your CheckFire staff or a security group

Only people in that group will be able to sign in.

---

## Sharing with the marketing team

Once deployed, the hub is just a URL. Share it like any internal link:
- Pin it in the Teams Marketing channel
- Add it as a SharePoint quick link
- Bookmark it in the browser

No app to install, no account to create — they sign in with the Microsoft account they already use.

---

## Updating content

Content in the hub comes from two places:

1. **SharePoint lists & libraries** — update these in SharePoint as normal; the hub reflects changes immediately on next page load
2. **Hub layout/design** — edit the HTML/CSS files and redeploy (or push to GitHub if using Option B above)

The goal is that marketing can manage almost everything through SharePoint, and only a developer needs to touch the code if the layout needs changing.

---

## Troubleshooting

| Problem | Likely cause | Fix |
|---|---|---|
| Sign-in loop / blank page | Redirect URI mismatch | Check the URI in App Registration matches your site URL exactly (including trailing `/`) |
| "List not found" error | List name doesn't match | Check `config.js` list names match exactly — case-sensitive |
| "Insufficient privileges" | Permissions not granted | Go to App Registration → API permissions → Grant admin consent |
| No files in Resources | Wrong library name | Check `documentsLibrary` in config matches the SharePoint library name |
| Works for you, not others | Assignment required | Check Enterprise Application → Users and groups |

