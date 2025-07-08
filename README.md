# Sample Web Dashboard

This repository contains a minimal web dashboard that demonstrates how to integrate **Azure Active Directory** authentication and the **Microsoft Graph API** in a static web site. The project is intended as a starting point for deploying an Azure Static Web App that displays and updates items stored in a SharePoint list.

## Features

- Sign in and sign out using [MSAL.js](https://github.com/AzureAD/microsoft-authentication-library-for-js).
- Retrieve items from a SharePoint list using the Microsoft Graph API.
- Optional form to add new items to the list.
- Automated deployment to Azure Static Web Apps via GitHub Actions.

## Project Structure

```
├── index.html             – Main page with sign in/out buttons and dashboard UI
├── index_no-input.html    – Variant of the page without the "Add item" form
├── style.css              – Basic styling
├── auth.js                – MSAL configuration and sign in/out helpers
├── graph.js               – Microsoft Graph calls (read/write)
├── graph_no-input.js      – Read‑only version of the Graph helpers
├── additem.js             – Handles form submission to create a new list item
├── ui.js                  – Updates the page based on the user's sign‑in state
└── .github/workflows/azure-static-web-apps-*.yml – Deployment workflow
```

## Setup Instructions

### 1. Register an Azure AD Application

1. In the [Azure portal](https://portal.azure.com/), open **Azure Active Directory** → **App registrations** → **New registration**.
2. Give your app a name (e.g., "SharePoint Dashboard")
3. Note the **Application (client) ID** and **Directory (tenant) ID** from the Overview page.
4. Under **Authentication**, add your redirect URIs:
   - For local development: `http://localhost:3000` or `http://localhost:8080`
   - For production: `https://your-domain.com` or `https://your-github-username.github.io/your-repo-name`
5. Under **API permissions**, add:
   - `User.Read` (Microsoft Graph)
   - `Sites.ReadWrite.All` (Microsoft Graph)
   - Grant admin consent if required by your organization

### 2. Find Your SharePoint Site and List Information

#### Get SharePoint Site Graph URL:
1. Navigate to https://graph.microsoft.com/v1.0/sites/YOUR_SHAREPOINT_DOMAIN.sharepoint.com
2. Replace `YOUR_SHAREPOINT_DOMAIN` with your actual SharePoint domain
3. Find your site in the response and copy the full site URL from the Graph response
4. The format will be: `https://graph.microsoft.com/v1.0/sites/yourdomain.sharepoint.com,site-id,web-id`

#### Get SharePoint List ID:
1. Go to your SharePoint list
2. Click the gear icon → **List settings**
3. Look at the URL - the `List` parameter contains your list ID (GUID format)
4. Alternatively, use [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer): 
   `GET https://graph.microsoft.com/v1.0/sites/YOUR_SITE_URL/lists`

### 3. Configure the Application

Update the following files with your specific values:

#### `auth.js`:
- Replace `YOUR_CLIENT_ID_HERE` with your Application (client) ID
- Replace `YOUR_TENANT_ID_HERE` with your Directory (tenant) ID  
- Replace `YOUR_REDIRECT_URI_HERE` with your redirect URI

#### `graph.js` and `graph_no-input.js`:
- Replace `YOUR_SHAREPOINT_SITE_GRAPH_URL_HERE` with your SharePoint site's Graph API URL
- Replace `YOUR_LIST_ID_HERE` with your SharePoint list's GUID

### 4. Customize Your SharePoint List

This sample expects a SharePoint list with the following columns:
- **Title** (Single line of text) - Default column
- **Description** (Single line of text)
- **Status** (Single line of text)
- **Cost** (Number)
- **Complete** (Date and time)
- **Location** (Single line of text)
- **AssignmentLookupId** (Person or Group) - for user assignments

You can modify the fields in `additem.js` and `ui.js` to match your list schema.

## Running Locally

1. Install a simple static web server:
   ```bash
   npm install -g serve
   # or
   npx serve
   ```
2. Update the configuration files as described above
3. Serve the site locally:
   ```bash
   serve .
   ```
4. Navigate to the displayed URL (usually `http://localhost:3000`) in your browser

## Deploy to Azure Static Web Apps

1. Create a new **Static Web App** in the Azure portal and connect it to your GitHub repository.
2. The portal generates a deployment token (used as a repository secret called `AZURE_STATIC_WEB_APPS_API_TOKEN_*`).
3. Commit your code to the `main` branch. The provided GitHub Actions workflow builds and deploys the site automatically.

## Deploy to GitHub Pages

1. In your GitHub repository, go to **Settings** → **Pages**
2. Select **Deploy from a branch** and choose `main` branch
3. Your site will be available at `https://your-username.github.io/your-repo-name`
4. Make sure your `redirectUri` in `auth.js` matches this URL

## Hosting on GitHub

The `.github/workflows/azure-static-web-apps-*.yml` file defines the CI/CD pipeline. When you push changes to the repository, GitHub Actions checks out the code and runs the `Azure/static-web-apps-deploy` action to publish the site.

## Customization Tips

- Use `index_no-input.html` and `graph_no-input.js` if you need a read‑only dashboard.
- Modify `style.css` and the HTML markup to fit your branding.
- Review the Graph calls in `graph.js` and `additem.js` to tailor the list schema or permissions.
- Update the column mappings in `additem.js` and `ui.js` to match your SharePoint list structure.

## Troubleshooting

- **CORS errors**: Make sure your redirect URI is properly configured in Azure AD
- **Authentication failures**: Verify your Client ID and Tenant ID are correct
- **Graph API errors**: Check that your site URL and list ID are correct
- **Permission errors**: Ensure the required Graph API permissions are granted and consented

This sample provides a foundation for building your own Azure‑authenticated dashboard hosted entirely from a GitHub repository.