# Setup Instructions for Vercel Deployment

## Step 1: Create a Microsoft Entra ID App for SharePoint Access

1. Go to [Entra ID Admin Center](https://entra.microsoft.com) (or [Azure Portal](https://portal.azure.com))
   - *Note: Entra ID and Azure AD are the same - Microsoft renamed it to Entra ID*
2. Sign in with your Microsoft account
3. Go to **Applications** → **App registrations** → **New registration**
   - (Or in Azure Portal: **Azure Active Directory** → **App registrations** → **New registration**)
4. Fill in:
   - Name: `ID Card PDF Proxy`
   - Account types: `Accounts in this organizational directory only`
5. Click **Register**

## Step 2: Get Your Credentials

On the app registration page:
1. Copy **Application (client) ID** → Save as `CLIENT_ID`
2. Copy **Directory (tenant) ID** → Save as `TENANT_ID`
3. Go to **Certificates & secrets** → **New client secret**
4. Set expiration to 24 months
5. Copy the secret value → Save as `CLIENT_SECRET`

## Step 3: Grant Permissions

1. Go to **API permissions**
2. Click **Add a permission**
3. Search for **Microsoft Graph**
4. Select **Application permissions**
5. Search and add:
   - `Files.Read.All`
   - `Sites.Read.All`
6. Click **Grant admin consent**

## Step 4: Deploy to Vercel

1. Go to [Vercel.com](https://vercel.com)
2. Sign in with GitHub
3. Click **Add New** → **Project**
4. Import your `ID_Card` repository
5. Click **Deploy**
6. Wait for deployment to complete (usually 1-2 minutes)

## Step 5: Add Custom Domain

1. After deployment, go to your project
2. Click **Settings** → **Domains**
3. Click **Add** and enter: `info.ramendu.com`
4. Vercel will show you DNS records to add
5. Go to your domain registrar (where you bought info.ramendu.com)
6. Add the DNS records in your registrar's settings
7. Wait 5-30 minutes for DNS to propagate
8. Your site will now be at **https://info.ramendu.com** ✅

## Step 6: Add Environment Variables in Vercel

1. After deployment, go to **Project Settings** → **Environment Variables**
2. Add three variables:
   - Name: `TENANT_ID` → Value: (your tenant ID from Step 2)
   - Name: `CLIENT_ID` → Value: (your client ID from Step 2)
   - Name: `CLIENT_SECRET` → Value: (your secret from Step 2)
3. Click **Save**

## Step 7: Test It

1. Your site will now be deployed at **https://info.ramendu.com** ✅
2. The PDF viewer should load and display your SharePoint PDF
3. Whenever you update the PDF on SharePoint, it will auto-load the new version!

## How It Works

- Your frontend HTML requests the PDF from `/api/getDownloadUrl`
- The backend API uses Microsoft Graph to authenticate with SharePoint
- It fetches the latest PDF and serves it with CORS headers enabled
- The PDF.js viewer displays it responsively on any device

## Troubleshooting

**"Server not configured" error:**
- Make sure all 3 environment variables are set in Vercel

**"Failed to obtain access token":**
- Check that your CLIENT_ID and TENANT_ID are correct
- Verify permissions are granted in Azure

**PDF still not loading:**
- Check browser console for errors
- Verify the SharePoint PDF link is still valid and accessible

