# Assessment Report Analyzer - Complete Setup Guide

**For users with no programming experience | Microsoft 365 Edition**

This guide assumes you have never set up a software application before. Follow each step exactly as written. The process takes about 45-60 minutes the first time.

---

## Table of Contents

1. [What You're Setting Up](#part-1-what-youre-setting-up)
2. [Getting Your Anthropic API Key](#part-2-getting-your-anthropic-api-key)
3. [Setting Up Excel Online with Microsoft 365](#part-3-setting-up-excel-online-with-microsoft-365)
4. [Creating a GitHub Account](#part-4-creating-a-github-account)
5. [Uploading the Application](#part-5-uploading-the-application)
6. [Deploying on Streamlit Cloud](#part-6-deploying-on-streamlit-cloud)
7. [Testing Your Application](#part-7-testing-your-application)
8. [Using the Application](#part-8-using-the-application)
9. [Troubleshooting](#part-9-troubleshooting)
10. [Managing Costs](#part-10-managing-costs)

---

## Part 1: What You're Setting Up

### The Application
You're setting up a web application that:
- Accepts assessment reports (PDF or Word)
- Analyzes them using AI (Claude by Anthropic)
- Provides constructive feedback
- Extracts and stores metadata in Excel Online (SharePoint)
- Tracks assessment data across multiple years

### The Components
1. **Anthropic API** - The AI that analyzes reports (you pay per report, ~$0.05-0.10 each)
2. **Excel Online (Microsoft 365)** - Where metadata is stored (uses your existing UTA Microsoft license)
3. **Streamlit Cloud** - Where the application runs (free)
4. **GitHub** - Where the application code is stored (free)

### Cost Summary
- One-time setup: Free
- Monthly hosting: Free
- Excel Online storage: Free (UTA Microsoft license)
- Per report analyzed: ~$0.05-0.10
- Estimated monthly cost for typical use (100 reports): ~$5-10

---

## Part 2: Getting Your Anthropic API Key

The API key is like a password that lets your application use Claude AI.

### Step 2.1: Create an Anthropic Account

1. Open your web browser
2. Go to: **https://console.anthropic.com/**
3. Click **"Sign up"**
4. Enter your email address (use your work email)
5. Create a password
6. Click the verification link sent to your email
7. Complete any additional verification steps

### Step 2.2: Add Billing Credits

You need to add money to your account before you can use the API.

1. After logging in, look at the left sidebar
2. Click **"Billing"** (or "Plans & Billing")
3. Click **"Add payment method"**
4. Enter your credit card information
5. Click **"Add credits"**
6. Add **$20** to start (this will process approximately 200-400 reports)
7. Click **"Confirm"**

### Step 2.3: Create Your API Key

1. In the left sidebar, click **"API Keys"**
2. Click the button **"Create Key"**
3. In the "Name" field, type: `Assessment Analyzer`
4. Click **"Create Key"**
5. You will see a long string starting with `sk-ant-api03-...`
6. **IMPORTANT**: Click the copy button next to the key
7. Open a text file (Notepad)
8. Paste the key and save the file somewhere safe
9. Name the file something like `api_keys_DO_NOT_SHARE.txt`

**WARNING**: 
- This key is like a password - never share it publicly
- If someone gets your key, they can use your credits
- You can only see the full key once - if you lose it, create a new one

---

## Part 3: Setting Up Excel Online with Microsoft 365

Since UTA is a Microsoft campus, you'll use Excel Online in SharePoint for data storage.

### Step 3.1: Create the Excel File in SharePoint

1. Go to your SharePoint site or OneDrive for Business
2. Navigate to a folder where you want to store the tracking file (e.g., "Institutional Effectiveness" folder)
3. Click **"New"** → **"Excel workbook"**
4. Name it: `Assessment_Metadata_Tracker.xlsx`
5. Open the file to confirm it was created
6. Close it (the app will create the worksheets automatically)

### Step 3.2: Register an Application in Azure AD

This allows the application to access your Excel file programmatically. **You may need UTA IT help for this section.**

**Step 3.2a: Access Azure Portal**

1. Go to **https://portal.azure.com/**
2. Sign in with your UTA Microsoft account (@uta.edu or @mavs.uta.edu)
3. In the search bar at the top, type **"App registrations"**
4. Click on **"App registrations"** under Services

**Step 3.2b: Create a New App Registration**

1. Click **"+ New registration"**
2. Fill in:
   - Name: `UTA Assessment Analyzer`
   - Supported account types: Select **"Accounts in this organizational directory only (UTA only - Single tenant)"**
   - Redirect URI: Leave blank
3. Click **"Register"**

**Step 3.2c: Note Your IDs**

After registration, you'll see an overview page. Copy these values to your text file:

1. **Application (client) ID**: Copy this value
   - Looks like: `12345678-abcd-1234-efgh-123456789012`
   - Label it in your text file as: `MS_CLIENT_ID`
   
2. **Directory (tenant) ID**: Copy this value
   - Same format as above
   - Label it as: `MS_TENANT_ID`

**Step 3.2d: Create a Client Secret**

1. In the left sidebar, click **"Certificates & secrets"**
2. Click **"+ New client secret"**
3. Description: `Assessment Analyzer Secret`
4. Expires: Select **"24 months"** (you'll need to renew in 2 years)
5. Click **"Add"**
6. **IMPORTANT**: Copy the **Value** column immediately (NOT the Secret ID)
   - You can only see this once!
   - Label it as: `MS_CLIENT_SECRET`

**Step 3.2e: Add API Permissions**

1. In the left sidebar, click **"API permissions"**
2. Click **"+ Add a permission"**
3. Click **"Microsoft Graph"**
4. Click **"Application permissions"** (not Delegated)
5. Search for and check these permissions:
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All`
6. Click **"Add permissions"**
7. **IMPORTANT**: Click the button **"Grant admin consent for University of Texas at Arlington"**
   - If this button is grayed out, you need IT admin help (see note below)
8. Confirm all permissions show green checkmarks under "Status"

> **Need IT Help?** If you can't grant admin consent, contact UTA IT and say:
> "I've created an App Registration called 'UTA Assessment Analyzer' in Azure AD. I need admin consent granted for Microsoft Graph application permissions: Files.ReadWrite.All and Sites.ReadWrite.All. My Application ID is [your client ID]."

### Step 3.3: Get the Drive ID and Item ID

You need to identify exactly where your Excel file is located in SharePoint/OneDrive.

**Step 3.3a: Using Graph Explorer**

1. Go to **https://developer.microsoft.com/en-us/graph/graph-explorer**
2. Click **"Sign in to Graph Explorer"** (use your UTA account)
3. Grant the requested permissions when prompted

**Step 3.3b: Find Your Drive ID**

If your file is in **OneDrive for Business**:

1. In the query box, replace the existing text with:
   ```
   https://graph.microsoft.com/v1.0/me/drive
   ```
2. Click **"Run query"**
3. In the Response Preview, find the line that says `"id":`
4. Copy that value - this is your **Drive ID**
5. Label it as: `MS_DRIVE_ID`

If your file is in **SharePoint**:

1. First, find your SharePoint site. Run this query:
   ```
   https://graph.microsoft.com/v1.0/sites?search=*
   ```
2. Find your site in the results and copy its `id`
3. Then run:
   ```
   https://graph.microsoft.com/v1.0/sites/{site-id}/drive
   ```
4. Copy the `id` from the response - this is your **Drive ID**

**Step 3.3c: Find Your File's Item ID**

1. Now run this query (replace {drive-id} with your actual Drive ID):
   ```
   https://graph.microsoft.com/v1.0/drives/{drive-id}/root/children
   ```
2. This shows all files/folders in the root
3. If your Excel file is in a subfolder, navigate deeper:
   ```
   https://graph.microsoft.com/v1.0/drives/{drive-id}/root:/{folder-name}:/children
   ```
4. Find your `Assessment_Metadata_Tracker.xlsx` file
5. Copy its `id` value - this is your **Item ID**
6. Label it as: `MS_ITEM_ID`

**Step 3.3d: Verify Your IDs**

Your text file should now have all 5 Microsoft values:
```
MS_CLIENT_ID = xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
MS_TENANT_ID = xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
MS_CLIENT_SECRET = your-secret-value-here
MS_DRIVE_ID = xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
MS_ITEM_ID = xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

---

## Part 4: Creating a GitHub Account

GitHub stores your application code and connects to Streamlit Cloud.

### Step 4.1: Create Account

1. Go to **https://github.com/**
2. Click **"Sign up"**
3. Enter your email
4. Create a password
5. Choose a username (e.g., `uta-ie-assessment`)
6. Complete the verification puzzle
7. Click **"Create account"**
8. Verify your email address

### Step 4.2: Create a Repository

A repository (repo) is like a folder for your project.

1. After logging in, click the **"+"** icon in the top right
2. Click **"New repository"**
3. Repository name: `assessment-analyzer`
4. Description: `Assessment Report Analyzer for UTA IE`
5. Select **"Private"** (important - keeps your code private)
6. Check **"Add a README file"**
7. Click **"Create repository"**

---

## Part 5: Uploading the Application

Now you'll upload the application files to GitHub.

### Step 5.1: Extract the ZIP File

You should have received `assessment-analyzer-v3.zip`.

**On Windows:**
1. Find the ZIP file in your Downloads folder
2. Right-click on it
3. Click **"Extract All..."**
4. Click **"Extract"**
5. A folder called `assessment-analyzer-v3` will appear

### Step 5.2: Upload Files to GitHub

1. Go to your GitHub repository (github.com/your-username/assessment-analyzer)
2. Click **"Add file"** → **"Upload files"**
3. Open the `assessment-analyzer-v3` folder on your computer
4. Drag ALL files from that folder into the GitHub upload area:
   - `app.py`
   - `requirements.txt`
   - `unit_registry_academic.csv`
   - `unit_registry_admin.csv`
   - `.gitignore`
   - `QUICK_REFERENCE.md`
   - `UTA_BRANDING.md`
   - etc.
5. Scroll down
6. In the "Commit changes" box, type: `Initial upload`
7. Click the green **"Commit changes"** button
8. Wait for the upload to complete

### Step 5.3: Upload the .streamlit Folder

1. Click **"Add file"** → **"Create new file"**
2. In the filename field, type: `.streamlit/config.toml`
3. Copy and paste this content:

```toml
[theme]
# UTA Brand Colors
primaryColor = "#0064b1"
backgroundColor = "#ffffff"
secondaryBackgroundColor = "#f5f7fa"
textColor = "#003865"
font = "sans serif"

[server]
maxUploadSize = 50

[browser]
gatherUsageStats = false
```

4. Click **"Commit changes"** → **"Commit changes"**

---

## Part 6: Deploying on Streamlit Cloud

Streamlit Cloud runs your application and makes it accessible via a web link.

### Step 6.1: Create Streamlit Account

1. Go to **https://share.streamlit.io/**
2. Click **"Sign in with GitHub"**
3. Authorize Streamlit to access your GitHub
4. Complete any additional setup steps

### Step 6.2: Deploy Your App

1. Click **"New app"** (or **"Create app"**)
2. Repository: Select `your-username/assessment-analyzer`
3. Branch: `main`
4. Main file path: `app.py`
5. Click **"Advanced settings"**

### Step 6.3: Add Your Secrets

In the Advanced settings, you'll see a "Secrets" text area. Paste this template and fill in your values:

```toml
# Anthropic API (from Part 2)
ANTHROPIC_API_KEY = "sk-ant-api03-paste-your-key-here"

# App Passwords (choose your own)
APP_PASSWORD = "choose-a-team-password"
ADMIN_PASSWORD = "choose-an-admin-password"

# Microsoft 365 / Azure AD (from Part 3)
MS_CLIENT_ID = "paste-your-client-id"
MS_CLIENT_SECRET = "paste-your-client-secret"
MS_TENANT_ID = "paste-your-tenant-id"
MS_DRIVE_ID = "paste-your-drive-id"
MS_ITEM_ID = "paste-your-item-id"
```

Replace each placeholder with your actual values.

Click **"Save"**

### Step 6.4: Deploy

1. Click **"Deploy!"**
2. Wait for the deployment (takes 2-5 minutes)
3. When ready, you'll see your app URL like: 
   `https://uta-ie-assessment-analyzer.streamlit.app`
4. **Save this URL** - this is how you and your team will access the app

---

## Part 7: Testing Your Application

### Step 7.1: Test Login

1. Go to your app URL
2. Enter your `APP_PASSWORD`
3. You should see the main interface with UTA branding

### Step 7.2: Test Admin Access

1. Log out (button in sidebar)
2. Enter your `ADMIN_PASSWORD`
3. You should see an additional "Configuration" tab

### Step 7.3: Test Report Analysis

1. Get a sample assessment report (PDF or Word)
2. Select "Results Report" from dropdown
3. Upload the file
4. Click "Analyze Report"
5. Wait 30-60 seconds
6. Review the analysis output on the right
7. Review the extracted metadata
8. Click "Save to Excel Online"

### Step 7.4: Verify Excel Online

1. Open your Excel file in SharePoint/OneDrive
2. You should see new worksheets:
   - `Results_Data`
   - `Improvement_Data` (created when needed)
   - `Plan_Data` (created when needed)
3. Verify that data was saved correctly

---

## Part 8: Using the Application

### Normal Analysis Workflow

1. **Go to app URL**
2. **Log in** with team password
3. **Select report type:**
   - Results Report: Full assessment with data
   - Improvement Report: Documents actions taken
   - Next Cycle Plan: Forward-looking plans
4. **Upload report** (PDF or Word)
5. **Click "Analyze Report"**
6. **Review analysis** (constructive feedback)
7. **Review metadata** (editable form)
8. **Correct any errors** in extracted metadata
9. **Save to Excel Online** or download as Excel file

### Batch Import (Historical Data)

Use this to load historical reports without running analysis.

1. Go to **"Batch Import"** tab
2. Select report type (all files must be same type)
3. Upload multiple files
4. Click **"Extract Metadata from All Files"**
5. Review extracted data
6. Click **"Save All to Excel Online"**

### Admin Functions

1. Log in with admin password
2. Go to **"Configuration"** tab
3. Available options:
   - **Rubric Guidance**: What criteria to evaluate
   - **Tone Instructions**: How feedback is written
   - **Analysis Prompts**: AI instructions for each report type
   - **Custom Rubric**: Upload your own evaluation criteria
   - **Unit Registry**: Manage canonical unit names

---

## Part 9: Troubleshooting

### "Excel Online Connection Failed"

**Cause**: Azure AD credentials incorrect or permissions not granted
**Fix**:
1. Double-check all 5 Microsoft values are correct
2. Verify admin consent was granted (green checkmarks in Azure)
3. Test the Drive ID and Item ID in Graph Explorer
4. Contact IT if admin consent button was grayed out

### "Invalid API Key" Error

**Cause**: Anthropic API key is incorrect
**Fix**: 
1. Go to console.anthropic.com
2. Create a new API key
3. Update in Streamlit secrets (Settings → Secrets)

### Data Not Appearing in Excel

**Cause**: Item ID or Drive ID incorrect
**Fix**:
1. Go back to Graph Explorer
2. Re-run the queries to find correct IDs
3. Make sure you're looking at the right file/folder
4. Update secrets in Streamlit

### "Module Not Found" Error

**Cause**: Missing dependency
**Fix**: Verify requirements.txt was uploaded to GitHub

### PDF Won't Process

**Cause**: PDF is scanned image (not text)
**Fix**: 
1. Use Word documents instead
2. Or OCR the PDF first using Adobe or another tool

---

## Part 10: Managing Costs

### Anthropic API Costs

| Report Size | Estimated Cost |
|-------------|----------------|
| Short (2-3 pages) | ~$0.03-0.04 |
| Medium (5-8 pages) | ~$0.06-0.08 |
| Long (10+ pages) | ~$0.10-0.12 |
| Batch import (no analysis) | ~$0.02-0.03 |

### Monthly Estimates

| Usage Level | Reports/Month | Est. Cost |
|-------------|---------------|-----------|
| Light | 50 | $3-5 |
| Moderate | 100 | $6-10 |
| Heavy | 250 | $15-25 |

### Setting Up Alerts

1. Go to console.anthropic.com
2. Click "Billing" → "Usage"
3. Set up alerts (e.g., email at $10, $20)

---

## Quick Reference Card

### Important URLs
| Resource | URL |
|----------|-----|
| Your App | (your Streamlit URL) |
| Anthropic Console | console.anthropic.com |
| Azure Portal | portal.azure.com |
| Graph Explorer | developer.microsoft.com/graph/graph-explorer |
| Streamlit Dashboard | share.streamlit.io |
| GitHub Repository | github.com/your-username/assessment-analyzer |

### Credentials to Keep Safe
- Anthropic API Key
- APP_PASSWORD (for team)
- ADMIN_PASSWORD (for you)
- MS_CLIENT_SECRET (Azure)

### Annual Maintenance
- Renew MS_CLIENT_SECRET every 24 months (Azure AD → App registrations → Certificates & secrets)
- Add Anthropic credits as needed

---

## Getting IT Help

For Azure AD setup, you may need IT assistance. Here's a template email:

---

**To**: IT Help Desk  
**Subject**: Azure AD App Registration - Admin Consent Needed

Hi,

I'm setting up an internal application for the Office of Institutional Effectiveness that needs to write assessment data to an Excel file in SharePoint.

I've created an App Registration in Azure AD with the following details:
- **Name**: UTA Assessment Analyzer
- **Application (Client) ID**: [paste your client ID]

I need admin consent granted for these Microsoft Graph **Application** permissions:
- Files.ReadWrite.All
- Sites.ReadWrite.All

Could you please grant admin consent for this application?

Thank you!

---

*Setup Guide Version 3.0 (Microsoft 365 Edition) - December 2025*
