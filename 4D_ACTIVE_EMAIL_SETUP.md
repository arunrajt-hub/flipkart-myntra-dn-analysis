# 4D Active Email - GitHub Actions Setup Guide

This guide will help you set up the 4D Active Email automation to run on GitHub Actions, eliminating the need for n8n.

---

## 📋 What This Script Does

This script replicates your n8n "4D Active eMail" workflow:

1. **Reads Google Sheets data** from range `A1:V23`
2. **Filters specific columns** (indices: 1, 2, 18, 19, 20, 21, 22)
3. **Calculates Attrition** = Column[19] - Column[20]
4. **Creates styled HTML table** with inline CSS
5. **Sends email** via Gmail to:
   - **To:** arunraj@loadshare.net
   - **CC:** lokeshh@loadshare.net, maligai.rasmeen@loadshare.net, rakib@loadshare.net

---

## 🚀 Setup Steps

### Step 1: Add GitHub Secrets

Go to your GitHub repository → **Settings** → **Secrets and variables** → **Actions** → **New repository secret**

Add these secrets:

#### 1. `SERVICE_ACCOUNT_JSON`
- **Value:** Your Google Service Account JSON (same as used for Flipkart Myntra DN Analysis)
- **How to get:** Copy the entire contents of `service_account_key.json`

#### 2. `GMAIL_SENDER_EMAIL`
- **Value:** `arunraj@loadshare.net`
- **Note:** This is the email address that will send the emails

#### 3. `GMAIL_APP_PASSWORD`
- **Value:** Your Gmail App Password
- **How to get:**
  1. Go to Google Account → Security
  2. Enable 2-Step Verification (if not already enabled)
  3. Go to App Passwords
  4. Generate a new app password for "Mail"
  5. Copy the 16-character password (no spaces)

---

### Step 2: Verify Files Are in Repository

Make sure these files are in your repository:

- ✅ `4d_active_email.py` (main script)
- ✅ `requirements_4d_active_email.txt` (dependencies)
- ✅ `.github/workflows/4d_active_email.yml` (GitHub Actions workflow)

---

### Step 3: Verify Google Sheets Access

The script needs access to:
- **Spreadsheet ID:** `1yexwDDUJn6frYet7AaoKkJRjqFQp_F40MtrxpOCKkv8`
- **Sheet ID:** `851267488`
- **Range:** `A1:V23`

Make sure your Google Service Account has **View** access to this spreadsheet.

---

### Step 4: Test the Workflow

1. Go to your GitHub repository
2. Click **Actions** tab
3. Find **"4D Active Email Automation"** workflow
4. Click **"Run workflow"** → **"Run workflow"** (manual trigger)
5. Wait for it to complete
6. Check your email!

---

## ⏰ Schedule

The workflow is currently set to run:
- **Once daily at 9:00 AM UTC** (2:30 PM IST)

### To Change Schedule

Edit `.github/workflows/4d_active_email.yml`:

```yaml
schedule:
  # Run once daily at 9:00 AM UTC (2:30 PM IST)
  - cron: '0 9 * * *'
```

**Common schedules:**
- Every 12 hours: `'0 */12 * * *'`
- Every 8 hours: `'0 */8 * * *'`
- Twice daily (10 AM and 6 PM IST): `'30 4,12 * * *'` (4:30 AM and 12:30 PM UTC)

---

## 🔍 How It Works

### Workflow Flow

```
GitHub Actions Trigger
    ↓
Read Google Sheets (A1:V23)
    ↓
Filter Columns [1,2,18,19,20,21,22]
    ↓
Calculate Attrition = Column[19] - Column[20]
    ↓
Create Styled HTML Table
    ↓
Send Email via Gmail
```

### Column Filtering

The script keeps only these columns (by index):
- **Index 1:** Second column
- **Index 2:** Third column
- **Index 18:** 19th column
- **Index 19:** 20th column
- **Index 20:** 21st column
- **Index 21:** 22nd column
- **Index 22:** 23rd column
- **Attrition:** Calculated as Column[19] - Column[20]

### HTML Styling

The HTML table is styled with inline CSS (matching n8n workflow):
- Table: Border, padding, Arial font, 13px
- Headers: Light gray background (#f6f8fa)
- Cells: Borders and padding

---

## 🐛 Troubleshooting

### Error: "Service account key file was not created"
- ✅ Check that `SERVICE_ACCOUNT_JSON` secret is set correctly
- ✅ Make sure the JSON is complete (copy entire file)

### Error: "Permission denied" when reading Google Sheets
- ✅ Verify Service Account has access to the spreadsheet
- ✅ Share the spreadsheet with the Service Account email

### Error: "Authentication failed" when sending email
- ✅ Check `GMAIL_APP_PASSWORD` secret is correct
- ✅ Make sure you're using App Password (not regular password)
- ✅ Verify 2-Step Verification is enabled on Gmail account

### Error: "No data found"
- ✅ Verify spreadsheet ID is correct
- ✅ Verify sheet ID is correct
- ✅ Check that range A1:V23 has data

---

## 📊 Comparison: n8n vs GitHub Actions

| Feature | n8n Workflow | GitHub Actions Script |
|--------|--------------|----------------------|
| **Runs on** | Your laptop (needs ngrok) | GitHub's cloud |
| **Laptop needed?** | Yes (24/7) | No |
| **ngrok needed?** | Yes | No |
| **Cost** | Free (self-hosted) | Free (GitHub Actions) |
| **Reliability** | Depends on laptop | High (cloud) |
| **Schedule** | n8n scheduler | GitHub Actions cron |

---

## ✅ Benefits

1. **No laptop needed** - Runs entirely in GitHub's cloud
2. **No ngrok needed** - No need to expose local services
3. **More reliable** - Runs even if your laptop is off
4. **Free** - GitHub Actions free tier is sufficient
5. **Easy to modify** - Just edit the Python script

---

## 📝 Notes

- The script uses the same Google Service Account as your other workflows
- Email is sent from `arunraj@loadshare.net` (configured in secrets)
- The schedule can be changed anytime by editing the workflow file
- You can manually trigger the workflow from GitHub Actions UI

---

## 🎉 You're Done!

Once set up, the workflow will:
- ✅ Run automatically on schedule
- ✅ Send emails without needing your laptop
- ✅ Work even when your laptop is off

No more n8n needed for this workflow! 🚀

