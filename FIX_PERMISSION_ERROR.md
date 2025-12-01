# Fix Permission Error - 4D Active Email Script

## 🔑 Service Account Email

Your service account email is:
```
emo-reports-automation@single-frame-467107-i1.iam.gserviceaccount.com
```

---

## 📋 Steps to Fix Permission Error

### Step 1: Open Your Google Sheet

Open this spreadsheet:
```
https://docs.google.com/spreadsheets/d/1yexwDDUJn6frYet7AaoKkJRjqFQp_F40MtrxpOCKkv8/edit
```

### Step 2: Share with Service Account

1. **Click the "Share" button** (top right corner)
2. **In the "Add people and groups" field**, paste this email:
   ```
   emo-reports-automation@single-frame-467107-i1.iam.gserviceaccount.com
   ```
3. **Set permission to "Viewer"** (or "Editor" if you need write access)
4. **Uncheck "Notify people"** (service accounts don't need notifications)
5. **Click "Share"**

### Step 3: Verify Access

After sharing, run the script again:
```bash
python 4d_active_email.py
```

It should now work! ✅

---

## 🔍 Why This Error Happened

The service account needs explicit permission to access your Google Sheet. Even though it's a service account, Google Sheets requires you to share the sheet with it, just like you would share with a regular user.

---

## ✅ After Sharing

Once you've shared the sheet, the script will be able to:
- ✅ Read data from range A1:V23
- ✅ Filter columns
- ✅ Calculate Attrition
- ✅ Send email

---

## 📝 Note

This is the same service account used for your other automations (Flipkart Myntra DN Analysis, etc.). If you've already shared other sheets with this email, you might just need to share this specific sheet.

