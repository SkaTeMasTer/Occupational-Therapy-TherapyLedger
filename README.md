# OT Billing System - Complete Setup Guide

Oppupational theropy practice, it's a spreadsheet with scripts, everything anyone with a practice could use to keep track of clients, creates invoices that get attached in email, checks for new paid invoices, keeps detailed graphs and stats for /weekly/month/year.




## 📦 WHAVENT'F Finishedd YET






















## 📦 What's Included


---

## 🚀 Step-by-Step Setup

### Step 1: Upload Files to Google Drive

1. Upload all files to Google Drive
2. Right-click each `.xlsx` file → "Open with" → Google Sheets
3. Right-click the `.docx` file → "Open with" → Google Docs
4. This converts them to native Google format

5. Save the script (Ctrl+S)
---


### Adding Clients

1. Open the **Clients** sheet
2. Add client info in each row:
   - Client ID (unique, like C001, C002)
   - Client Name (must match exactly when entering billable hours)
   - Email, Phone, Address, City, State
   - Client Since date

### Logging Billable Hours

1. Open **Billable Hours** sheet
2. Enter in row 3 or below:
   - **Client Name**: Type exactly as it appears in Clients sheet
   - **Client ID**: Auto-fills via formula
   - **Period Start/End**: Date range for services
   - **Hours**: Number of hours worked
   - **Hourly Rate**: Your rate for this client
   - **Total**: Auto-calculates (Hours × Rate)

### Sending Invoices

**Single Invoice:**
1. Click on any cell in the row you want to invoice
2. Menu → 📧 Invoice System → "Send Invoice for Selected Row"
3. The "Sent ✓" column shows ☑ when sent

**Mark as Paid:**
1. Select the row
2. Menu → 📧 Invoice System → Payment → "Mark Selected as Paid"

The dashboard shows:

- **Monthly Income**: Revenue by month with bar chart
- **Weekly Income**: Week-by-week trends with line chart
- **Yearly Income**: Annual comparisons
- **Client Analysis**: Revenue per client with pie chart
- **Payment Status**: Paid vs unpaid tracking


---

## 📊 Understanding the Analytics

### Key Metrics Explained

| Metric | What It Shows |
|--------|---------------|
| Total Revenue | Sum of all invoiced amounts |
| Paid Revenue | Sum of invoices marked as PAID |
| Unpaid Revenue | Total - Paid (outstanding balance) |
| Collection Rate | Percentage of revenue collected |
| Total Hours | All hours billed |
| Avg Hourly Rate | Revenue ÷ Hours |

### Charts Available

- **Monthly Revenue Bar Chart**: Compare months visually
- **Weekly Trend Line Chart**: See patterns over time
- **Revenue by Client Pie Chart**: Who generates the most revenue
- **Paid vs Unpaid Pie Chart**: Collection status overview

---

## ⚙️ Customization

### Change Invoice Appearance

1. Open **Invoice Template** in Google Sheets
2. Modify colors, fonts, layout
3. Add your logo (Insert → Image)
4. Keep `{{PLACEHOLDER}}` format for dynamic values

### Change Email Wording

1. Open **Email Template** in Google Docs
2. Edit the text as desired
3. Keep placeholders like `{{CLIENT_NAME}}` intact

### Available Placeholders

| Placeholder | Replaced With |
|-------------|---------------|
| `{{INVOICE_NUMBER}}` | Auto-generated invoice number |
| `{{CLIENT_NAME}}` | Client's name |
| `{{CLIENT_EMAIL}}` | Client's email |
| `{{PERIOD_START}}` | Billing period start date |
| `{{PERIOD_END}}` | Billing period end date |
| `{{HOURS}}` | Total hours billed |
| `{{RATE}}` | Hourly rate |
| `{{TOTAL_DUE}}` | Invoice total |
| `{{DUE_DATE}}` | Payment due date |
| `{{BUSINESS_NAME}}` | Your practice name |
| `{{BUSINESS_PHONE}}` | Your phone |
| `{{BUSINESS_EMAIL}}` | Your email |

---

## ❓ Troubleshooting

### "Client not found" error
- Client name in Billable Hours must **exactly match** the name in Clients sheet
- Check for extra spaces or different capitalization

### Email not sending
- Verify client has a valid email in Clients sheet
- Check your spam/junk folder
- Make sure you authorized the script

### "Mark Paid" checkbox not working
- Run `installTriggers` from the script editor
- Or manually add an onEdit trigger

### Menu not showing
- Refresh the page
- Run the `onOpen` function from script editor
- Clear browser cache

### Charts not updating
- Click "Refresh Dashboard Data" in the menu
- Or close and reopen the dashboard

### PDF looks wrong
- Check that placeholders weren't deleted in template
- Verify the Invoice sheet name matches CONFIG

---

## 💡 Pro Tips

1. **Backup regularly**: Export your Billable Hours sheet periodically
2. **Consistent naming**: Always use the same client name spelling
3. **Date format**: Use MM/DD/YYYY format for dates
4. **Rate changes**: Update rates in Billable Hours, not Clients sheet
5. **Archive old data**: Move completed years to a separate sheet

---

## 🆘 Need Help?

The script is fully customizable. Common modifications:
- Add more service line items per invoice
- Include CPT codes for insurance
- Add payment tracking with dates
- Generate end-of-year tax reports
- Set up automatic payment reminders

Feel free to modify the code to fit your specific practice needs!
