# Sample 02: Order & Payment Follow-up Tracker (Google Sheets + Apps Script)

**This is a portfolio sample with fictional data.** Customer names and phone numbers are made up (the phone numbers are masked).

## The problem it solves
Small distributors, wholesalers and service businesses often track orders and pending payments in a spreadsheet that nobody updates, and they chase overdue payments from memory. This tracker:

| Feature | How |
|---|---|
| Live balance and days-overdue per order | Sheet formulas |
| Overdue rows turn red automatically | Conditional formatting, plus `highlightOverdue()` |
| Dashboard: sales, collected, outstanding, overdue, collection % | COUNTIFS / SUMIFS |
| Overdue ageing (1-15 / 16-30 / 31+ days) and outstanding by customer | COUNTIFS / SUMIFS / SUMPRODUCT |
| "Last Updated" timestamp when status or payment changes | `onEdit` trigger |
| Marking a payment "Paid" fills in the paid amount | `onEdit` trigger |
| **Daily 9 AM email** listing overdue customers, phone numbers and balances | Time-driven trigger, `MailApp` |
| One-click month-wise summary sheet | `buildMonthlySummary()` |
| Custom **Automation** menu | `onOpen()` |

## Files
- `Sample_02_Order_Payment_Tracker.xlsx`: Dashboard, Orders and Lists sheets
- `Code.gs`: the Apps Script automation
- `preview_Dashboard.png`, `preview_Orders.png`: screenshots

## Turning it into a live Google Sheet (about 5 minutes)
1. In Google Drive, choose **New → File upload**, upload the .xlsx, then open it with Google Sheets. Use **File → Save as Google Sheets** if Drive asks.
2. Open **Extensions → Apps Script**, delete the default code, paste in `Code.gs`, and save.
3. Reload the sheet. An **Automation** menu appears.
4. Choose **Automation → Turn on daily 9 AM automation** and approve the permissions. The prompt is Google's standard authorisation for your own script.
5. Test it with **Automation → Email me the overdue digest now**.
