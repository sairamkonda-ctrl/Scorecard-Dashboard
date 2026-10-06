# Scorecard dashboard — Nigeria access and reports

## Changes
- Yusuf Aderinto and Samuel Samuel can browse Nigeria team and individual performance without needing a personal Team List row. Server responses omit other countries' identities, metrics, ratings, leave records, and benchmark totals.
- Existing lead and individual access is retained. Unknown sign-in identities are denied instead of falling back to the script owner.
- Nigeria performance reports support daily, weekly, monthly, and quarterly delivery, including the existing HTML report and PDF attachment.
- To: yusuf.aderinto@mediamint.com, samuel.samuel@mediamint.com.
- CC: avinash.vellore@mediamint.com, pavan.davuluri@mediamint.com, sharath.upadhyay@mediamint.com.
- Report sending and schedule installation require a configured LEAD_EMAILS account.

## Apply to Google Apps Script
1. Replace Code.gs and Index.html in the existing spreadsheet's Apps Script project with these files.
2. Confirm Nigeria employees have Country = Nigeria (or Region = Nigeria) in Team List. Blank or other country values are excluded.
3. Update the web app deployment to the new version. Restrict access to your organization, and ensure signed-in visitor emails are available. If a visitor email is unavailable, access is denied; do not restore the old owner-email fallback. The two users must sign in with the exact addresses above. Spreadsheet sharing is not required for an execute-as-owner web app.
4. As the configured lead account (sairam.konda@mediamint.com), run testNigeriaDigestDaily, testNigeriaDigestWeekly, testNigeriaDigestMonthly, and testNigeriaDigestQuarterly. Preview reports go only to the existing DIGEST_PREVIEW_EMAIL; check that setting in Code.gs.
5. Run installNigeriaDigestTriggers once as that lead and authorize the requested Google permissions. Re-running replaces that account's Nigeria triggers; remove any older Nigeria triggers installed by other accounts to avoid duplicate schedules.

Schedules use the spreadsheet timezone (normally Asia/Kolkata), at approximately 09:15: daily for the previous day; Monday for the previous Monday–Sunday; the 2nd for the previous month; and January/April/July/October 2nd for the previous quarter. Apps Script time triggers are approximate. The Email Status tab records delivery and existing duplicate-period checks remain in place.

## Validation
Local mocked checks passed for both Nigeria viewers, unknown and blank identities, ordinary-member and lead access, country filtering/index remapping, Nigeria benchmark totals, recipient counts, and JavaScript syntax. Live Google authentication, email delivery, PDF generation, and trigger execution require verification in the deployed Apps Script project. No emails were sent and no live deployment was changed during this update.
