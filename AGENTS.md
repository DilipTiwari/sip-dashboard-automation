\# Instructions for Codex



This is a Google Apps Script project connected to a Google Sheet.



\## Important Rule



Google Apps Script treats all .gs files as one project. Avoid duplicate global constants and duplicate function names.



\## Do Not Rename These Sheet Names



\- SIP Daily Tracker

\- SIP Dashboard

\- Weekly SIP Data

\- master data

\- SIP

\- Monthly Ledger

\- Current financial Detail



\## Keep These Flows Separate



1\. Daily tracker flow:

&#x20;  - recordDailyTracker



2\. Dashboard email flow:

&#x20;  - setupUltraDashboard

&#x20;  - sendWeeklyDashboardMail



3\. Separate tracker graph email flow:

&#x20;  - sendWeeklySipTrackerGraphMail



4\. NAV/value update flow:

&#x20;  - updateMutualFundNAV

&#x20;  - updateNiftyValue



\## Important



Do not merge the dashboard email and separate tracker graph email unless explicitly asked.



Do not change email recipients unless explicitly asked.



When editing triggers, only trigger the actual worker functions, not installer functions.

