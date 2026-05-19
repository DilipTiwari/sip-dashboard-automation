\# SIP Dashboard Automation



This is a Google Sheets + Apps Script automation project for SIP portfolio tracking.



\## Main Sheets



\- SIP Daily Tracker

\- SIP Dashboard

\- Weekly SIP Data

\- master data

\- SIP

\- Monthly Ledger

\- Current financial Detail



\## Main Functions



\- recordDailyTracker - records daily SIP tracker data

\- setupUltraDashboard - refreshes SIP Dashboard

\- sendWeeklyDashboardMail - sends weekly dashboard email body

\- sendWeeklySipTrackerGraphMail - sends separate SIP Daily Tracker graph email

\- updateMutualFundNAV - updates mutual fund NAV data

\- updateNiftyValue - updates Nifty value

\- sendPortfolioAlert - sends daily portfolio alert



\## Trigger Setup



Recommended active triggers:



\- updateNiftyValue

\- updateMutualFundNAV

\- recordDailyTracker

\- sendPortfolioAlert

\- sendWeeklyDashboardMail

\- sendWeeklySipTrackerGraphMail



Do not create triggers for installer functions like:



\- installSundayDashboardMailTrigger

\- installSundaySipTrackerGraphTrigger

