# excel-scheduled-refresh
This Python script is designed to automate the refresh of Excel workbooks on a scheduled basis using Windows Task Scheduler.

-The script opens a specified Excel file, triggers a refresh of all data connections, queries, and pivot tables, and saves the workbook.
-By pairing the script with Task Scheduler, you can configure automatic refreshes at specific times (e.g., daily at 8 AM).
-This is especially useful for workbooks that pull data from databases, APIs, or Power Query sources, ensuring that reports always have the most up-to-date information without requiring manual refreshes.

Key Features:

- Refreshes all data connections within the Excel workbook.
- Runs silently in the background — no manual Excel interaction required.
- Compatible with Task Scheduler for fully automated refresh cycles.
- Useful for automating recurring reporting workflows.

Typical Use Case:
Automating the daily refresh of a sales dashboard or financial report in Excel so that stakeholders always see the latest data when opening the file.
