from win32com.client import Dispatch

# Starts Excel COM automation
xl = Dispatch("Excel.Application")

# Runs Excel in the background
xl.visible = False

# Opens the workbook
workbook = xl.Workbooks.open("file_path.xlsx")

# Refreshes all data connections, Power Queries, and pivots
workbook.RefreshAll()

# Waits until async queries (like Power Query, data model refreshes) are complete
xl.CalculateUntilAsyncQueriesDone()

# Closes the workbook and saves changes
workbook.Close(SaveChanges=1)

# Quits Excel completely
xl.Quit()
