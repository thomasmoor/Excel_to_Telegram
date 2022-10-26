# Excel_to_Telegram

This tool lets you post directly a table of data from Excel into Telegram

Open the Excel file excel_to_telegram.xlsm
Open your own spreadsheet
Select the range you want to post on Telegram
Click on Add-ins in your toolbar
Then click on ThomasMoor
And finally Table to Telegram
You should see a popup window appear and then disappear
And you should see your table appear in Telegram

Pre-Requisites:
- Python installed
- Dowload code from this Github repository
- Connect and get API parameters from my.telegram.org
- Set the parameters in the Params sheet of the Excel workbook:
  - API_ID   (from my.telegram.org)
  - API_HASH (from my.telegram.org)
  - Channel: URL from Telegram app > Channel you want to post in
  - Script Directory: Directory where you downloaded this code to
  - Python executable full path (from venv\scripts if you install in a virtual environment - recommnded)
  - Python script: you should not have to change this one
