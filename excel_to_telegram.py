import json
import logging
import openpyxl
import os
import prettytable as pt
import sys
from telethon import TelegramClient
import telethon.sync

# pip install openpyxl prettytable telethon

name="excel_to_telegram"
xlsm="excel_to_telegram.xlsm"

# Parameters - will be taken from the Excel params sheet
API_ID=""
API_HASH=""
Channel=""

# Change to the script directory
os.chdir(sys.path[0])

logging.basicConfig(
  filename='excel_to_telegram.log',
  # encoding='utf-8',
  format='%(asctime)s %(levelname)s:%(message)s',
  level=logging.DEBUG
)
logging.debug("Logging activated")

# Log Command line parameters
logging.debug(f"args:{len(sys.argv)}")
for arg in sys.argv:
  logging.debug(f"arg:{arg}")

# Get the data
data=sys.argv[1]
logging.debug(f"data: {data}")
print(f"data: {data}")

def create_table(cells):

  logging.debug(f"create_table - cells: {cells}")

  # Use the first row of the cells as the headers of the prettytable
  r=cells[0]
  headers=[]
  for k in r:
    print(f"{k} - {r[k]}")
    headers.append(r[k])

  # Create a formattable table for Telegram
  table = pt.PrettyTable(headers)
  for k in r:
    table.align[r[k]] = 'l'

  # Add the Array data to the prettytable
  n=0
  for r in cells:
    n+=1
    if n==1: continue
    a=[]
    for k in r:
      # logging.debug(f"{k} - {r[k]}\r\n")
      # print(f"{k} - {r[k]}")
      a.append(r[k])
    table.add_row(a)
  return table
# create_table

def get_param(name,wb):
  v=""
  dest = wb.defined_names[name].destinations
  for title,coord in dest:
    range = wb[title][coord]
    v=range.value
  return v
# get_param

def get_params():
  global API_ID
  global API_HASH
  global Channel
  logging.debug(f"get_params from {xlsm}")
  wb = openpyxl.load_workbook(xlsm)
  API_ID=get_param('API_ID',wb)
  API_HASH=get_param('API_HASH',wb)
  Channel=get_param('Channel',wb)
  # logging.debug(f"Params: {API_ID} {API_HASH} {Channel}")
  # print(f"Params: {API_ID} {API_HASH} {Channel}")
  wb.close()
# get_params

if __name__ == '__main__':

  # Get the parameters
  get_params()

  # Connect to Telegram
  # logging.debug("TelegramClient {API_ID} {API_HASH} {Channel}")
  client = TelegramClient(name, API_ID, API_HASH)
  
  try:
  
    # Start the connection
    logging.debug("client.start")
    client.start()
      
    if sys.argv[1].lower()!='init':
      # Get json from data
      # logging.debug("get json=cells")
      cells=json.loads(data)

      # Create the table
      # logging.debug("create_table")
      table=create_table(cells)
      
      # Get the entity of the selected channel
      # logging.debug(f"get_entity {Channel}")
      entity = client.get_entity(Channel)
      logging.debug(f"got entity {entity} for Channel {Channel}")

      # Send the table to Telegram
      client.send_message(entity, 
      f'<pre>{table}</pre>',parse_mode='html')
      
      logging.debug("Sending Done.")
    
  finally:
    print("Disconnect...")
    client.disconnect()
