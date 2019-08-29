from apiclient.discovery import build
from httplib2 import Http 
from oauth2client import file, client, tools
import datefinder
import re 
import pandas as pd
import numpy as np
import xlwt
import xlrd
from xlutils.copy import copy
import xlsxwriter
import openpyxl

args = tools.argparser.parse_args()
args.noauth_local_webserver = True

SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
store = file.Storage('credentials.json')
creds = store.get()

if not creds or creds.invalid:
	flow = client.flow_from_clientsecrets('C:\\Users\\huong.vu\\Desktop\\Gmail Flight\\client_secret.json', SCOPES)
	creds = tools.run_flow(flow, store, args)

service = build('gmail','v1', http=creds.authorize(Http()))

#calll the Gmail API, only get 1 of the recent message ids
# First get the message if for the message
results = service.users().messages().list(userId='me',
											maxResults = 10,  # max record to  obtain
											q='from: noreply-travel@google.com label:inbox ').execute() #include filter for message

time = []
price_now = []
price_before = []

for i in range(results['resultSizeEstimate']):
	# get the message id from the results object
	message_id = results['messages'][i]['id']

	# use the message id to get the actual message, including any attachments
	message = service.users().messages().get(userId='me',id=message_id).execute()
	
	# get line that contains received date of email
	date = message['payload']['headers'][2]['value']
	# find date from the string
	matches = datefinder.find_dates(date, strict = True)
	# convert to time format
	for match in matches:
		time.append(match.strftime('%m/%d/%Y'))

	# get line that contains prices info
	snippet_index = [m.start() for m in re.finditer('\$',message['snippet'])]
	# get prices from string
	price_now.append(message['snippet'][(snippet_index[0]+ 1):(snippet_index[0]+5)])
	price_before.append(message['snippet'][(snippet_index[1]+ 1):(snippet_index[1]+5)])

# reorder data
time.reverse()
price_before.reverse()
price_now.reverse()

# bring result into data frame
data = pd.DataFrame(data = {'time': time,
							'price_now': price_now,
							'price_before': price_before})

# open excel
workbook = openpyxl.load_workbook('C:\\Users\\huong.vu\\Desktop\\Gmail Flight\\gmail_flight.xlsx')
# getting sheet in excel
sheet = workbook.get_sheet_by_name('Sheet1')
# getting last row in excel sheet
last_row = sheet.max_row

# getting historical data
hist_data = pd.DataFrame(data = {'time': [sheet.cell(row = d, column = 1).value for d in range(2,last_row+1)],
									'price_now': [sheet.cell(row = d, column = 2).value for d in range(2,last_row+1)],
									'price_before': [sheet.cell(row = d, column = 3).value for d in range(2,last_row+1)]})

# remove duplicates
insert_data = data.join(hist_data, lsuffix = '_new', rsuffix = '_hist')
insert_data = insert_data[pd.isnull(insert_data['time_hist'])].iloc[:,0:3]


# write data into excel file
for k in range(last_row + 1, len(insert_data)+last_row+1):
	try:
		sheet.cell(row = k, column = 1).value =  insert_data['time_new'].iloc[k - last_row-1]
		sheet.cell(row = k, column = 2).value =  insert_data['price_now_new'].iloc[k - last_row-1]
		sheet.cell(row = k, column = 3).value =  insert_data['price_before_new'].iloc[k - last_row-1]
	except Exception as e:
		print(str(e))

workbook.save('C:\\Users\\huong.vu\\Desktop\\Gmail Flight\\gmail_flight.xlsx')
workbook.close()

