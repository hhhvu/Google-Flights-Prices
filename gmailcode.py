from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import datefinder
import re
import pandas as pd
import numpy as np
import xlrd
import openpyxl
import datetime
from datetime import datetime as dt

import base64
import email
from apiclient import errors


args = tools.argparser.parse_args()
args.noauth_local_webserver = True

SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
store = file.Storage('credentials.json')
creds = store.get()

if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets('C:\\Users\\huong.vu\\Desktop\\Google-Flights-Prices\\client_secret.json',
                                          SCOPES)
    creds = tools.run_flow(flow, store, args)

service = build('gmail', 'v1', http=creds.authorize(Http()))

# calll the Gmail API, only get 1 of the recent message ids
# First get the message if for the message
results = service.users().messages().list(userId='me',
                                          maxResults=10,  # max record to  obtain
                                          q='from: noreply-travel@google.com label:inbox ').execute()  # include filter for message

time = []
price_now = []
price_before = []
airlines = []

for i in range(results['resultSizeEstimate']):
    # get the message id from the results object
    message_id = results['messages'][i]['id']

    # use the message id to get the actual message, including any attachments
    message = service.users().messages().get(userId='me', id=message_id).execute()

    for key in message:
        if key == 'snippet':
            # # get line that contains prices info
            # snippet_index = [m.start() for m in re.finditer('\$', message[key])]

            # # get prices from string
            # price_1 = message[key][(snippet_index[0]+1):(snippet_index[1] - 1)]
            # price_2 = message[key][(snippet_index[1] + 1) : (snippet_index[1] + 5)]

            # try:
            #     price_now.append(int(price_1))
            # except:
            #     price_now.append(price_1)
            
            # try:
            #     price_before.append(int(price_2))
            # except:
            #     price_before.append(price_2)

            # to get airline, first get position of colons:
            colon_index = [m.start() for m in re.finditer('\:', message[key])] #first colon is fixed, always go back from second colon
            # get posistion of word 'adult'
            adult_index = message[key].find('adult')
            # select a string from adult word to the second colon
            selected_string = message[key][adult_index:colon_index[1]].split()
            # getting two words before the last word 
            airline_name = selected_string[len(selected_string) - 3] + ' ' + selected_string[len(selected_string) - 2]
            airlines.append(airline_name)

        if key == 'payload':
            for name in message[key]:
                if name == 'headers':
                    for i in range(len(message[key][name])):
                        if message[key][name][i]['name'] == 'Subject':
                            # find '$' position
                            subject = message[key][name][i]['value']
                            price_indexes = [m.start() for m in re.finditer('\$', subject)]
                            # find '(' position
                            open_bracket_index = subject.find('(')
                            close_bracket_index = subject.find(')')
                            # selecting price
                            price_1 = subject[(price_indexes[0]+1):(open_bracket_index - 1)]
                            price_2 = subject[(price_indexes[1]+1):(close_bracket_index)]
                            price_now.append(price_1)
                            price_before.append(price_2)
                        if message[key][name][i]['name'] == 'Date':
                            # convert to date 
                            date = dt.strptime(message['payload']['headers'][i]['value'], '%a, %d %b %Y %H:%M:%S %z').strftime('%m/%d/%y')
                            time.append(date)


    # # get line that contains received date of email
    # date = message['payload']['headers'][2]['value']
    # # find date from the string
    # matches = datefinder.find_dates(date, strict=True)
    # # convert to time format
    # for match in matches:
    #     time.append(match.strftime('%m/%d/%Y'))

    # # get line that contains prices info
    # snippet_index = [m.start() for m in re.finditer('\$', message['snippet'])]
    # # get prices from string
    # price_now.append(message['snippet'][(snippet_index[0] + 1):(snippet_index[0] + 5)])
    # price_before.append(message['snippet'][(snippet_index[1] + 1):(snippet_index[1] + 5)])

# reorder data
time.reverse()
price_before.reverse()
price_now.reverse()

# bring result into data frame
data = pd.DataFrame(data={'time': time,
                          'price_now': price_now,
                          'price_before': price_before})

# open excel
workbook = openpyxl.load_workbook('C:\\Users\\huong.vu\\Desktop\\Google-Flights-Prices\\gmail_flight.xlsx')
# getting sheet in excel
sheet = workbook.get_sheet_by_name('Sheet1')
# getting last row in excel sheet
last_row = sheet.max_row

# getting historical data
hist_data = pd.DataFrame(data={'time': [sheet.cell(row=d, column=1).value for d in range(2, last_row + 1)],
                               'price_now': [sheet.cell(row=d, column=2).value for d in range(2, last_row + 1)],
                               'price_before': [sheet.cell(row=d, column=3).value for d in range(2, last_row + 1)]})

# remove duplicates
insert_data = data.join(hist_data, lsuffix='_new', rsuffix='_hist')
insert_data = insert_data[pd.isnull(insert_data['time_hist'])].iloc[:, 0:3]

# write data into excel file
for k in range(last_row + 1, len(insert_data) + last_row + 1):
    try:
        sheet.cell(row=k, column=1).value = insert_data['time_new'].iloc[k - last_row - 1]
        sheet.cell(row=k, column=2).value = insert_data['price_now_new'].iloc[k - last_row - 1]
        sheet.cell(row=k, column=3).value = insert_data['price_before_new'].iloc[k - last_row - 1]
    except Exception as e:
        print(str(e))

workbook.save('C:\\Users\\huong.vu\\Desktop\\Google-Flights-Prices\\gmail_flight.xlsx')
workbook.close()


message = service.users().messages().get(userId='me', id='16cfd5f696777359', format='raw').execute()
msg_str = base64.urlsafe_b64decode(message['raw'].replace('-_', '+/').encode('ASCII'))
mime_msg = email.message_from_bytes(msg_str)
messageMainType = mime_msg.get_content_maintype()
if messageMainType == 'multipart':
    for part in mime_msg.get_payload():
        print(part.get)
        if part.get_content_maintype() == 'text':
            output = part.get_payload()
    return ""
elif messageMainType == 'text':
    return mime_msg.get_payload()