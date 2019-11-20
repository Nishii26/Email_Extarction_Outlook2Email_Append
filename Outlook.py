import win32com.client
import os
import time
import datetime as dt
import xlrd
from xlutils.copy import copy as copy


# this is set to the current time
date_time = dt.datetime.now()
# this is set to one hour ago
lastHourDateTime = dt.datetime.now() - dt.timedelta(hours = 20)
#This is set to one minute ago; you can change timedelta's argument to whatever you want it to be
# lastMinuteDateTime = dt.datetime.now() - dt.timedelta(minutes = 1)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)


# retrieve all emails in the inbox, then sort them from most recently received to oldest (False will give you the reverse). Not strictly necessary, but good to know if order matters for your search
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)

# restrict to messages from the past hour based on ReceivedTime using the dates defined above.
# lastHourMessages will contain only emails with a ReceivedTime later than an hour ago
# The way the datetime is formatted DOES matter; You can't add seconds here.
lastHourMessages = messages.Restrict("[ReceivedTime] >= '" +lastHourDateTime.strftime('%m/%d/%Y %H:%M %p')+"'")

# lastMinuteMessages = messages.Restrict("[ReceivedTime] >= '" +lastMinuteDateTime.strftime('%m/%d/%Y %H:%M %p')+"'")
print(len(lastHourMessages))

print("Current time: "+date_time.strftime('%m/%d/%Y %H:%M %p'))
print("Messages from the past hour:")

# for message in lastHourMessages:
#     print(message.Subject)
#     print(message.ReceivedTime)

# print ("Messages from the past minute:")

# for message in lastMinuteMessages:
#     print(message.Subject)
#     print(message.ReceivedTime)

# # GetFirst/GetNext will also work, since the restricted message list is just a shortened version of your full inbox.
# print ("Using GetFirst/GetNext")
message = lastHourMessages.GetFirst()
col=0
workbook = xlrd.open_workbook("C:\\Users\\Innovation\\Desktop\\Email_Output.xls")
workbook_sheet = workbook.sheet_by_index(0)

# worksheet.write(0,5,"Received Time")
row = workbook_sheet.nrows
wb = copy(workbook)
worksheet = wb.get_sheet(0)
while message:
    worksheet.write(row, col, row)
    worksheet.write(row, col + 1,message.SenderName)
    try:
        worksheet.write(row, col+2,message.To)
    except:
        worksheet.write(row,col+2,"")
    worksheet.write(row, col+3,message.Subject)
    worksheet.write(row, col+4,message.Body)
    worksheet.write(row,col+5,message.ReceivedTime.replace(tzinfo=None))
    row = row + 1
    # print(message.Subject)
    #print(message.Body)
    # print(message.SenderName)
    # print(message.To)
    #print(message.Recipient)
    # #print(message.Sender.Address)
    # print(message.ReceivedTime)
    # d = message.ReceivedTime
    # print(d.replace(tzinfo=None))
    message = lastHourMessages.GetNext()
    wb.save("C:\\Users\\Innovation\\Desktop\\Email_Output.xls")