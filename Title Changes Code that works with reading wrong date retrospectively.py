import openpyxl
import pandas
import win32com.client
import pandas as pd
import os
from datetime import datetime, timedelta

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
gal = outlook.Session.GetGlobalAddressList()
entries = gal.AddressEntries

data = []
for entry in entries:
    if entry.Type == "EX":
        user = entry.GetExchangeUser()
        if user is not None:
            data.append([user.Name, user.JobTitle])

df = pd.DataFrame(data, columns=['Name', 'Title'])

today = datetime.today().strftime('%Y-%m-%d')
filename = f'Titles_{today}.xlsx'
df.to_excel(filename, index=False)

df1 = pd.read_excel(filename)

today_string = '2023-3-21'

today = datetime.strptime(today_string, '%Y-%m-%d')

yesterday = today - timedelta(days=1)

yesterday_file = f'Titles_{yesterday.strftime("%Y-%m-%d")}.xlsx'

df2 = pd.read_excel(yesterday_file)

difference = pd.merge(df1, df2, how='outer', indicator=True).query('_merge != "both"').drop(columns='_merge')

difference.to_excel('Title Changes.xlsx')

df = pd.read_excel('Title Changes.xlsx')

column = 'Name'

mask = ~df.duplicated(subset=column, keep=False)

df = df[~mask]

df.to_excel('Title Changes For Today.xlsx', index=False)

outlook = win32com.client.Dispatch('Outlook.Application')

namespace = outlook.GetNamespace('MAPI')

mail = outlook.CreateItem(0)

mail.Subject = 'Test'
mail.Body = 'This is a test for the automated title change report'
mail.To = 'pehlinger@customtruck.com'

attachment = r'C:\Users\pehlinger\PycharmProjects\pythonProject1\Title Changes For Today.xlsx'
mail.Attachments.Add(attachment)

mail.Send()