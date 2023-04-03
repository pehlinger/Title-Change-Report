import win32com.client
import pandas as pd
from datetime import datetime, timedelta, date

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

today_string = date.today()

past = today_string - timedelta(days=7)

past_file = f'Titles_{past.strftime("%Y-%m-%d")}.xlsx'

df2 = pd.read_excel(past_file)

difference = pd.merge(df1, df2, how='outer', indicator=True).query('_merge != "both"').drop(columns='_merge')

difference.to_excel('Title Changes.xlsx')

df = pd.read_excel('Title Changes.xlsx')

column = 'Name'

mask = ~df.duplicated(subset=column, keep=False)

df = df[~mask]

df.to_excel('Title Changes For Today.xlsx', index=False)

df = pd.read_excel('Title Changes For Today.xlsx')

df = df[df.duplicated(subset='Name', keep=False)]

df.to_excel('Title Changes For Today.xlsx', index=False)

df = pd.read_excel('Title Changes For Today.xlsx')

first_duplicate_index = df.index[df.duplicated(subset='Name', keep='first')][0]

df1 = df.iloc[:first_duplicate_index]
df2 = df.iloc[first_duplicate_index:]

with pd.ExcelWriter('Title Changes For Today.xlsx') as writer:
    df1.to_excel(writer, sheet_name='Today', index=False)
    df2.to_excel(writer, sheet_name='Last Week', index=False)


outlook = win32com.client.Dispatch('Outlook.Application')

namespace = outlook.GetNamespace('MAPI')

mail = outlook.CreateItem(0)

mail.Subject = 'Test'
mail.Body = 'Here are this weeks title changes, ' \
            'please review ' \
            'and adjust access accordingly '

#Emails removed for privacy
mail.To = 'example@email.com'

attachment = r'C:\Users\pehlinger\PycharmProjects\pythonProject1\Title Changes For Today.xlsx'
mail.Attachments.Add(attachment)

mail.Send()
