
#Importing the Packages

import win32com.client as client      #To acess Outlook
import pandas as pd                   #To read the XLSX file
import datetime as dt                 #to have today as a refence

#Reading the XLSX file
table = pd.read_excel('Billing.xlsx')
#Setting today as the reference
today = dt.datetime.now()

#Choosing only clients that  are within the "Overdue" filter
table_bills = table.loc[table['Status']=='Not payed']
table_bills = table_bills.loc[table_bills['Due date']<today]

#Setting the Outlook session
outlook = client.Dispatch('Outlook.Application')
emitter = outlook.session.Accounts['PLACEMAILHERE@gmail.com']

#Setting the format
data= table_bills[['Amount','Due date','E-mail','Invoice']].values.tolist()

for datum in data:
    receiver = data[2]
    invoice = data[3]
    due = data[1]
    due = due.strftime("%d/%m/%Y")
    amount = data[0]
    assunto = 'Overdue Payment'
    message = outlook.CreateItem(0)
    message.To = receiver
    message.Subject = receiver
    message_body = f'''
    
    Dear customer,

    We have noticed that your invoice: {invoice} due to the date {due} , in the amout of  R${amount:.2f} is overdue.
    We would like to know if there are any issues that you might have come by and if you currently are in need of our 
    aid. If any doubts or questions persist, please contact us by this very e-mail"
    
    Best regards,
    Douglas Oliveira
    '''
    message.Body =  message_body
    message._oleobj_.Invoke(*(64209, 0, 8, 0, emitter))
    message.Save()
    message.Send()
