# Python-Automated-Billing-Mail
Simple application for automaticaly sending a mail via Outlook to a list of people(xlsx file)  contacts which present certain criteria. 

The application imports three packages. They will do most of the job.
win32com.client will allow us to acess Outlook, pandas will allow us to read the excel file, and datetime will be our time reference.
As we want to send billing mails to clients who are overdue with their invoces, we need to compare "today" with the date present on the xlsx file.

If the the critirea is met, the aplication sends a string formatted text which contains all of the necessary info about the overdue invoce. 
