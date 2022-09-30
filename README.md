# Python-Automated-Billing-Mail
 Application that automatically sends  e-mails via Outlook to a list of contacts(xlsx file) which present certain criteria. 

The application imports three packages. They will do most of the job.
win32com.client will allow us to acess Outlook, pandas will allow us to read the excel file, and datetime will be our time reference.
As we want to send billing mails to clients who are overdue with their invoices, we need to compare "today" with the dates present on the xlsx file.

If the the critirea is met, the aplication sends a string formatted text which contains all of the necessary info about the overdue invoce. 
