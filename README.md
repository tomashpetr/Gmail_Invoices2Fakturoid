**Script using Google Sheet and Google Appscript to upload invoices to Fakturoid.cz accounting service**

Initial setup:
1. in Google drive create empty google sheet file
2. create two lists - A) one called "FakturoidLog" and B) second called "EmlsRecords"
3. open appscript editor (from the opened google sheet menu)
4. copy the code from this repo to the appscript editor
5. adjust the config values to match your fakturoid account (clientID, ClientSecret, company account slug, etc.)
6. run the function "processBoltInvoices"
7. approve all access rigths needed (using your own google account)

The function will:
1. reach to fakturoid API to get the security http token
2. process your emails in your inbox (for the logged in google account gmail)
3. if the emails match the query, then search for linked PDF file
4. download the PDF file and upload it to the fakturoid inbox (for further processing)
5. mark the message ID of successfully uploaded PDF into second gsheet list - to have a record of emails which should be skipped if function is launched in the future again
6. the script generates multiple logging records in the first gsheet list

In fakturoid it is possible to automatically extract the data from PDF for easier accounting
see: https://www.fakturoid.cz/
