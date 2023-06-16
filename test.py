import win32com.client

ol = win32com.client.Dispatch("outlook.application")

olmailitem = 0x0  # size of the new email
newmail = ol.CreateItem(olmailitem)

newmail.Subject = 'Testing Mail'
newmail.To = 'xyz@gmail.com'
newmail.CC = 'xyz@gmail.com'

newmail.Body = 'Hello, this is a test email to showcase how to send emails from Python and Outlook.'
# attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
# newmail.Attachments.Add(attach)

# Display the mail before sending it
newmail.Display()
# newmail.Send()
