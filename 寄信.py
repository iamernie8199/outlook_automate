import win32com.client

ol = win32com.client.Dispatch("outlook.application")

olmailitem = 0x0  # size of the new email
newmail = ol.CreateItem(olmailitem)

newmail.Subject = 'Testing Mail'
newmail.To = 'xyz@gmail.com'
newmail.CC = 'xyz@gmail.com; abc@gmail.com'

# newmail.Body = 'Hello, this is a test email to showcase how to send emails from Python and Outlook.'

newmail.Attachments.Add(os.path.join(os.getcwd(), '123.png')).PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "123")
body = "<b>blablabla</b><br><br>"
body += """<img src="cid:123"  width="800"><br>"""
newmail.HTMLBody = body

# attach='C:\\Users\\admin\\Desktop\\Python\\Sample.xlsx'
# newmail.Attachments.Add(attach)

# Display the mail before sending it
newmail.Display()
# newmail.Send()
