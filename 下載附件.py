import win32com.client
import os

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# 收件匣
inbox = outlook.GetDefaultFolder(6) # 預設資料夾
# inbox = outlook.Folders.Item(1).Folders['收件匣'] # 備份資料夾

messages = inbox.Items
# messages = outlook.Folders.Item("Inbox").Items
# messages = inbox.Items.Restrict("@SQL=""http://schemas.microsoft.com/mapi/proptag/0x0E1D001F"" like '%abc%' ") # 搜尋關鍵字

# 倒序排序
# messages = [m for m in messages][::-1]
messages.Sort("[ReceivedTime]", True)

for m in messages:
    # 搜尋特定標題
    if 'xxx' in m.Subject:
        # mail內文
        # body = m.Body
        
        # mail日期
        # date = m.senton.date()
        
        # 取得附件
        attachments = m.Attachments
        attachment = attachments.Item(1)
        attachment.SaveASFile(os.path.join(os.getcwd(), attachment.FileName))
        break
