import win32com.client
import os

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
# 收件匣
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
# messages = outlook.Folders.Item("Inbox").Items
# 倒序排序
messages = [m for m in messages][::-1]

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
