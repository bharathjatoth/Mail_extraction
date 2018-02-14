import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
print(len(messages))
message = messages.GetLast()
body_content = message.body
print(body_content)
list1 = []
#all the fields are self explanatory 
for i in range(len(messages)):
    print(messages[i].CreationTime)
    print(messages[i].Subject)
    print(messages[i].body)
    print(messages[i].attachments)
    print(messages[i].SenderName)
    print(messages[i].Sender)
    print(messages[i].ReceivedTime)
