import pandas as pd
import win32com.client as win32
import re
import os
from dotenv import load_dotenv
from config import *

load_dotenv()

def processTemplate(templateFilePath, filesPath, encodings, replacers):
    baseHTML = None

    for encoding in encodings:
        try:
            with open(templateFilePath, 'r', encoding=encoding) as file:
                baseHTML = file.read()
            print(f"Successfully read the template using {encoding} encoding.")
            break
        except UnicodeDecodeError:
            print(f"Failed to read with {encoding} encoding, trying next...")

    if baseHTML is None:
        print("Failed to read the template file with any of the attempted encodings.")
        exit()
    
    tmp1 = baseHTML.split('<div class=WordSection1>',1)[0]
    tmp2 = baseHTML.split('<p class=MsoNormal><o:p>&nbsp;</o:p></p>',1)[1]

    baseHTML = f"{tmp1}<div class=WordSection1>{tmp2}"

    for pattern in replacers:
        baseHTML = baseHTML.replace(pattern, replacers[pattern])

    imageAttachments = []
    matches = re.findall(fr'src="{filesPath}/(.*?)"', baseHTML)
    id = 0
    cid = f"cid:image{id}"
    for i, img_filename in enumerate(matches):
        img_path = os.path.join(filesPath, img_filename)
        if os.path.exists(img_path):
            baseHTML = baseHTML.replace(f'src="{filesPath}/{img_filename}"', f'src="{cid}"')
            if img_path not in imageAttachments:
                imageAttachments.append(img_path)
                id+=1
            cid = f"cid:image{id}"

    return baseHTML, imageAttachments

def createMail(outlook, account, baseHTML, imageAttachments, senderName, mailSubject, replacers):
    mail = outlook.CreateItem(0)
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))
    mail.SentOnBehalfOfName = senderName
    mailHTML = baseHTML

    for pattern in replacers:
        if replacers[pattern] in row:
            mailHTML = mailHTML.replace(pattern, row[replacers[pattern]])

    for i in range(0, len(imageAttachments)):
        attachment = mail.Attachments.Add(os.path.abspath(imageAttachments[i]))
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", f"image{i}")
        attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", True)

    mail.Subject = mailSubject
    mail.HTMLBody = mailHTML
    mail.To = email
    if action == "draft":
        mail.Save()
        print(f"Saved mail to drafts folder")
    elif action == "send":
        mail.Send()
        print(f"Sent mail")

if __name__ == "__main__":
    outlook = win32.Dispatch('outlook.application')
    namespace = outlook.GetNamespace("MAPI")
    account = namespace.Accounts.Item(os.getenv('OUTLOOK_ACCOUNT_INDEX'))
    df = pd.read_csv(os.getenv('CSV_FILE'))
    templateFilePath = os.getenv('TEMPLATE_FILE')
    filesPath = os.getenv('FILES_PATH')
    if filesPath is None:
        filesPath = f"{templateFilePath.split('.')[0]}_files"
    nrRows = len(df)

    print("== Outlook Mass Mail ==")
    print(f"Using {nrRows} contacts from {os.getenv('CSV_FILE')}")
    print(f"Using {templateFilePath} as base message")
    if action == "draft":
        print("Mails will be saved to the drafts folder")
    elif action == "send":
        print("Mails will be sent immediately")
    else:
        print("Invalid action {action}, aborting")
        exit()
    print(f"Sending mail as {os.getenv('SENDER_NAME')}")
    print(f"Using subject: '{os.getenv('MAIL_SUBJECT')}'")

    if waitForInput:
        input("Press enter to continue")

    baseHTML, imageAttachments = processTemplate(templateFilePath, filesPath, encodings, replacers['consistent'])

    for index, row in df.iterrows():
        name = row[os.getenv('CSV_FILE_NAME_COLUMN_LABEL')]
        email = row[os.getenv('CSV_FILE_EMAIL_COLUMN_LABEL')]
        if pd.isnull(email):
            print(f"Skipping {name}: No Email set")
            continue
        print(f"{index+1}/{nrRows} Creating mail for {name}")
        createMail(outlook, account, baseHTML, imageAttachments, os.getenv('SENDER_NAME'), os.getenv('MAIL_SUBJECT'), replacers['csv'])
        