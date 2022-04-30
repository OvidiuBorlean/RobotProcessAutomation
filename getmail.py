import imaplib
import base64
import os
import email
import random

email_user = 'magazzino@test.biz'
email_pass = 'mypass'

mail = imaplib.IMAP4_SSL('mail.blogistica.biz',993)
mail.login(email_user, email_pass)
mail.select('Inbox')
type, data = mail.search(None, 'UNSEEN')
mail_ids = data[0]
id_list = mail_ids.split()


for num in data[0].split():
    typ, data = mail.fetch(num, '(RFC822)' )
    raw_email = data[0][1]
# converts byte literal to string removing b''
    raw_email_string = raw_email.decode('utf-8')
    email_message = email.message_from_string(raw_email_string)
# downloading attachments
    for part in email_message.walk():
        # this part comes from the snipped I don't understand yet... 
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        fileName = part.get_filename()
        if bool(fileName):
            randomNr = random.randint(0,9)
            filePath = os.path.join('./', str(randomNr) + fileName)
            if not os.path.isfile(filePath) :
                
                fp = open(filePath, 'wb')
                print(filePath)
                
                fp.write(part.get_payload(decode=True))
                fp.close()
            subject = str(email_message).split("Subject: ", 1)[1].split("\nTo:", 1)[0]
            fromsplit = str(email_message).split("From: ", 1)[1].split("\nDate:", 1)[0]
            print(fromsplit)
            #print('Downloaded "{file}" from email titled "{subject}" with UID {uid}.'.format(file=fileName, subject=subject, uid=latest_email_uid.decode('utf-8')))


for response_part in data:
        if isinstance(response_part, tuple):
            msg = email.message_from_string(response_part[1].decode('utf-8'))
            email_subject = msg['subject']
            email_from = msg['from']
            #print ('From : ' + email_from + '\n')   
            #print ('Subject : ' + email_subject + '\n')
            #print(msg.get_payload(decode=True))












