##import imaplib
## 
##IMAP_SERVER='imap.gmail.com'
##IMAP_PORT=993
## 
##M = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
##rc, resp = M.login('testxstar@gmail.com', 'vdblnsgdjllamdxb')
##print rc, resp
## 
##M.logout()
import imaplib, email

#log in and select the inbox
mail = imaplib.IMAP4_SSL('imap.gmail.com')
mail.login('testxstar@gmail.com', 'vdblnsgdjllamdxb')
mail.select('inbox')

#get uids of all messages
result, data = mail.uid('search', None, 'ALL') 
uids = data[0].split()

#read the lastest message
result, data = mail.uid('fetch', uids[-1], '(RFC822)')
m = email.message_from_string(data[0][1])

if m.get_content_maintype() == 'multipart': #multipart messages only
    for part in m.walk():
        #find the attachment part
        if part.get_content_maintype() == 'multipart': continue
        if part.get('Content-Disposition') is None: continue

        #save the attachment in the program directory
        filename = part.get_filename()
        fp = open(filename, 'wb')
        fp.write(part.get_payload(decode=True))
        fp.close()
        print '%s saved!' % filename
