__author__ = 'lepik'
import imaplib
import email

# gmail account
m_user = 'testxstar@gmail.com'  # gmail address
m_pass = 'vdblnsgdjllamdxb'  # application secret key, allowed in your gmail account


def extract_body(payload):
    if isinstance(payload, str):
        return payload
    else:
        return '\n'.join([extract_body(part.get_payload()) for part in payload])

conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
conn.login(m_user, m_pass)
conn.select()
typ, data = conn.search(None, 'UNSEEN')
print('Unseen messages: ', len(data[0].split()))
if len(data[0].split()) == 0:
    print('Nothing to download.')
else:
    print('Downloading:')
try:
    for num in data[0].split():
        typ, msg_data = conn.fetch(num, '(RFC822)')
        #print(msg_data)
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                #print(response_part[1])
                msg = email.message_from_bytes(response_part[1])
                subject = msg['subject']
                #print(subject)
                payload = msg.get_payload()
                body = extract_body(payload)
                #print(body)
                if msg.get_content_maintype() == 'multipart':  # multipart messages only
                    for part in msg.walk():
                        #find the attachment part
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get('Content-Disposition') is None:
                            continue

                        #save the attachment in the program directory
                        filename = part.get_filename()
                        fp = open(r"./XLS_Received_from_elders/" + filename, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
                        print('%s saved!' % filename)

                typ, response = conn.store(num, '+FLAGS', r'(\Seen)')
finally:
    try:
        conn.close()
    except:
        pass
    conn.logout()