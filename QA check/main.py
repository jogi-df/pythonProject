import extract_msg
import requests
import glob
import pandas as pd

def follow_url(url_wrapped):
    response = requests.get(url_wrapped)
    return response.url


f = glob.glob(r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check/*.msg")


for filename in f:
    msg = extract_msg.Message(filename)
#    msg_sender = msg.sender
#    msg_date = msg.date
#    msg_subj = msg.subject
    msg_message = msg.body
#    msg_html = msg.htmlBody
#    print('Sender: {}'.format(msg_sender))
    df1 = pd.DataFrame([x.split(';') for x in msg_message.split('\n')])
#    df = pd.DataFrame([msg_message.split('\n')])
    df1 = df1.replace(r"^ |\t+$|\s+$", r"", regex=True)
    df1 = df1.replace(r" <", r"xx<", regex=True)
#    df1.iloc[:,0] = df1.iloc[:,0].apply(lambda x: x.split('xx'))
#    a = df1.explode

    #print('Sent On: {}'.format(msg_date))
#    print('Subject: {}'.format(msg_subj))
    #print('Body: {}'.format(msg_message))
    #print(msg_subj)
 #   print('HTML: {}'.format(msg_html))


print(df1)

earl = follow_url("https://nam04.safelinks.protection.outlook.com/?url=https%3A%2F%2Ft-trg.email.adobe.com%2Fr%2F%3Fid%3Dh8ac57dfb%2C8eff47ee%2C84e306a9%26e%3DcDE9Y3JlYXRpdmVjbG91ZC5hZG9iZS5jb20vY2MvZGlzY292ZXIvYXJ0aWNsZS90aGUtdW5sb2NrLWthcmVuLXgtY2hlbmctb24tZ29pbmctdmlyYWwtYW5kLW9wdGltaXppbmctZm9yLWZ1bj9jb250ZXh0Lmd1aWQ9MzUxNzlkMGEtYTkyNC00Y2QzLWE3ZDItNDZhYjBiZDE4NmJjJmNvbnRleHQuaW5pdD1mYWxzZSZwMj04MUc1NVZTNw%26s%3DXb8a8FowjkD7NDoGQ7DcMZNp7r63D0A62b8qDtVHuaY&data=05%7C01%7Cdat83604%40adobe.com%7C82ee8903d3ad448a0eff08daa7103bee%7Cfa7b1b5a7b34438794aed2c178decee1%7C0%7C0%7C638005985407860464%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C3000%7C%7C%7C&sdata=y%2FGj0%2BEYpuJ1FcUIJW4lsWMlph3R%2FHqKCD2mJ46AvFM%3D&reserved=0")


print(earl)

df1.to_excel(r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check\working_copy.xlsx", index=False)