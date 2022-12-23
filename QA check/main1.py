import extract_msg
import requests
import glob
import pandas as pd
import re
import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

WINDOW_SIZE = "1920,1080"
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=%s" % WINDOW_SIZE)
driver = webdriver.Chrome(options=chrome_options)

f = glob.glob(r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check\emails/*.msg")

df2 = pd. DataFrame()


def follow_url(url_wrapped):
    response = requests.get(url_wrapped)
    return response.url


def tracking_id(unwrapped_url):
    if 'trackingid=' in unwrapped_url:
        trackid = re.search("trackingid=(.{8})", unwrapped_url)
        return trackid.group(1)

def cid(unwrapped_url):
    if 'rtid=' in unwrapped_url:
        cid2 = re.search("rtid=(.{18})", unwrapped_url)
        return cid2.group(1)
def short_url(long_url):    #removes marketo tracking in url
    # a = [long_url]
    if "mkt_tok" in long_url:
        b = re.split('(\&mkt_tok)|(\?mkt_tok)', long_url)
        print(b[0])



###for filename in f:
###    msg = extract_msg.Message(filename)
#    msg_sender = msg.sender
#    msg_date = msg.date
#    msg_subj = msg.subject
###    msg_message = msg.body
#    msg_message = msg.htmlBody
#    print('Sender: {}'.format(msg_sender))
###    df1 = pd.DataFrame([x.split(';') for x in msg_message.split('\n')])
#    df1 = pd.DataFrame([msg_message.split('\n')])
###    df1 = df1.rename(columns={df1.columns[0]: 'Email'})
#    df = pd.DataFrame([msg_message.split('\n')])
###    df1 = df1.replace(r"^ |\t*$|\s*$", r"", regex=True)
#    df1 = df1.replace(r" <", r"xx<", regex=True)
#     df2['Email'] = df1['Email'].str.split('xx')

# this part extracts all the URLs in long form, unwraps from outlook, follows the urls
# and prints the response and short form as well as tracking id if exists
for filename in f:
    msg = extract_msg.Message(filename)
    msg_message = msg.htmlBody
    msg_subj = msg.subject
    msg_body = msg.body
    soup = BeautifulSoup(msg_message, "lxml")
    x = soup.find_all('a')
    legal = soup.find_all("td", class_="legal")
#    print(x)  # prints list of a anchors
    print(msg_subj) # prints subject line
#    print(msg_body) # prints entire body
    print(legal)
    for link in soup.find_all('a'):
        matches = ["vcf", "jsp", "/go/", "facebook", "instagram", "linkedin", "tiktok", "twitter", "youtube", "pinterest",
                   "Subscription", "trademarks", "privacy",
                   "unsubscribe", "emailWebview", "mktoTestLink", "policy"]  # list of words for links to stop following
        matches1 = ["vcf", "jsp", "/go/", "facebook", "instagram", "linkedin", "tiktok", "twitter", "youtube", "pinterest",
                    "Subscription", "trademarks", "privacy",
                    "unsubscribe", "emailWebview", "mktoTestLink", "policy",
                    "trackingid"]  # list of words for links to stop following
        url1 = link.get('href')
#        print(url1) #wrapped url
        if "http" in url1:
#            print("url1", url1) #this should be the original url wrapped by email client or spam filter
            # url2 = follow_url(link.get('href'))
            driver.get(url1)
            url2 = driver.current_url
            print("unwrap", url2)  # possibly a marketo link or go url
            s_url1 = short_url(url2)
            t_id1 = tracking_id(url2)
            cid1 = cid(url2)
            print(t_id1)
            print(cid1)
driver.close()

#    df1.set_index(['Email']).apply(lambda x: x.str.split('xx').explode()).reset_index()

#    df1.iloc[:,0] = df1.iloc[:,0].apply(lambda x: x.split('xx'))
#    a = df1.explode

    # print('Sent On: {}'.format(msg_date))
#    print('Subject: {}'.format(msg_subj))
    # print('Body: {}'.format(msg_message))
    # print(msg_subj)
#   print('HTML: {}'.format(msg_html))


# print(df1)
# print(df2)


# earl = follow_url("http://source.adobe.com/MzYwLUtDSS04MDQAAAGIHIXfp-nQM4dwwbAP2SWFBdipMxJvBHYUVxzie8XHw6A-eUrw4EiBp1vyrK1JLYP0GDGSG8Y=")
# earl = follow_url("https://nam04.safelinks.protection.outlook.com/?url=https%3A%2F%2Ft-trg.email.adobe.com%2Fr%2F%3Fid%3Dh8ac57dfb%2C8eff47ee%2C84e306a9%26e%3DcDE9Y3JlYXRpdmVjbG91ZC5hZG9iZS5jb20vY2MvZGlzY292ZXIvYXJ0aWNsZS90aGUtdW5sb2NrLWthcmVuLXgtY2hlbmctb24tZ29pbmctdmlyYWwtYW5kLW9wdGltaXppbmctZm9yLWZ1bj9jb250ZXh0Lmd1aWQ9MzUxNzlkMGEtYTkyNC00Y2QzLWE3ZDItNDZhYjBiZDE4NmJjJmNvbnRleHQuaW5pdD1mYWxzZSZwMj04MUc1NVZTNw%26s%3DXb8a8FowjkD7NDoGQ7DcMZNp7r63D0A62b8qDtVHuaY&data=05%7C01%7Cdat83604%40adobe.com%7C82ee8903d3ad448a0eff08daa7103bee%7Cfa7b1b5a7b34438794aed2c178decee1%7C0%7C0%7C638005985407860464%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C3000%7C%7C%7C&sdata=y%2FGj0%2BEYpuJ1FcUIJW4lsWMlph3R%2FHqKCD2mJ46AvFM%3D&reserved=0")
# earl2 = follow_url("https://nam04.safelinks.protection.outlook.com/?url=https%3A%2F%2Ft-trg.email.adobe.com%2Fr%2F%3Fid%3Dh8ac57dfb%2C8eff47ee%2C84e306ac%26e%3DcDE9Y3JlYXRpdmVjbG91ZC5hZG9iZS5jb20vY2MvZGlzY292ZXIvYXJ0aWNsZS9tYWtlLWFuLWFuaW1hdGVkLXBob3RvLWlsbHVzdHJhdGlvbi1pbi1qdXN0LWZpdmUtc3RlcHM_Y29udGV4dC5ndWlkPTM1MTc5ZDBhLWE5MjQtNGNkMy1hN2QyLTQ2YWIwYmQxODZiYyZjb250ZXh0LmluaXQ9ZmFsc2UmcDI9ODU2NjVSMjY%26s%3DZSQpLhYobEYqwlBXGdJ9SxzKjcb5iD4Nd-nRNUG3dvY&data=05%7C01%7Cdat83604%40adobe.com%7C82ee8903d3ad448a0eff08daa7103bee%7Cfa7b1b5a7b34438794aed2c178decee1%7C0%7C0%7C638005985408016680%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C3000%7C%7C%7C&sdata=OEVmf916q3I0GXEXcJHlwWvJ4Q8ZKRZNSonkFLrvyGw%3D&reserved=0")
# print(earl)
# s_u = short_url(earl)
# t_id = tracking_id(earl)
# print(t_id)


# print(earl2)
# s_u2 = short_url(earl2)
# t_id2 = tracking_id(earl2)
# print(t_id2)

#DRIVER = 'chromedriver'
#driver = webdriver.Chrome(DRIVER)
#driver.get(earl2)

#time.sleep(5)
#screenshot = driver.save_screenshot(r'C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check\my_screenshot.png')
#driver.quit()


###df1.to_excel(r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check\working_copy.xlsx", index=False)
