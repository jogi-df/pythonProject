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
#f = glob.glob(r"C:\test\emails/*.msg")
df2 = pd. DataFrame()

# declaring variables
platform = ""
emailid = ""

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
def short_url(long_url):    #removes tracking in url
    if "mkt_tok" in long_url:
        b = re.split('(\&mkt_tok)|(\?mkt_tok)', long_url)
        print(b[0])
    elif "?context.guid" in long_url:
        b = re.split('(\?context.guid)', long_url)
        print(b[0])

def platform(url):
    if "marketo" in url:
        p = "Marketo"
        return p
    elif any(re.findall(r'campaign|m3-page', url)):
        p = "Campaign"
        return p


###for filename in f:
###    msg = extract_msg.Message(filename)
###    msg_message = msg.body
###    df1 = pd.DataFrame([x.split(';') for x in msg_message.split('\n')])
###    df1 = df1.rename(columns={df1.columns[0]: 'Email'})
###    df1 = df1.replace(r"^ |\t*$|\s*$", r"", regex=True)
#    msg_sender = msg.sender
#    msg_date = msg.date
#    msg_subj = msg.subject
#    msg_message = msg.htmlBody
#    print('Sender: {}'.format(msg_sender))
#    df1 = pd.DataFrame([msg_message.split('\n')])
#    df = pd.DataFrame([msg_message.split('\n')])
#    df1 = df1.replace(r" <", r"xx<", regex=True)
#     df2['Email'] = df1['Email'].str.split('xx')

i = 0
# this part extracts all the URLs in long form, unwraps from outlook, follows the urls
# and prints the response and short form as well as tracking id if exists
for filename in f:
    msg = extract_msg.Message(filename)
    msg_message = msg.htmlBody
    msg_subj = msg.subject
    msg_body = msg.body
#    print(msg_message)
    df1 = pd.DataFrame([x.split(';') for x in msg_body.split('\n')])
    df1 = df1.rename(columns={df1.columns[0]: 'Email'})
    df1 = df1.replace(r"^ |\t*$|\s*$", r"", regex=True)
    df1 = df1.replace(r'<[^<]*?/?>', r'', regex=True)
    df1 = df1.replace(r' / [\u200B -\u200D\uFEFF] /', '', regex=True)
    soup = BeautifulSoup(msg_message, "lxml")
    x = soup.find_all('a')
    text = soup.find_all(text=True)
#    legal = soup.find_all("td", class_="legal") # testing printing out sections by class this is the footer
#    print(x)  # prints list of a anchors
    print(msg_subj) # prints subject line
#    print(msg_body) # prints entire body
#    print(legal)
    for link in soup.find_all('a'):
        url1 = link.get('href')
#        print(url1) #wrapped url
        if "http" in url1:
#            print("url1", url1) #this should be the original url wrapped by email client or spam filter
            # url2 = follow_url(link.get('href'))
            driver.get(url1)
            url2 = driver.current_url
            print("unwrap", url2)  # possibly a marketo link or go url
            if "trackingid" in url2:
                t_id1 = tracking_id(url2)
                print(t_id1)
            if "rtid=" in url2:
                cid1 = cid(url2)
                print(cid1)
            s_url1 = short_url(url2)
            if any(re.findall(r'marketo|mkt_tok|mkto', url2)):
                platform = "Marketo"
            elif any(re.findall(r'campaign|m3-page', url2)):
                platform = "Campaign"
    print("Email platform is:", platform)
    i = i + 1
driver.close()

#    df1.set_index(['Email']).apply(lambda x: x.str.split('xx').explode()).reset_index()

#    df1.iloc[:,0] = df1.iloc[:,0].apply(lambda x: x.split('xx'))
#    a = df1.explode


#time.sleep(5)
#screenshot = driver.save_screenshot(r'C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check\my_screenshot.png')
#driver.quit()


df1.to_excel(r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check\working_copy.xlsx", index=False)
