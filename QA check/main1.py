import extract_msg
import requests
import glob
import pandas as pd
import re
import openpyxl
from openpyxl import load_workbook
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

WINDOW_SIZE = "1920,1080"
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=%s" % WINDOW_SIZE)
driver = webdriver.Chrome(options=chrome_options)

#file_name = r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check\working_copy.xlsx"
writer = pd.ExcelWriter(r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check\working_copy.xlsx", engine='openpyxl', mode='a')

f = glob.glob(r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check\emails/*.msg")
#f = glob.glob(r"C:\test\emails/*.msg")
df2 = pd. DataFrame()
data = []

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
        data.append(b[0])
#        print(b[0])
    elif "?context.guid" in long_url:
        b = re.split('(\?context.guid)', long_url)
        data.append(b[0])
#        print(b[0])

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

#wb=load_workbook(r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check\working_copy.xlsx") #open workbook to add sheets to--need to fix this
i=0
#book = openpyxl.load_workbook(r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check\working_copy.xlsx", 'rb')
#writer.book = book

# this part extracts all the URLs in long form, unwraps from outlook, follows the urls
# and prints the response and short form as well as tracking id if exists
for filename in f:
    msg = extract_msg.Message(filename)
    msg_message = msg.htmlBody
    msg_subj = msg.subject
#    msg_subj_title = msg_subj[0:30] # using first 30 characters of SL for ws title
    msg_body = msg.body
    soup = BeautifulSoup(msg_message, "lxml")
    text = soup.find_all(text=True)
    df1 = pd.DataFrame([x.split(';') for x in msg_body.split('\n')])
    df1 = df1.rename(columns={df1.columns[0]: 'Email'})
    df1 = df1.replace(r"^ |\t*$|\s*$", r"", regex=True)
    df1 = df1.replace(r'<[^<]*?/?>', r'', regex=True)
    df1 = df1.replace(r' ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌  ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌  ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌  ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌  ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌  ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌  ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌  ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌', '', regex=True)
    df1 = df1.replace(r' ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌ ‌', '', regex=True)
    soup = BeautifulSoup(msg_message, "lxml")
    x = soup.find_all('a')
#    legal = soup.find_all("td", class_="legal") # testing printing out sections by class this is the footer
#    print(x)  # prints list of a anchors
    if any(re.findall(r'TEST \|', msg_subj)): # removes TEST from SL - marketo
        msg_subj = re.sub("TEST \| ", "", msg_subj)
    if any(re.findall(r'^\[.*?\]', msg_subj)): # removes bracketed test info from SL - campaign
        msg_subj = re.sub("^\[.*?\] ", "", msg_subj)
    data.append(msg_subj)
    msg_subj_title = msg_subj[0:30] # using first 30 characters of SL for ws title
#    print(msg_subj) # prints subject line
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
            data.append("unwrap " + url2)
            # print("unwrap", url2)  # possibly a marketo link or go url
            if "trackingid" in url2:
                t_id1 = tracking_id(url2)
                data.append(t_id1)
                # print(t_id1)
            if "rtid=" in url2:
                cid1 = cid(url2)
                data.append(cid1)
                # print(cid1)
            s_url1 = short_url(url2)
            if any(re.findall(r'marketo|mkt_tok|mkto', url2)):
                platform = "Marketo"
            elif any(re.findall(r'campaign|m3-page', url2)):
                platform = "Campaign"
    print("Email platform is:", platform)
    df2 = pd.DataFrame({'Email':data})
    df1 = df1.append(df2)
#    ws=wb.create_sheet(msg_subj)
#    print(df1)
    df1.to_excel(writer, sheet_name=msg_subj_title, index=False)
    writer.save()
    df1 = df1[0:0]
    data = []
driver.close()
writer.close()

#time.sleep(5)
#screenshot = driver.save_screenshot(r'C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check\my_screenshot.png')
#driver.quit()


# df1.to_excel(r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check\working_copy.xlsx", sheet_name=cid1, index=False)