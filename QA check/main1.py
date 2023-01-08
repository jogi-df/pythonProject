import extract_msg
import requests
import glob
import pandas as pd
import re
import time
import openpyxl
from openpyxl import load_workbook
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from fake_useragent import UserAgent

WINDOW_SIZE = "1024,768"
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=%s" % WINDOW_SIZE)
chrome_options.add_argument("user-agent={userAgent}")
driver = webdriver.Chrome(options=chrome_options)

save_to_folder = r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check"

writer = pd.ExcelWriter(r"C:\Users\jogi\OneDrive - binary-tech.com\Consulting\DF 2022\QA check\working_copy.xlsx", engine='xlsxwriter', engine_kwargs={'options':{'strings_to_formulas': False}})

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

def activityid(unwrapped_url):
    if 'ActivityID=' in unwrapped_url:
        aid2 = re.search("ActivityID=(.{7})", unwrapped_url)
        return aid2.group(1)

def short_url(long_url):    #removes tracking in url
    if "mkt_tok" in long_url:
        b = re.split('(\&mkt_tok)|(\?mkt_tok)', long_url)
        data.append(b[0])
#        print(b[0])
    elif "?context.guid" in long_url:
        b = re.split('(\?context.guid)', long_url)
        data.append(b[0])
#        print(b[0])

def platform(url): # figures out if email is sent from marketo or AC.  could be useful in the future
    if "marketo" in url:
        p = "Marketo"
        return p
    elif any(re.findall(r'campaign|m3-page', url)):
        p = "Campaign"
        return p

def sender(email_sender): # extracts display name and sender email from the sender field
   email_1 = re.search("^([\w\s]+)\s\<(.*?)\>", email_sender)
   dn = email_1.group(1)
   em = email_1.group(2)
   return dn, em


i=0


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
        aid3 = re.search("^\[(.{7})", msg_subj)
        aid3 = aid3.group(1)
        msg_subj = re.sub("^\[.*?\] ", "", msg_subj)
    data.append("SL:" + msg_subj)
    dn, em = sender(msg.sender)
    data.append("Display name: " + dn)
    data.append("Sender email: " + em)
#    df1.loc[df1.index[0], 'Email'] = msg_subj  # adds in SL in xl file - using this overwrites PH because of indexing that i haven't fixed
    print("Processing: " + msg_subj) # prints email SL so it looks like it's doing something
#    msg_subj_title = msg_subj[0:30] # using first 30 characters of SL for ws title
#    print(msg_subj) # prints subject line
#    print(msg_body) # prints entire body
#    print(legal)
    j = 0
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
            else:
                data.append("no tracking ID")
                # print("no tracking ID")
            if "rtid=" in url2:
                cid1 = cid(url2)
                data.append(cid1)
                # print(cid1)
            else:
                data.append("no CID")
            if "ActivityID=" in url2:
                aid1 = activityid(url2)
                data.append(aid1)
                # print(cid1)
            else:
                data.append("no ActivityID")
                # print("no ActivityID")
            s_url1 = short_url(url2)
            if any(re.findall(r'marketo|mkt_tok|mkto', url2)):
                platform = "Marketo"
            elif any(re.findall(r'campaign|m3-page|m2-page|m1-page', url2)):
                platform = "Campaign"

            # this section creates the screenshots from the URLs
            driver.get(url2)
            ss_file = "ss" + str(j) + ".png"
            time.sleep(5)
            driver.save_screenshot(save_to_folder + "\\" + ss_file)
            j = j + 1


    # print("Email platform is:", platform)
    df2 = pd.DataFrame({'Email':data})
    #df1 = df1.append(df2)
    frames = [df1, df2]
    df1 = pd.concat(frames) # using concat instead of append due to deprecation of append
#    ws=wb.create_sheet(msg_subj)
#    print(df1)
#    ws_title = "Email " + str(i)
    ws_title = aid3
    df1.to_excel(writer, sheet_name=ws_title, index=False)
    #writer.save()
    df1 = df1[0:0]
    data = []
    i = i + 1
driver.close()
writer.close()
writer.quit()
driver.quit()

