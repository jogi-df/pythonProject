import requests
import mechanize
import logging
import sys
import http.cookiejar
import mechanicalsoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

#logger = logging.getLogger("mechanize")
#logger.addHandler(logging.StreamHandler(sys.stdout))
#logger.setLevel(logging.INFO)

# br = mechanize.Browser()
#br = mechanicalsoup.StatefulBrowser()

# Cookie Jar
#cj = http.cookiejar.LWPCookieJar()
#br.set_cookiejar(cj)

# Browser options
#br.set_handle_equiv(True)
#br.set_handle_gzip(True)
#br.set_handle_redirect(True)
#br.set_handle_referer(True)
#br.set_handle_robots(False)

# User-Agent (this is cheating, ok?)
#br.addheaders = [('User-agent', 'Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.1) Gecko/2008071615 Fedora/3.0.1-1.fc9 Firefox/3.0.1')]

#s = requests.Session()
#s.get('https://resources.marketo.com/dc/R0TZFG3XghX3T46Uif6Cy_mSGVz-v5v6N4139DVISLFaXWIK4Q63h2NVxQY5H5yu3vE0fDpzDVXGevBy5cRLvw==/NDYwLVRESC05NDUAAAGIyk9A2xmA5X9LsKAH5eR9JWuCl397P6sSdQN2cgNd6-MMQP6agxKPHdbNOfltqXlRYDoN7n8=')
#r = s.get('https://resources.marketo.com/dc/R0TZFG3XghX3T46Uif6Cy_mSGVz-v5v6N4139DVISLFaXWIK4Q63h2NVxQY5H5yu3vE0fDpzDVXGevBy5cRLvw==/NDYwLVRESC05NDUAAAGIyk9A2xmA5X9LsKAH5eR9JWuCl397P6sSdQN2cgNd6-MMQP6agxKPHdbNOfltqXlRYDoN7n8=')
#print(r.text)



# r=br.open("https://resources.marketo.com/dc/R0TZFG3XghX3T46Uif6Cy_mSGVz-v5v6N4139DVISLFaXWIK4Q63h2NVxQY5H5yu3vE0fDpzDVXGevBy5cRLvw==/NDYwLVRESC05NDUAAAGIyk9A2xmA5X9LsKAH5eR9JWuCl397P6sSdQN2cgNd6-MMQP6agxKPHdbNOfltqXlRYDoN7n8=")
#print (x.read())

url = "https://nam04.safelinks.protection.outlook.com/?url=http%3A%2F%2Fsource.adobe.com%2FMzYwLUtDSS04MDQAAAGIHIXfpy-b4GpxhgpWMGFZF1K93Rekvnks3KIW-Ei04ensPu5iuG5wAmyN-9Wm87pyKcaZ5ks%3D&amp;data=05%7C01%7Cdat03013%40adobe.com%7Cd6633b69c2a44f80ff2808dac75e5353%7Cfa7b1b5a7b34438794aed2c178decee1%7C0%7C0%7C638041505181415941%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C3000%7C%7C%7C&amp;sdata=3O7Ga78C3uLh3UIUN8jTrt8DPLH6HR1tiOLnIOQ7arQ%3D&amp;reserved=0"
#url = "https://resources.marketo.com/dc/R0TZFG3XghX3T46Uif6Cy_mSGVz-v5v6N4139DVISLFaXWIK4Q63h2NVxQY5H5yu3vE0fDpzDVXGevBy5cRLvw==/NDYwLVRESC05NDUAAAGIyk9A2xmA5X9LsKAH5eR9JWuCl397P6sSdQN2cgNd6-MMQP6agxKPHdbNOfltqXlRYDoN7n8="
#url = "https://apps.enterprise.adobe.com/go/7015Y000004BYeUQAW?mkt_tok=NDYwLVRESC05NDUAAAGIyk9A28703gSAwNTGN7iL4hzG_3OQLw12k4Jat5V-oun2rCoYFOCPY-IYqEcM8CawPm9pDvNm6fXuRS98b13j4Hr445LheVH8FJ47YUsJUFuUrIa8"
#url = "https://nam04.safelinks.protection.outlook.com/?url=http%3A%2F%2Fsource.adobe.com%2FMzYwLUtDSS04MDQAAAGIHIXfpwwYEKmMmTFUgiRv3xhpVHjGe-sqZf3tPwNkio0-NMdbZJ-kftYDn_Kz-BCHsf-CYsA%3D&data=05%7C01%7Cdat03013%40adobe.com%7Cd6633b69c2a44f80ff2808dac75e5353%7Cfa7b1b5a7b34438794aed2c178decee1%7C0%7C0%7C638041505181415941%7CUnknown%7CTWFpbGZsb3d8eyJWIjoiMC4wLjAwMDAiLCJQIjoiV2luMzIiLCJBTiI6Ik1haWwiLCJXVCI6Mn0%3D%7C3000%7C%7C%7C&sdata=pIsvkS%2FipXqHXYtMpDwoWUpZNX6AFTex9qQICfGByrw%3D&reserved=0"


WINDOW_SIZE = "1920,1080"
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--window-size=%s" % WINDOW_SIZE)


driver = webdriver.Chrome(options=chrome_options)
driver.get(url)
u2 = driver.current_url
driver.get(u2)
u3 = driver.current_url
print(u3)
driver.close()