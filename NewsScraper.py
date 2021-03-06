# ============================================================================
# USA companies news web scraping
# Author - P. Dyachenko
# =============================================================================
from MailSender import sendme_dataframe
import requests
from bs4 import BeautifulSoup
import re
import numpy as np
import pandas as pd
import sys, traceback

import schedule
import time

def job():
    print("I'm working...")

    news_links = {}
    urls_to_scrape = ["https://www.marketwatch.com/", "https://www.reuters.com/"]
    keywords_to_parse = [".*Carnival", "Cruise", "Boeing", "Delta", "trump", "Covid", "Oil", "Opek", "Airline.*",]
    pattern_to_match = ".* | .*".join(keywords_to_parse)
    print("Looking for " + pattern_to_match)

    for url in urls_to_scrape:
        # Http request getting soup object
        request_object = requests.get(url)
        coverpage = request_object.content
        soup_object = BeautifulSoup(coverpage, "html.parser")
        # Scanning the site filling the news dict
        page_articles = soup_object.find_all('a', href=True, )
        for art in page_articles:
            try:
                news_text = art.get_text().strip()
                if re.match(pattern_to_match, news_text, flags=re.I | re.X):
                    print("Adding news: " + news_text + " to the send list")
                    # print(art.get("href")[0:3])
                    url_prefix = url[0:-1] if art.get("href")[0:4] != "http" else ""
                    news_links[url_prefix + art.get("href")] = news_text
            except Exception as e:
                print(e)
                print('-' * 60)
                traceback.print_exc(file=sys.stdout)
                print('-' * 60)
                continue

    news_table = pd.DataFrame(list(news_links.items()), columns=["URL","Text"])
    sendme_dataframe(news_table)

# schedule.every(10).minutes.do(job)
schedule.every().hour.do(job)
# schedule.every().day.at("10:30").do(job)

while 1:
    schedule.run_pending()
    time.sleep(1)



