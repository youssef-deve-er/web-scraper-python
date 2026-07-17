import sys
import io
# إجبار بايثون على استخدام ترميز UTF-8 لدعم اللغة العربية في الـ Terminal والـ Output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import requests
from bs4 import BeautifulSoup
import pandas as pd

def scrape_tech_news():
    url = "https://news.ycombinator.com/" 
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    
    print("جاري الاتصال بالموقع واستخراج البيانات...")
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, 'html.parser')
        articles = soup.find_all('span', class_='titleline')
        
        news_list = []
        for article in articles:
            a_tag = article.find('a')
            if a_tag:
                title = a_tag.text
                link = a_tag['href']
                news_list.append({"العنوان (Title)": title, "الرابط (Link)": link})
        
        df = pd.DataFrame(news_list)
        df.to_excel("tech_news.xlsx", index=False)
        print("تم استخراج البيانات بنجاح وحفظها في ملف 'tech_news.xlsx'!")
    else:
        print(f"فشل الاتصال بالموقع. رمز الخطأ: {response.status_code}")

if __name__ == "__main__":
    scrape_tech_news()