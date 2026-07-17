import sys
import io
import requests
from bs4 import BeautifulSoup
import pandas as pd

# إجبار بايثون على استخدام ترميز UTF-8 لدعم اللغة العربية في الـ Terminal والـ Output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def scrape_tech_news(max_pages=3):
    """
    سكربت متقدم لسحب الأخبار التقنية من Hacker News مع دعم الصفحات المتعددة وتنسيق البيانات.
    :param max_pages: عدد الصفحات المراد سحبها (تم ضبطها افتراضياً على 3 صفحات)
    """
    base_url = "https://news.ycombinator.com/"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    }
    
    news_list = []
    
    print(f"🚀 جاري بدء سحب البيانات (المستهدف: {max_pages} صفحات)...")
    
    for page in range(1, max_pages + 1):
        # الموقع يستخدم بارامتر p لتحديد رقم الصفحة
        url = f"{base_url}?p={page}" if page > 1 else base_url
        print(f"⏳ جاري فحص الصفحة رقم {page}...")
        
        try:
            response = requests.get(url, headers=headers, timeout=10)
            
            if response.status_code != 200:
                print(f"⚠️ خطأ في الصفحة {page}. رمز الحالة: {response.status_code}")
                continue
                
            soup = BeautifulSoup(response.text, 'html.parser')
            articles = soup.find_all('span', class_='titleline')
            
            if not articles:
                print("🛑 لم يتم العثور على مقالات إضافية، قد يكون هذا نهاية الصفحات المتاحة.")
                break
                
            for article in articles:
                a_tag = article.find('a')
                if a_tag:
                    title = a_tag.text.strip()
                    link = a_tag['href']
                    
                    # إصلاح الروابط الداخلية (لو كان الرابط يؤدي لنفس الموقع)
                    if link.startswith('item?id='):
                        link = base_url + link
                        
                    news_list.append({
                        "رقم الصفحة (Page)": page,
                        "العنوان (Title)": title,
                        "الرابط (Link)": link
                    })
                    
        except requests.exceptions.RequestException as e:
            print(f"❌ حدث خطأ في الاتصال بالشبكة أثناء فحص الصفحة {page}: {e}")
            break

    # التحقق من وجود بيانات قبل الحفظ
    if news_list:
        df = pd.DataFrame(news_list)
        file_name = "tech_news.xlsx"
        
        # حفظ الملف وتنسيقه تلقائياً باستخدام ExcelWriter لتعديل عرض الأعمدة
        try:
            with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='HackerNews')
                
                # كود احترافي لضبط عرض الأعمدة تلقائياً بناءً على طول النص لراحة العميل
                worksheet = writer.sheets['HackerNews']
                for col in worksheet.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = col[0].column_letter
                    worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
            print(f"\n🎉 تم استخراج {len(news_list)} مقال بنجاح!")
            print(f"📁 تم حفظ وتنسيق الملف كـ '{file_name}' المخصص للعميل.")
            
        except Exception as e:
            print(f"❌ فشل حفظ ملف الـ Excel. تأكد من إغلاق الملف إذا كان مفتوحاً: {e}")
    else:
        print("⚠️ لم يتم جمع أي بيانات لحفظها.")

if __name__ == "__main__":
    # يمكنك زيادة عدد الصفحات كما تحب، جهازك القوي سيتحمل معالجة البيانات بكفاءة
    scrape_tech_news(max_pages=3)
