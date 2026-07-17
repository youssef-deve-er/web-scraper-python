import sys
import io
import asyncio
import pandas as pd
from playwright.async_api import async_playwright
from bs4 import BeautifulSoup

# إجبار بايثون على استخدام ترميز UTF-8 لدعم اللغة العربية في الـ Terminal والـ Output
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

async def scrape_tech_news_advanced(max_pages=3):
    """
    سكربت متقدم واحترافي يعتمد على متصفح حقيقي خفي (Playwright) 
    لسحب البيانات بأمان وتخطي حمايات المواقع مع تنسيق ملف الـ Excel.
    """
    base_url = "https://news.ycombinator.com/"
    news_list = []
    
    print(f"🚀 جاري تشغيل المتصفح الذكي (المستهدف: {max_pages} صفحات)...")
    
    # تشغيل Playwright
    async with async_playwright() as p:
        # تشغيل متصفح Chromium في الخلفية (headless=True لسرعة الأداء)
        # جهازك القوي يمكنه تشغيل عشرات المتصفحات من هذا النوع في وقت واحد
        browser = await p.chromium.launch(headless=True)
        
        # إنشاء سياق متصفح جديد مع بايو وإعدادات مستخدم حقيقي لتفادي الحظر
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            viewport={"width": 1920, "height": 1080}
        )
        
        page_tab = await context.new_page()
        
        for page_num in range(1, max_pages + 1):
            url = f"{base_url}?p={page_num}" if page_num > 1 else base_url
            print(f"⏳ جاري تصفح وقراءة الصفحة رقم {page_num} عبر المتصفح...")
            
            try:
                # الانتقال للموقع وانتظار تحميل شبكة البيانات بالكامل للـ JavaScript
                await page_tab.goto(url, wait_until="networkidle", timeout=15000)
                
                # استخراج كود الـ HTML بعد معالجته بالكامل داخل المتصفح
                html_content = await page_tab.content()
                
                # تمرير الكود لـ BeautifulSoup للفرز السريع
                soup = BeautifulSoup(html_content, 'html.parser')
                articles = soup.find_all('span', class_='titleline')
                
                if not articles:
                    print("🛑 لم يتم العثور على مقالات إضافية، قد يكون هذا نهاية الصفحات المتاحة.")
                    break
                    
                for article in articles:
                    a_tag = article.find('a')
                    if a_tag:
                        title = a_tag.text.strip()
                        link = a_tag['href']
                        
                        if link.startswith('item?id='):
                            link = base_url + link
                            
                        news_list.append({
                            "رقم الصفحة (Page)": page_num,
                            "العنوان (Title)": title,
                            "الرابط (Link)": link
                        })
                        
                # مهلة تريّث عشوائية قصيرة بين الصفحات لتبدو كبشري طبيعي
                await asyncio.sleep(1.5)
                
            except Exception as e:
                print(f"❌ حدث خطأ أثناء فحص الصفحة {page_num}: {e}")
                break
                
        # إغلاق المتصفح بعد الانتهاء
        await browser.close()

    # مرحلة حفظ البيانات وتنسيقها للعميل
    if news_list:
        df = pd.DataFrame(news_list)
        file_name = "tech_news_advanced.xlsx"
        
        try:
            with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='HackerNews')
                
                # تنسيق عرض الأعمدة تلقائياً ليفتح العميل الملف ويجده منظماً
                worksheet = writer.sheets['HackerNews']
                for col in worksheet.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = col[0].column_letter
                    worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
            print(f"\n🎉 تم استخراج {len(news_list)} مقال بنجاح عبر المتصفح الحقيقي!")
            print(f"📁 تم حفظ وتنسيق الملف الاحترافي كـ '{file_name}'")
            
        except Exception as e:
            print(f"❌ فشل حفظ ملف الـ Excel: {e}")
    else:
        print("⚠️ لم يتم جمع أي بيانات لحفظها.")

if __name__ == "__main__":
    # تشغيل السكربت غير المتزامن باستخدام asyncio
    asyncio.run(scrape_tech_news_advanced(max_pages=3))
