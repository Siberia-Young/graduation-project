from selenium import webdriver
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import time

filename = "data/jd/京东_华为移动快充_2023-11-20_11-17-29_(5925 of 5970).xlsx"
workbook = load_workbook(filename)
sheet = workbook.active
start_time = time.time()

start_row = 2
end_row = sheet.max_row
aim_col = 10

total = end_row - start_row + 1
current = 0

options = webdriver.FirefoxOptions()
driver = webdriver.Remote(command_executor="http://127.0.0.1:4444", options=options)

# options = webdriver.FirefoxOptions()
# driver = webdriver.Firefox(options=options)

try:
    for row in range(start_row, end_row + 1):
        current+=1
        res = (total - current) / (current / ((time.time() - start_time) / 60))
        print(f"\r当前进度：{current}/{total}，预计仍需：{res:.2f} min", end="")
        goods_link = sheet.cell(row=row, column=aim_col).value
        driver.get(goods_link)
        time.sleep(0.3)
        tempHTML = driver.execute_script("return document.documentElement.outerHTML")
        tempSoup = BeautifulSoup(tempHTML, "html.parser")

        elements = tempSoup.select('div.hxm_hide_page')
        if len(elements) == 0:
            elements = tempSoup.select('div.itemover-tip')
        if len(elements) == 0:
            try:
                goods_brand_element = tempSoup.find_all('ul',id='parameter-brand')
                goods_brand = len(goods_brand_element) == 0 and '暂无' or goods_brand_element[0].select('li a')[0].text
                sheet.cell(row=row, column=9, value=goods_brand)
                if sheet.cell(row=row, column=4).value is None:
                    shop_element = tempSoup.select('div.popbox-inner h3 a')
                    
                    shop_name = len(shop_element) == 0 and '暂无' or shop_element[0].text
                    shop_name_cell = sheet[f"{get_column_letter(4)}{row}"]
                    shop_name_cell.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
                    sheet.cell(row=row, column=4, value=shop_name)

                    shop_link = len(shop_element) == 0 and '暂无' or ('https:'+shop_element[0].get('href'))
                    shop_link_cell = sheet[f"{get_column_letter(5)}{row}"]
                    shop_link_cell.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
                    shop_link_cell.font = Font(underline="single", color="0563C1")
                    shop_link_cell.hyperlink = shop_link
                    sheet.cell(row=row, column=5, value=shop_link)
            except:
                workbook.save(filename)
                driver.quit()
                print('与现有浏览器连接断开')
        else:
            sheet.cell(row=row, column=12, value='delete')
except:
    print('连接超时')
    print('主动中断')
finally:
    # 保存文件
    workbook.save(filename)
    driver.quit()
    print('与现有浏览器连接断开')
    end_time = time.time()
    duration = end_time - start_time
    print(f"爬虫耗时：{duration:.2f} 秒")
    print(f"目标数量：{total} 条")
    print(f"已获取数量：{current} 条")
    unit = current / (duration / 60)
    print(f"每分钟爬取数量：{unit:.2f} 条")
