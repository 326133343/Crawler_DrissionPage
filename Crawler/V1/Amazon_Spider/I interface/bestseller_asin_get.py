'''
BSR中的ASIN码获取
'''
from DrissionPage import ChromiumPage
import pandas as pd
import re
import os

def generate_amazon_link(asin, base_url='https://www.amazon.com/dp/'):
    return f'{base_url}{asin}'

def scroll_page(page, scroll_times=6, delay=1):
    for _ in range(scroll_times):
        page.actions.scroll(0, 1000)  # 向下滚动页面
        page.wait(delay)  # 等待页面加载

def get_asins_from_bestseller_page(page, url):
    page.get(url)
    scroll_page(page)  # 滚动页面以加载内容
    page.wait(3)  # 等待页面加载
    html = page.html
    asin_pattern = re.compile(r'data-asin="([A-Z0-9]{10})"')
    asins = re.findall(asin_pattern, html)
    return asins[:50]  # 只取前50个ASIN

def scrape_bestsellers(url):
    page = ChromiumPage()  # 创建主页面实例
    asins = get_asins_from_bestseller_page(page, url)
    page.quit()
    return [{'url': generate_amazon_link(asin), 'asin': asin} for asin in asins]

def save_to_excel(data, file_name):
    df = pd.DataFrame(data)
    output_folder = 'output'
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    file_path = os.path.join(output_folder, file_name)
    df.to_excel(file_path, index=False)

# 测试URL
url = "https://www.amazon.com/BestSellers/zgbs/hi/3744271"

# 获取并保存结果
products = scrape_bestsellers(url)
save_to_excel(products, 'bestsellers_towel.xlsx')

