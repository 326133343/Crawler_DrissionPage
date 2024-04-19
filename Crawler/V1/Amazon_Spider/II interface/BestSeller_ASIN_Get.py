from DrissionPage import ChromiumPage
import pandas as pd
import re
import os
from concurrent.futures import ThreadPoolExecutor

def generate_amazon_link(asin, base_url='https://www.amazon.com/dp/'):
    return f'{base_url}{asin}'
        
def scroll_page(page, scroll_times=6, delay=1):
    for _ in range(scroll_times):
        page.actions.scroll(0, 1000)
        page.wait(delay)

def get_asins_from_bestseller_page(page, url):
    page.get(url)
    scroll_page(page)
    page.wait(3)
    html = page.html
    asin_pattern = re.compile(r'data-asin="([A-Z0-9]{10})"')
    asins = re.findall(asin_pattern, html)
    return asins[:50]

def scrape_bestsellers(url, num_tabs=2):
    all_products = []
    main_page = ChromiumPage()
    tabs = [main_page.new_tab() for _ in range(num_tabs)] 

    with ThreadPoolExecutor(max_workers=num_tabs) as executor:
        futures = []
        for i in range(1,3):
            current_tab = tabs[i-1]
            page_url = f'{url}/ref=zg_bs_pg_{i}?_encoding=UTF8&pg={i}'
            futures.append(executor.submit(get_asins_from_bestseller_page, current_tab, page_url))

        for future in futures:
            asins = future.result()
            for asin in asins:
                all_products.append({'url': generate_amazon_link(asin), 'asin': asin})

    for tab in tabs:
        tab.close()
    main_page.quit()

    return all_products

def main(file_path):
    df = pd.read_excel(file_path)
    
    for index, row in df.iterrows():
        category_url = row['url']
        category_name = row['category_name']
        products = scrape_bestsellers(category_url)

        # 保存结果到Excel
        result_df = pd.DataFrame(products)
        file_name = f'bestsellers_{category_name}.xlsx'
        file_path = os.path.join('output', file_name)
        result_df.to_excel(file_path, index=False)

file_path = 'path_to_your_excel.xlsx'
main(file_path)
