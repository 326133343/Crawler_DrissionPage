from DrissionPage import ChromiumPage
from concurrent.futures import ThreadPoolExecutor
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
import time
import random
import os
 
def scrape_data_from_eMAG(url,tab):
    tab.get(url)
    time.sleep(random.randint(15, 25))  
    soup = BeautifulSoup(tab.html, 'html.parser')
    
    rating = soup.select_one('p.review-rating-data').text.strip() if soup.select_one('p.review-rating-data') else '无法获取'
    positive_rate = soup.select_one('div.recommending-reviews.js-recommending-reviews p.semibold').text.strip() if soup.select_one('div.recommending-reviews.js-recommending-reviews p.semibold') else '无法获取'
    comments_count = soup.select_one('div.verified-reviews.js-verified-reviews p.semibold').text.strip() if soup.select_one('div.verified-reviews.js-verified-reviews p.semibold') else '无法获取'
    stock_status = "有存货" if soup.select_one('p.stock-and-genius span.label-in_stock') else "无存货"

    return {
        'URL': url,
        '评分': rating,
        '好评率': positive_rate,
        '评论数': comments_count,
        '存货状况': stock_status
    }


def process_tab(page, df_slice):
    results = []
    tab = page.new_tab()
    for index, row in df_slice.iterrows():
        url = row['URL']
        product_info = scrape_data_from_eMAG(url, tab)
        if product_info:
            data = {**row.to_dict(), **product_info}
            results.append(data)
    tab.close()
    return pd.DataFrame(results)

def process_file(page, input_file, output_folder):
    """处理单个文件，并保存更新后的数据到指定输出文件夹"""
    df = pd.read_excel(input_file)
    num_tabs = 15  #并发数
    df_slices = np.array_split(df, num_tabs)
    
    all_results = []
    with ThreadPoolExecutor(max_workers=num_tabs) as executor:
        futures = [executor.submit(process_tab, page, df_slice) for df_slice in df_slices]
        for future in futures:
            result = future.result()
            all_results.append(result)
    
    updated_df = pd.concat(all_results, ignore_index=True)
    output_file = os.path.join(output_folder, f'updated_{os.path.basename(input_file)}')
    updated_df.to_excel(output_file, index=False)
    print(f'Updated data for {os.path.basename(input_file)} saved to {output_file}')

def main(input_folder, output_folder):
    """遍历指定文件夹中的所有 Excel 文件，并更新数据"""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    page = ChromiumPage()

    for file_name in os.listdir(input_folder):
        if file_name.endswith('.xlsx'):
            input_file = os.path.join(input_folder, file_name)
            process_file(page, input_file, output_folder)
    
    page.quit()

if __name__ == "__main__":
    input_folder = 'C:/Users/Administrator/Desktop/python/eMAG/eMAG_First'
    output_folder = 'C:/Users/Administrator/Desktop/python/eMAG/eMAG_Second'
    main(input_folder, output_folder)
