#C:/Users/Administrator/Desktop/python/amazon/products_us.xlsx

from DrissionPage import ChromiumPage
import pandas as pd
import numpy as np
from concurrent.futures import ThreadPoolExecutor
from parsel import Selector
from urllib.parse import urlparse, unquote,urljoin
import re
import os
import requests
from PIL import Image
from io import BytesIO

def get_variation_count(response, selector):
    var = response.css(selector)
    return max(len(var.css('li')), len(var.css('option')), 1) if var else 1

def extract_category_name(url):
    parsed_url = urlparse(url)
    path = parsed_url.path
    parts = path.split('/')
    return unquote(parts[-3]) if len(parts) > 3 else "Unknown"

def clean_brand(brand_text):
    if brand_text.startswith('Visit the'):
        return brand_text.split(' ')[2]
    elif 'Brand:' in brand_text:
        return brand_text.split(': ')[1]
    else:
        return brand_text

def download_image(image_url, save_path):
    directory = os.path.dirname(save_path)
    if not os.path.exists(directory):
        os.makedirs(directory)

    try:
        response = requests.get(image_url)
        image = Image.open(BytesIO(response.content))
        image.save(save_path)
        return True
    except Exception as e:
        print(f"下载图片失败：{e}")
        return False
    
def scrape_product_info(html, base_url,tab):
    response = Selector(text=html)
    
    if not response.css('#productTitle::text') and not response.css('#corePrice_feature_div'):
        return None
    #避免商品下架导致死循环
    
    img_url = response.css('img#landingImage::attr(src)').get()
    if not img_url:
        print("图片URL未找到")
        img_local_path = "图片URL未找到"
    else:
        img_filename = img_url.split('/')[-1]
        img_local_path = os.path.join('downloaded_images', img_filename)
        if not download_image(img_url, img_local_path):
            img_local_path = "下载失败"
    title = response.css('#productTitle::text').get().strip() if response.css('#productTitle::text') else 'Title not found'
    
    price = ''.join(response.css('#corePrice_feature_div>div>div>span>span *::text').get()) if response.css('#corePrice_feature_div>div>div>span>span *::text') else '未知'
    
    bought_in_past_month_selector = 'span#social-proofing-faceout-title-tk_bought span'
    bought_in_past_month = response.css(bought_in_past_month_selector).get().strip() if response.css(bought_in_past_month_selector) else '信息未提供'

    
    size_count = get_variation_count(response, '#variation_size_name')
    color_count = get_variation_count(response, '#variation_color_name')
    style_count = get_variation_count(response, '#variation_style_name')
    package_count = get_variation_count(response, '#variation_item_package_quantity')
    variant_count = size_count * color_count * style_count * package_count
    
    category_href = response.css('#wayfinding-breadcrumbs_feature_div li a::attr(href)').getall()[-1] if response.css('#wayfinding-breadcrumbs_feature_div li a::attr(href)') else None
    category_full_url = urljoin(base_url, category_href) if category_href else '' 
    category = extract_category_name(category_full_url) if category_full_url else 'Unknown'
    
    review_count = response.css('#acrCustomerReviewText::text').get() if response.css('#acrCustomerReviewText::text') else 0
    
    star = response.css('#acrPopover a>span::text').get().strip() if response.css('#acrPopover a') else ''
    
    raw_brand = response.css('#bylineInfo').css('::text').get() if response.css('#bylineInfo') else '无品牌'
    brand = clean_brand(raw_brand)
    
    bsr_text = response.css('table#productDetails_detailBullets_sections1 tr>td>span>span::text').get()
    bsr = bsr_text.split('(')[0].strip() if bsr_text else '无排名'

    details_texts = response.css('table#productDetails_detailBullets_sections1 td::text').getall() + response.css('#detailBulletsWrapper_feature_div>#detailBullets_feature_div>ul>li>span *::text').getall()
    begin = '无上架时间'
    for text in details_texts:
        date_match = re.search(r"\d{1,2}[.,]? \w{1,20}[.,]? \d{4}", text) or re.search(r"\w{1,20}[.,]? \d{1,2}[.,]? \d{4}", text)
        if date_match:
            begin = date_match.group().strip()
            break
    
    delivery = response.css('div.offer-display-feature-text[offer-display-feature-name="desktop-fulfiller-info"] span::text').get() or ""
    
    if 'Amazon' in delivery:
        delivery = 'FBA'
    else:
        delivery = 'FBM'
    
    return {
        '图片本地路径': img_local_path,
        '标题': title,
        '价格': price,
        '类目': category,
        '评论数': review_count,
        '评分': star,
        '品牌': brand,
        'BSR排名': bsr,
        '上架时间': begin,
        '变体数': variant_count,
        '配送方式': delivery,
        '上个月购买人数':bought_in_past_month
    }

def process_tab(page, df_slice):
    results = []
    tab = page.new_tab()
    for index, row in df_slice.iterrows():
        url = row['link']
        tab.get(url)
        tab.wait(20)
        html = tab.html
        product_info = scrape_product_info(html, url, tab)
        if product_info:
            data = {**row.to_dict(), **product_info}
            results.append(data)
    tab.close()
    return pd.DataFrame(results)

def process_file(page, input_file, output_folder):
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
    input_folder = 'C:/Users/Administrator/Desktop/python/amazon/input'
    output_folder = 'C:/Users/Administrator/Desktop/python/amazon/output'
    main(input_folder, output_folder)
