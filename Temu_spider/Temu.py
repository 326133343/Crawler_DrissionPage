from DrissionPage import ChromiumPage
import time
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urljoin
import requests
from pathlib import Path
from PIL import Image
import io

def download_image(image_url, folder="images"):
    Path(folder).mkdir(parents=True, exist_ok=True)
    image_url = image_url.strip("'\"")
    # 生成唯一的文件名
    file_name = f"{int(time.time() * 1000)}.jpg"  # 使用时间戳作为文件名
    file_path = f"{folder}/{file_name}"

    try:
        response = requests.get(image_url, headers={'User-Agent': 'Mozilla/5.0'})
        if response.status_code == 200:
            # 将webp图片数据加载到一个Pillow图像对象中
            image = Image.open(io.BytesIO(response.content))
            # 转换图片格式为JPG并保存
            rgb_im = image.convert('RGB')
            rgb_im.save(file_path, format='JPEG')
        return file_path
    except Exception as e:
        print(f"Error downloading {image_url}: {e}")
        return None

def scrape_data(page):
    target_url = "https://www.temu.com/pl-en/channel/best-sellers.html?filter_items=1%3A1&scene=home_title_bar_recommend&refer_page_el_sn=201341&refer_page_name=home&refer_page_id=10005_1710234521439_ggpidxmgdf&refer_page_sn=10005&_x_sessn_id=6ch12q7zxb"
    page.get(target_url)
    time.sleep(120)  # 给页面加载预留时间

    # 再次获取更新后的页面内容
    soup = BeautifulSoup(page.html, 'html.parser')
    
    products = []
    containers = soup.select('._6q6qVUF5._1UrrHYym')
    print(f"找到{len(containers)}个商品容器")
    
    for container in containers:
        partial_url = container.select_one('a._2Tl9qLr1._3ZME5MBZ')['href']
        full_url = urljoin("https://www.temu.com", partial_url)

        image_element = container.select_one('img.goods-img-external')
        image_url = image_element['src'] if image_element and 'src' in image_element.attrs else "图片URL缺失"
        image_path = download_image(image_url) if image_url != "图片URL缺失" else image_url

        title_element = container.select_one('h3._2BvQbnbN')
        title = title_element.get_text(strip=True) if title_element else "未查询到标题"

        price_element = container.select_one('span.LiwdOzUs')
        price = price_element.get_text(strip=True) if price_element else "未查询到价格"

        purchase_count_element = container.select_one('span._3vfo0XTx')
        purchase_count = purchase_count_element.get_text(strip=True) if purchase_count_element else "未查询到购买人数"

        products.append([image_path,title,full_url,price, purchase_count ])

    return products


def main():
    page = ChromiumPage()  # 开启headless模式以便观察
    products = scrape_data(page)
    if products:
        df = pd.DataFrame(products, columns=['商品图片', '标题', 'URL', '价格(PL)','购买人数'])
        df.to_excel(r"C:\Users\Administrator\Desktop\python\TEMU\product_info\Cell Phone & Accessories.xlsx", index=False)
        print("数据已保存")
    else:
        print("没有提取到任何商品数据")


if __name__ == '__main__':
    main()
