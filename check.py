import os
import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from PIL import Image
from io import BytesIO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as OpenpyxlImage

# 初始化Excel工作簿
wb = Workbook()
ws = wb.active
ws.append(["序号", "URL", "状态码", "访问结果", "页面标题", "截图"])

# 设置Selenium浏览器选项
options = webdriver.ChromeOptions()
options.add_argument("--headless")  # 无头模式

# 初始化WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# 读取URL列表
with open('list.txt', 'r') as f:
    urls = [url.strip() for url in f.readlines() if url.strip()]

for index, url in enumerate(urls):
    print(f"正在访问: {url}")

    # 使用requests检查状态码
    try:
        resp = requests.get(url)
        status_code = resp.status_code
    except requests.RequestException as e:
        status_code = str(e)
        print(f"{url} 请求失败: {status_code}")
        ws.append([index + 1, url, status_code, "失败", "", ""])
        continue

    # 检查状态码是否为200系列
    if 200 <= status_code < 300:
        print(f"{url} 返回状态码: {status_code}，访问成功")
        access_result = "成功"
        try:
            driver.get(url)
            page_title = driver.title
            img_path = f"screenshot_{index + 1}.png"
            driver.save_screenshot(img_path)
            img = Image.open(img_path)
            img = img.resize((int(img.width * 0.75), int(img.height * 0.75)))  # 调整图片大小
            bio = BytesIO()
            img.save(bio, format="PNG")
            img_stream = BytesIO()
            img.save(img_stream, format="PNG")
            img_stream.seek(0)
            img_id = wb.add_image(OpenpyxlImage(img_stream)).id
            img = OpenpyxlImage(img_stream)
            img.width = 400000
            img.height = 720000  # 设置图片高度最多2厘米（约720000EMU）
            ws.add_image(img, f"F{index + 2}")
            ws.append([index + 1, url, status_code, access_result, page_title, ""])
        except Exception as e:
            print(f"截图失败: {e}")
            ws.append([index + 1, url, status_code, access_result, page_title, ""])
    else:
        print(f"{url} 返回状态码: {status_code}，访问失败")
        access_result = "失败"
        ws.append([index + 1, url, status_code, access_result, "", ""])

# 保存Excel文件
wb.save("访问结果.xlsx")
print("所有URL访问完毕，结果已保存到Excel文件。")
driver.quit()  # 最后关闭浏览器