import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageFilter
import img2pdf
import os
import numpy as np

pdf_path = "附件BZ015(2).pdf"
output_pdf = "去水印后.pdf"
temp_img_dir = "temp_imgs"

# 步骤1：PDF转图片
os.makedirs(temp_img_dir, exist_ok=True)
doc = fitz.open(pdf_path)
img_paths = []
for i, page in enumerate(doc):
    pix = page.get_pixmap(dpi=300)
    img_path = os.path.join(temp_img_dir, f"page_{i}.png")
    pix.save(img_path)
    img_paths.append(img_path)

# 步骤2：智能去除水印和二维码
def remove_watermark(img_path):
    # 打开图像
    img = Image.open(img_path)
    w, h = img.size
    
    # 定义水印和二维码区域（右下角25%）
    margin_x, margin_y = int(w * 0.75), int(h * 0.75)
    
    # 创建原图的副本
    result = img.copy()
    
    # 裁剪水印区域
    watermark_area = result.crop((margin_x, margin_y, w, h))
    
    # 将图像转换为numpy数组以便更精细的处理
    watermark_np = np.array(watermark_area)
    
    # 创建一个与水印区域大小相同的遮罩
    mask = np.ones(watermark_np.shape[:2], dtype=np.uint8) * 255
    
    # 对遮罩应用渐变效果，使边缘平滑
    for i in range(mask.shape[0]):
        for j in range(mask.shape[1]):
            # 创建径向渐变遮罩
            dist_x = j / mask.shape[1]
            dist_y = i / mask.shape[0]
            gradient = max(dist_x, dist_y)
            mask[i, j] = int(255 * (1 - gradient))
    
    # 将遮罩应用到水印区域
    for c in range(3):  # 处理RGB三个通道
        watermark_np[:, :, c] = (
            watermark_np[:, :, c] * (mask / 255.0) + 
            get_background_color(result, margin_x, margin_y, w, h)[c] * (1 - mask / 255.0)
        ).astype(np.uint8)
    
    # 将处理后的区域转回图像并粘贴
    processed_area = Image.fromarray(watermark_np)
    result.paste(processed_area, (margin_x, margin_y))
    
    # 保存处理后的图像
    result.save(img_path)

def get_background_color(img, margin_x, margin_y, w, h):
    """获取周围背景的平均颜色"""
    # 从图像边缘和周围区域取样
    samples = [
        img.getpixel((w-10, h-10)),   # 右下角
        img.getpixel((w-10, margin_y-10)),  # 上方
        img.getpixel((margin_x-10, h-10)),  # 左方
    ]
    
    # 计算平均颜色
    avg_color = np.mean(samples, axis=0).astype(int)
    return avg_color

# 处理每张图片
for img_path in img_paths:
    remove_watermark(img_path)

# 步骤3：图片转回PDF
with open(output_pdf, "wb") as f:
    img_bytes = [p for p in img_paths]
    f.write(img2pdf.convert(img_bytes))

# 清理临时文件
import shutil
shutil.rmtree(temp_img_dir)