# pip install pandas requests openpyxl Pillow
import pandas as pd
import requests
import os
from PIL import Image as PILImage
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
import shutil

# 读取表格并指定列名
df = pd.read_excel('test.xlsx', header=None, names=['col1', 'col2', 'col3', 'col4', 'image_url', 'col6'])
print("读取表格完成")

# 创建存储图片的文件夹
image_folder = 'images'
# 创建辅助文件
modified_excel_file = 'modified_excel_file.xlsx'
os.makedirs(image_folder, exist_ok=True)
print("创建存储图片的文件夹完成")

# 遍历每一行，下载图片并替换链接
for index, row in df.iterrows():
    image_url = row['image_url']  # 使用指定的列名'image_url'
    image_name = f"image_{index}.jpg"  # 为每张图片生成一个唯一的文件名
    image_path = os.path.join(image_folder, image_name)
    print(f"处理第 {index + 1} 行：{image_url}")

    # 下载图片
    response = requests.get(image_url)
    with open(image_path, 'wb') as f:
        f.write(response.content)
    print(f"下载图片完成：{image_url}")

    # 替换链接为本地路径
    df.at[index, 'image_url'] = image_path
    print(f"替换链接为本地路径完成：{image_path}")

# 保存修改后的表格
df.to_excel(modified_excel_file, index=False, header=False)  # 不包含行和列的标签
print("保存修改后的表格完成")

# 加载新的 Excel 文件
wb = load_workbook(modified_excel_file)
ws = wb.active

# 在 Excel 中插入图片并调整大小
for index, row in df.iterrows():
    # 使用 Pillow 库打开图片并转换格式
    image = PILImage.open(row['image_url'])
    image.save(row['image_url'].replace('.webp', '.jpg'))  # 将 .webp 格式转换为 .jpg 格式

    # 插入图片到 Excel 中
    img = ExcelImage(row['image_url'].replace('.webp', '.jpg'))

    # 计算调整后的图片尺寸
    width, height = img.width, img.height
    max_width, max_height = 100, 100  # 设定最大宽度和高度
    if width > max_width or height > max_height:
        ratio = min(max_width / width, max_height / height)
        img.width = int(width * ratio)
        img.height = int(height * ratio)

    # 插入到每一行的 'e' 列
    ws.add_image(img, f'E{index + 1}')  

    # 设置单元格大小
    ws.column_dimensions['E'].width = 15
    ws.row_dimensions[index + 1].height = 100

# 保存 Excel 文件
wb.save('output.xlsx')
print("保存 Excel 文件完成")

# 删除图片文件夹
if os.path.exists(image_folder):
    shutil.rmtree(image_folder)
    print(f"已删除文件夹：{image_folder}")
else:
    print(f"文件夹不存在：{image_folder}")

# 删除修改后的 Excel 文件
if os.path.exists(modified_excel_file):
    os.remove(modified_excel_file)
    print(f"已删除文件：{modified_excel_file}")
else:
    print(f"文件不存在：{modified_excel_file}")