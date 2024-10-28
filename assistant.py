from PyPDF2 import PdfMerger
from PIL import Image
import os
import datetime

A4_WIDTH = 2480
A4_HEIGHT = 3508
MARGIN = 118  # 定义 1 厘米的页边距

CONTENT_WIDTH = A4_WIDTH - 2 * MARGIN  # 可用内容宽度
CONTENT_HEIGHT = A4_HEIGHT - 2 * MARGIN  # 可用内容高度

folder_count = 0

def merge_images_to_a4_page(image_files):
    images = [Image.open(image_file) for image_file in image_files]

    num_images = len(images)
    if num_images == 2:
        # 横向拼接两张图片
        max_height = CONTENT_HEIGHT  # 修改这里
        max_width = CONTENT_WIDTH // 2  # 修改这里

        resized_images = []
        for image in images:
            scale_factor = min(max_width / image.width, max_height / image.height)
            new_size = (int(image.width * scale_factor), int(image.height * scale_factor))
            resized_image = image.resize(new_size, Image.LANCZOS)
            resized_images.append(resized_image)

        # 创建A4大小的画布
        merged_image = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))

        # 拼接图片
        x_offset = MARGIN  # 起始位置加上左边距
        for image in resized_images:
            y_offset = MARGIN + (CONTENT_HEIGHT - image.height) // 2  # 垂直居中，考虑上边距
            merged_image.paste(image, (x_offset, y_offset))
            x_offset += max_width  # 移动到下一个位置
    elif num_images in [3, 4]:
        # 2x2 网格拼接
        cols = 2
        rows = 2
        max_cell_width = CONTENT_WIDTH // cols  # 修改这里
        max_cell_height = CONTENT_HEIGHT // rows  # 修改这里

        resized_images = []
        for image in images:
            scale_factor = min(max_cell_width / image.width, max_cell_height / image.height)
            new_size = (int(image.width * scale_factor), int(image.height * scale_factor))
            resized_image = image.resize(new_size, Image.LANCZOS)
            resized_images.append(resized_image)

        # 创建A4大小的画布
        merged_image = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))

        # 拼接图片到网格中
        positions = [
            (MARGIN, MARGIN),
            (MARGIN + max_cell_width, MARGIN),
            (MARGIN, MARGIN + max_cell_height),
            (MARGIN + max_cell_width, MARGIN + max_cell_height)
        ]
        for idx, image in enumerate(resized_images):
            x_offset, y_offset = positions[idx]
            x_padding = (max_cell_width - image.width) // 2
            y_padding = (max_cell_height - image.height) // 2
            merged_image.paste(image, (x_offset + x_padding, y_offset + y_padding))
    else:
        # 对于其他数量的图片，横向拼接
        max_height = CONTENT_HEIGHT
        total_width = sum(image.width for image in images)

        # 如果总宽度超过内容宽度，则按比例缩放
        if total_width > CONTENT_WIDTH:
            scale_factor = CONTENT_WIDTH / total_width
            images = [image.resize((int(image.width * scale_factor), int(image.height * scale_factor)), Image.LANCZOS) for image in images]
            total_width = sum(image.width for image in images)
            max_height = max(image.height for image in images)

        # 创建A4大小的画布
        merged_image = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))

        x_offset = MARGIN + (CONTENT_WIDTH - total_width) // 2
        y_offset = MARGIN + (CONTENT_HEIGHT - max_height) // 2
        for image in images:
            merged_image.paste(image, (int(x_offset), int(y_offset)))
            x_offset += image.width

    return merged_image


def merge_invoice_and_images_to_total_pdf(folder_path, merger):
    files = os.listdir(folder_path)
    pdf_files = [os.path.join(folder_path, f) for f in files if f.endswith('.pdf')]
    image_files = [os.path.join(folder_path, f) for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg'))]

    # 检查是否符合条件（一个PDF和至少两张图片）
    if len(pdf_files) != 1 or len(image_files) < 2:
        print(f"已忽略{folder_path}：不符合要求，只找到 {len(pdf_files)} 个PDF和 {len(image_files)} 张图片。")
        return 0

    # 合并发票PDF
    invoice_pdf_path = pdf_files[0]
    with open(invoice_pdf_path, 'rb') as f:
        merger.append(f)

    # 将所有图片合并到一页
    merged_image = merge_images_to_a4_page(image_files)
    # 将合并后的图片保存为PDF格式
    image_pdf_path = os.path.join(folder_path, 'temp_image_merge.pdf')
    merged_image.save(image_pdf_path, 'PDF')

    # 将合并的图片PDF添加到合并器，并确保文件关闭
    with open(image_pdf_path, 'rb') as f:
        merger.append(f)

    # 删除临时合成的PDF
    os.remove(image_pdf_path)

    return 1

def process_all_subfolders_to_total_pdf(base_folder):
    global folder_count
    merger = PdfMerger()

    for subfolder in os.listdir(base_folder):
        subfolder_path = os.path.join(base_folder, subfolder)
        if os.path.isdir(subfolder_path):
            folder_count += merge_invoice_and_images_to_total_pdf(subfolder_path, merger)

    # 写入总的PDF文件
    output_pdf = f'./报销单_自动生成_{folder_count}张发票_{datetime.datetime.now().strftime("%Y%m%d%H%M%S")}.pdf'
    with open(output_pdf, 'wb') as output_file:
        merger.write(output_file)
    merger.close()


base_folder = './'

process_all_subfolders_to_total_pdf(base_folder)

