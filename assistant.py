import fitz  # PyMuPDF
from PIL import Image
import os
import datetime
import argparse
from tqdm import tqdm

# 定义 A4 纸尺寸（300 DPI）
A4_WIDTH = 2480  # 8.27 inches * 300 dpi
A4_HEIGHT = 3508  # 11.69 inches * 300 dpi
MARGIN = 118  # 1 cm ≈ 118 pixels at 300 DPI
MIN_SPACE_FOR_COLLAGE = 210  # 7 cm in pixels

CONTENT_WIDTH = A4_WIDTH - 2 * MARGIN  # 2244
CONTENT_HEIGHT = A4_HEIGHT - 2 * MARGIN  # 3272

folder_count = 0
success_folders = []
ignored_folders = []

def log_debug(message):
    if debug_mode:
        print(message)

def create_collage_image(image_files, max_width, cell_height):
    log_debug(f'创建拼图 {image_files} (最大宽度 {max_width}, 单元高度 {cell_height})')
    images = []
    for image_file in image_files:
        try:
            img = Image.open(image_file)
            images.append(img)
        except Exception as e:
            print(f"打开图片错误 {image_file}: {e}")

    if len(images) == 0:
        return None

    num_images = len(images)

    if 2 <= num_images <= 4:
        # 将 2 到 4 张图像按行排列
        target_height = cell_height
        scaled_images = []
        total_width = 0
        for img in images:
            scale_factor = target_height / img.height
            new_size = (int(img.width * scale_factor), int(img.height * scale_factor))
            resized_image = img.resize(new_size, Image.LANCZOS)
            scaled_images.append(resized_image)
            total_width += resized_image.width

        if total_width > max_width:
            scale_factor = max_width / total_width
            for i in range(len(scaled_images)):
                new_size = (int(scaled_images[i].width * scale_factor), int(scaled_images[i].height * scale_factor))
                scaled_images[i] = scaled_images[i].resize(new_size, Image.LANCZOS)

        collage_width = sum(img.width for img in scaled_images)
        collage_image = Image.new('RGB', (collage_width, target_height), (255, 255, 255))
        x_offset = 0
        for img in scaled_images:
            collage_image.paste(img, (x_offset, (target_height - img.height) // 2))
            x_offset += img.width

    else:
        # 对超过 4 张的图像进行 4 列网格排列
        images_per_row = 4
        target_height_per_image = cell_height // (num_images // images_per_row + (1 if num_images % images_per_row != 0 else 0))
        scaled_images = []
        
        for img in images:
            scale_factor = target_height_per_image / img.height
            new_size = (int(img.width * scale_factor), target_height_per_image)
            scaled_images.append(img.resize(new_size, Image.LANCZOS))

        # 计算拼图的总高度
        rows_needed = (num_images + images_per_row - 1) // images_per_row
        collage_height = target_height_per_image * rows_needed
        collage_image = Image.new('RGB', (max_width, collage_height), (255, 255, 255))

        # 按照 4 列的网格排布图像
        x_offset = y_offset = 0
        for idx, img in enumerate(scaled_images):
            if x_offset + img.width > max_width or (idx % images_per_row == 0 and idx != 0):
                x_offset = 0
                y_offset += target_height_per_image
            collage_image.paste(img, (x_offset, y_offset))
            x_offset += img.width

    return collage_image

def merge_invoice_and_images_to_total_pdf(folder_path, doc):
    global folder_count
    try:
        log_debug(f"\n正在处理文件夹: {folder_path}")
        files = os.listdir(folder_path)
        pdf_files = [os.path.join(folder_path, f) for f in files if f.lower().endswith('.pdf')]
        image_files = [os.path.join(folder_path, f) for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg'))]

        # 筛选 NEWPAGE 和 NEWLINE 图片
        newline_images = [f for f in image_files if os.path.basename(f).startswith('NEWLINE')]
        newpage_images = [f for f in image_files if os.path.basename(f).startswith('NEWPAGE')]
        other_images = [f for f in image_files if f not in newline_images and f not in newpage_images]

        if len(pdf_files) != 1 or len(other_images) < 2:
            reason = f"仅找到 {len(pdf_files)} 个 PDF 文件与 {len(other_images)} 个图片文件."
            ignored_folders.append((folder_path, reason))
            log_debug(f"Ignored {folder_path}: {reason}")
            return 0

        invoice_pdf_path = pdf_files[0]
        invoice_doc = fitz.open(invoice_pdf_path)
        invoice_page = invoice_doc.load_page(0)
        
        scale = 5
        matrix = fitz.Matrix(scale, scale)
        pix = invoice_page.get_pixmap(matrix=matrix)
        invoice_image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        scale_factor = CONTENT_WIDTH / invoice_image.width
        new_height = int(invoice_image.height * scale_factor)
        resized_invoice_image = invoice_image.resize((CONTENT_WIDTH, new_height), Image.LANCZOS)

        remaining_space = CONTENT_HEIGHT - resized_invoice_image.height
        create_new_page_for_collage = remaining_space < MIN_SPACE_FOR_COLLAGE
        
        collage_image = create_collage_image(other_images, CONTENT_WIDTH, remaining_space if not create_new_page_for_collage else CONTENT_HEIGHT)
        if collage_image is None:
            log_debug("没有足够的图片，无法创建拼图.")
            return 0

        if create_new_page_for_collage:
            merged_image = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))
            merged_image.paste(resized_invoice_image, (MARGIN, MARGIN))
            output_pdf_path = os.path.join(folder_path, 'invoice_page.pdf')
            merged_image.save(output_pdf_path, 'PDF', resolution=300.0)
            doc.insert_pdf(fitz.open(output_pdf_path))
            os.remove(output_pdf_path)

            collage_page = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))
            collage_x_offset = (A4_WIDTH - collage_image.width) // 2
            collage_y_offset = (A4_HEIGHT - collage_image.height) // 2
            collage_page.paste(collage_image, (collage_x_offset, collage_y_offset))
            collage_pdf_path = os.path.join(folder_path, 'collage_page.pdf')
            collage_page.save(collage_pdf_path, 'PDF', resolution=300.0)
            doc.insert_pdf(fitz.open(collage_pdf_path))
            os.remove(collage_pdf_path)
        else:
            merged_image = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))
            merged_image.paste(resized_invoice_image, (MARGIN, MARGIN))
            collage_y_offset = MARGIN + new_height + (remaining_space - collage_image.height) // 2
            merged_image.paste(collage_image, (MARGIN, collage_y_offset))
            output_pdf_path = os.path.join(folder_path, 'combined_page.pdf')
            merged_image.save(output_pdf_path, 'PDF', resolution=300.0)
            doc.insert_pdf(fitz.open(output_pdf_path))
            os.remove(output_pdf_path)

        # 处理 NEWLINE 图片
        for newline_image_path in newline_images:
            img = Image.open(newline_image_path)
            scale_factor = CONTENT_WIDTH / img.width
            resized_newline_image = img.resize((CONTENT_WIDTH, int(img.height * scale_factor)), Image.LANCZOS)

            newline_page = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))
            newline_y_offset = (A4_HEIGHT - resized_newline_image.height) // 2
            newline_page.paste(resized_newline_image, (MARGIN, newline_y_offset))
            newline_pdf_path = os.path.join(folder_path, 'newline_page.pdf')
            newline_page.save(newline_pdf_path, 'PDF', resolution=300.0)
            doc.insert_pdf(fitz.open(newline_pdf_path))
            os.remove(newline_pdf_path)

        # 处理 NEWPAGE 图片
        for newpage_image_path in newpage_images:
            img = Image.open(newpage_image_path)
            scale_factor = CONTENT_WIDTH / img.width
            resized_newpage_image = img.resize((CONTENT_WIDTH, int(img.height * scale_factor)), Image.LANCZOS)

            newpage = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))
            newpage_y_offset = (A4_HEIGHT - resized_newpage_image.height) // 2
            newpage.paste(resized_newpage_image, (MARGIN, newpage_y_offset))
            newpage_pdf_path = os.path.join(folder_path, 'newpage_image.pdf')
            newpage.save(newpage_pdf_path, 'PDF', resolution=300.0)
            doc.insert_pdf(fitz.open(newpage_pdf_path))
            os.remove(newpage_pdf_path)

        folder_count += 1
        success_folders.append(folder_path)
        return 1

    except Exception as e:
        ignored_folders.append((folder_path, str(e)))
        log_debug(f"处理文件夹错误 {folder_path}: {e}")
        return 0

def process_all_subfolders_to_total_pdf(base_folder, output_path):
    global folder_count
    doc = fitz.open()

    # 获取上级文件夹的名称
    parent_folder_name = os.path.basename(os.path.abspath(base_folder))

    subfolders = [os.path.join(base_folder, subfolder) for subfolder in os.listdir(base_folder) if os.path.isdir(os.path.join(base_folder, subfolder))]

    for subfolder_path in tqdm(subfolders, desc="正在处理文件夹"):
        merge_invoice_and_images_to_total_pdf(subfolder_path, doc)

    if folder_count > 0:
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        # 使用上级文件夹名称作为文件名前缀
        default_output_filename = f'{parent_folder_name}_报销单_自动生成_{folder_count}张发票_{timestamp}.pdf'
        
        # 根据不同的output_path参数值判断保存路径
        if os.path.isdir(output_path):
            output_pdf = os.path.join(output_path, default_output_filename)
        elif os.path.splitext(output_path)[1].lower() == '.pdf':
            output_pdf = output_path
            # 检查文件是否存在，如果存在则询问覆盖
            if os.path.exists(output_pdf):
                user_input = input(f"{output_pdf} 已存在，是否覆盖？ (y/n): ").strip().lower()
                if user_input != 'y':
                    print("操作已取消。")
                    doc.close()
                    return
        else:
            output_pdf = os.path.join('./', default_output_filename) if output_path == '' else output_path

        doc.save(output_pdf)
        doc.close()
        print(f"成功创建 {output_pdf}")

        # 打印被忽略的文件夹信息
        if len(ignored_folders) > 0:
            print("\n以下文件夹被忽略：")
            for folder_path, reason in ignored_folders:
                print(f"{folder_path}: {reason}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='合并发票和图片为 PDF')
    parser.add_argument('--debug', '-D', action='store_true', help='启用调试信息')
    parser.add_argument('--input-folder', '-I', type=str, default='./', help='输入文件夹路径（默认当前文件夹）')
    parser.add_argument('--output', '-O', type=str, default='', help='输出 PDF 文件路径或文件名（默认自动生成）')

    args = parser.parse_args()
    debug_mode = args.debug

    input_folder = args.input_folder
    output = args.output if args.output else ''

    process_all_subfolders_to_total_pdf(input_folder, output)
