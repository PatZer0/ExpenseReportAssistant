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
    log_debug(f'Creating collage with images {image_files} (max width {max_width}, cell height {cell_height})')
    images = []
    for image_file in image_files:
        try:
            img = Image.open(image_file)
            # if img.height < img.width:
            #     img = img.rotate(90, expand=True)
            images.append(img)
        except Exception as e:
            print(f"Error opening image {image_file}: {e}")

    if len(images) == 0:
        return None

    # Handle image arrangement based on the number of images
    num_images = len(images)
    
    if 2 <= num_images <= 4:
        # Scale all images to the same height
        target_height = cell_height
        scaled_images = []
        total_width = 0
        for img in images:
            scale_factor = target_height / img.height
            new_size = (int(img.width * scale_factor), int(img.height * scale_factor))
            resized_image = img.resize(new_size, Image.LANCZOS)
            scaled_images.append(resized_image)
            total_width += resized_image.width

        # Adjust collage width to stay within max_width
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
        # More than 4 images - arrange in a grid to fit within max_width
        images_per_row = 2  # Adjust as needed
        target_height_per_image = cell_height // (num_images // images_per_row)
        scaled_images = []
        collage_height = target_height_per_image * (num_images // images_per_row)

        for img in images:
            scale_factor = target_height_per_image / img.height
            new_size = (int(img.width * scale_factor), target_height_per_image)
            scaled_images.append(img.resize(new_size, Image.LANCZOS))

        collage_image = Image.new('RGB', (max_width, collage_height), (255, 255, 255))
        x_offset = y_offset = 0

        for idx, img in enumerate(scaled_images):
            if x_offset + img.width > max_width:
                x_offset = 0
                y_offset += target_height_per_image
            collage_image.paste(img, (x_offset, y_offset))
            x_offset += img.width

    return collage_image

def merge_invoice_and_images_to_total_pdf(folder_path, doc):
    global folder_count
    try:
        log_debug(f"\nProcessing folder: {folder_path}")
        files = os.listdir(folder_path)
        pdf_files = [os.path.join(folder_path, f) for f in files if f.lower().endswith('.pdf')]
        image_files = [os.path.join(folder_path, f) for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg'))]

        if len(pdf_files) != 1 or len(image_files) < 2:
            reason = f"Found {len(pdf_files)} PDFs and {len(image_files)} images."
            ignored_folders.append((folder_path, reason))
            log_debug(f"Ignored {folder_path}: {reason}")
            return 0

        invoice_pdf_path = pdf_files[0]
        invoice_doc = fitz.open(invoice_pdf_path)
        invoice_page = invoice_doc.load_page(0)
        
        # Convert invoice PDF to image
        scale = 5
        matrix = fitz.Matrix(scale, scale)
        pix = invoice_page.get_pixmap(matrix=matrix)
        invoice_image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        log_debug(f"Original invoice image size: {invoice_image.size}")

        # Resize invoice to fit CONTENT_WIDTH
        scale_factor = CONTENT_WIDTH / invoice_image.width
        new_height = int(invoice_image.height * scale_factor)
        resized_invoice_image = invoice_image.resize((CONTENT_WIDTH, new_height), Image.LANCZOS)
        log_debug(f"Resized invoice image size: {resized_invoice_image.size}")

        # Calculate remaining space below the invoice
        remaining_space = CONTENT_HEIGHT - resized_invoice_image.height
        log_debug(f"Remaining space below invoice: {remaining_space}")
        
        # Check if we need a separate page for the collage
        if remaining_space < MIN_SPACE_FOR_COLLAGE:
            create_new_page_for_collage = True
            log_debug("Remaining space is insufficient; creating a new page for the collage.")
        else:
            create_new_page_for_collage = False
        
        # Create collage image
        collage_image = create_collage_image(image_files, CONTENT_WIDTH, remaining_space if not create_new_page_for_collage else CONTENT_HEIGHT)
        if collage_image is None:
            log_debug("Failed to create collage image; possibly no valid images.")
            return 0
        log_debug(f"Collage image size: {collage_image.size}")

        if create_new_page_for_collage:
            # First page: invoice only
            merged_image = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))
            merged_image.paste(resized_invoice_image, (MARGIN, MARGIN))
            output_pdf_path = os.path.join(folder_path, 'invoice_page.pdf')
            merged_image.save(output_pdf_path, 'PDF', resolution=300.0)
            doc.insert_pdf(fitz.open(output_pdf_path))
            os.remove(output_pdf_path)
            log_debug("Saved invoice-only page.")

            # Second page: collage only, centered
            collage_page = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))
            collage_x_offset = (A4_WIDTH - collage_image.width) // 2
            collage_y_offset = (A4_HEIGHT - collage_image.height) // 2
            collage_page.paste(collage_image, (collage_x_offset, collage_y_offset))
            collage_pdf_path = os.path.join(folder_path, 'collage_page.pdf')
            collage_page.save(collage_pdf_path, 'PDF', resolution=300.0)
            doc.insert_pdf(fitz.open(collage_pdf_path))
            os.remove(collage_pdf_path)
            log_debug("Saved collage-only page.")
        else:
            # Single page: invoice and collage combined
            merged_image = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))
            merged_image.paste(resized_invoice_image, (MARGIN, MARGIN))
            collage_y_offset = MARGIN + new_height + (remaining_space - collage_image.height) // 2
            merged_image.paste(collage_image, (MARGIN, collage_y_offset))
            output_pdf_path = os.path.join(folder_path, 'combined_page.pdf')
            merged_image.save(output_pdf_path, 'PDF', resolution=300.0)
            doc.insert_pdf(fitz.open(output_pdf_path))
            os.remove(output_pdf_path)
            log_debug("Saved combined invoice and collage page.")

        folder_count += 1
        success_folders.append(folder_path)
        return 1

    except Exception as e:
        ignored_folders.append((folder_path, str(e)))
        log_debug(f"Error processing folder {folder_path}: {e}")
        return 0

def process_all_subfolders_to_total_pdf(base_folder):
    global folder_count
    doc = fitz.open()

    subfolders = [os.path.join(base_folder, subfolder) for subfolder in os.listdir(base_folder) if os.path.isdir(os.path.join(base_folder, subfolder))]

    for subfolder_path in tqdm(subfolders, desc="正在处理文件夹"):
        merge_invoice_and_images_to_total_pdf(subfolder_path, doc)

    if folder_count > 0:
        timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        output_pdf = f'./报销单_自动生成_{folder_count}张发票_{timestamp}.pdf'
        doc.save(output_pdf)
        doc.close()
        print(f"成功创建 {output_pdf}")
    else:
        print("没有找到有效的文件夹进行处理。")

    # 输出成功与忽略的文件夹
    print("\n处理结果:")
    if success_folders:
        print("成功整理的文件夹:")
        for folder in success_folders:
            print(f"- {folder}")
    if ignored_folders:
        print("被忽略的文件夹及原因:")
        for folder, reason in ignored_folders:
            print(f"- {folder}：{reason}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='合并发票和图片为 PDF')
    parser.add_argument('--debug', action='store_true', help='启用调试信息')
    args = parser.parse_args()
    debug_mode = args.debug

    base_folder = './'
    process_all_subfolders_to_total_pdf(base_folder)