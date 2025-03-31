import os
import fitz  # PyMuPDF
from PIL import Image
import datetime
from tqdm import tqdm
from prettytable import PrettyTable
from colorama import init, Fore, Style
from file_utils import windows_sort_key, log_debug
from collage_creator import create_collage_image

# 初始化colorama
init()

# 定义 A4 纸尺寸（300 DPI）
A4_WIDTH = 2480  # 8.27 inches * 300 dpi
A4_HEIGHT = 3508  # 11.69 inches * 300 dpi
MARGIN = 118  # 1 cm ≈ 118 pixels at 300 DPI
MIN_SPACE_FOR_COLLAGE = 210  # 7 cm in pixels

CONTENT_WIDTH = A4_WIDTH - 2 * MARGIN  # 2244
CONTENT_HEIGHT = A4_HEIGHT - 2 * MARGIN  # 3272
HEIGHT_THRESHOLD = CONTENT_HEIGHT * 0.7  # 内容区域高度的70%

class PDFMerger:
    def __init__(self, debug_mode=False):
        self.debug_mode = debug_mode
        self.folder_count = 0
        self.success_folders = []
        self.ignored_folders = []

    def create_document(self):
        """创建一个新的PDF文档"""
        return fitz.open()

    def get_timestamp(self):
        """获取当前时间戳"""
        return datetime.datetime.now().strftime("%Y%m%d%H%M%S")

    def merge_invoice_and_images_to_total_pdf(self, folder_path, doc):
        """合并单个文件夹中的发票PDF和图片至总文档"""
        try:
            log_debug(f"\n正在处理文件夹: {folder_path}", self.debug_mode)
            files = os.listdir(folder_path)
            pdf_files = [os.path.join(folder_path, f) for f in files if f.lower().endswith('.pdf')]
            image_files = [os.path.join(folder_path, f) for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg'))]

            # 筛选 NEWPAGE 和 NEWLINE 图片
            newline_images = [f for f in image_files if os.path.basename(f).startswith('NEWLINE')]
            newpage_images = [f for f in image_files if os.path.basename(f).startswith('NEWPAGE')]
            other_images = [f for f in image_files if f not in newline_images and f not in newpage_images]

            if len(pdf_files) != 1 or len(other_images) < 2:
                # 修改存储结构，保存PDF和图片数量信息
                reason = f"仅找到 {len(pdf_files)} 个 PDF 文件与 {len(other_images)} 个图片文件."
                self.ignored_folders.append((folder_path, len(pdf_files), len(other_images), reason))
                log_debug(f"Ignored {folder_path}: {reason}", self.debug_mode)
                return 0

            invoice_pdf_path = pdf_files[0]
            invoice_doc = fitz.open(invoice_pdf_path)
            
            # 检查PDF页数
            pdf_page_count = len(invoice_doc)
            
            if pdf_page_count > 1:
                # 多页PDF，直接整个插入
                log_debug(f"处理多页PDF（{pdf_page_count}页）: {invoice_pdf_path}", self.debug_mode)
                doc.insert_pdf(invoice_doc)
                
                # 为多页PDF创建独立的拼图页
                collage_image = create_collage_image(other_images, CONTENT_WIDTH, CONTENT_HEIGHT, self.debug_mode)
                if collage_image is None:
                    log_debug("没有足够的图片，无法创建拼图.", self.debug_mode)
                else:
                    collage_page = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))
                    collage_x_offset = (A4_WIDTH - collage_image.width) // 2
                    collage_y_offset = (A4_HEIGHT - collage_image.height) // 2
                    collage_page.paste(collage_image, (collage_x_offset, collage_y_offset))
                    collage_pdf_path = os.path.join(folder_path, 'collage_page.pdf')
                    collage_page.save(collage_pdf_path, 'PDF', resolution=300.0)
                    doc.insert_pdf(fitz.open(collage_pdf_path))
                    os.remove(collage_pdf_path)
            else:
                # 单页PDF，按原逻辑处理
                invoice_page = invoice_doc.load_page(0)
                
                scale = 5
                matrix = fitz.Matrix(scale, scale)
                pix = invoice_page.get_pixmap(matrix=matrix)
                invoice_image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        
                scale_factor = CONTENT_WIDTH / invoice_image.width
                new_height = int(invoice_image.height * scale_factor)
                resized_invoice_image = invoice_image.resize((CONTENT_WIDTH, new_height), Image.LANCZOS)
        
                remaining_space = CONTENT_HEIGHT - resized_invoice_image.height
                create_new_page_for_collage = resized_invoice_image.height > HEIGHT_THRESHOLD
        
                collage_image = create_collage_image(
                    other_images, 
                    CONTENT_WIDTH, 
                    remaining_space if not create_new_page_for_collage else CONTENT_HEIGHT, 
                    self.debug_mode
                )
                if collage_image is None:
                    log_debug("没有足够的图片，无法创建拼图.", self.debug_mode)
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
            self._process_special_images(newline_images, folder_path, doc, 'newline_page')

            # 处理 NEWPAGE 图片
            self._process_special_images(newpage_images, folder_path, doc, 'newpage_image')

            self.folder_count += 1
            self.success_folders.append(folder_path)
            return 1

        except Exception as e:
            self.ignored_folders.append((folder_path, str(e)))
            log_debug(f"处理文件夹错误 {folder_path}: {e}", self.debug_mode)
            return 0
            
    def _process_special_images(self, image_paths, folder_path, doc, prefix):
        """处理特殊图片（NEWLINE或NEWPAGE）"""
        for image_path in image_paths:
            img = Image.open(image_path)
            scale_factor = CONTENT_WIDTH / img.width
            resized_image = img.resize((CONTENT_WIDTH, int(img.height * scale_factor)), Image.LANCZOS)

            page = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))
            y_offset = (A4_HEIGHT - resized_image.height) // 2
            page.paste(resized_image, (MARGIN, y_offset))
            pdf_path = os.path.join(folder_path, f'{prefix}.pdf')
            page.save(pdf_path, 'PDF', resolution=300.0)
            doc.insert_pdf(fitz.open(pdf_path))
            os.remove(pdf_path)

    def process_all_subfolders_to_total_pdf(self, base_folder, output_path=''):
        """处理所有子文件夹并合并为一个PDF文件"""
        doc = self.create_document()

        # 获取上级文件夹的名称
        parent_folder_name = os.path.basename(os.path.abspath(base_folder))

        subfolders = [os.path.join(base_folder, subfolder) for subfolder in os.listdir(base_folder) 
                    if os.path.isdir(os.path.join(base_folder, subfolder))]
        # 按照Windows的排序规则（包括中文拼音）对子文件夹进行排序
        subfolders.sort(key=windows_sort_key)

        for subfolder_path in tqdm(subfolders, desc="正在处理文件夹"):
            self.merge_invoice_and_images_to_total_pdf(subfolder_path, doc)

        if self.folder_count > 0:
            timestamp = self.get_timestamp()
            # 使用上级文件夹名称作为文件名前缀
            default_output_filename = f'{parent_folder_name}_报销单_自动生成_{self.folder_count}张发票_{timestamp}.pdf'
            
            output_pdf = self._determine_output_path(output_path, default_output_filename)
            if not output_pdf:
                doc.close()
                return

            doc.save(output_pdf)
            doc.close()
            print(f"成功创建 {output_pdf}")

            # 显示忽略的文件夹信息
            self._display_ignored_folders()

    def _determine_output_path(self, output_path, default_filename):
        """确定输出文件路径"""
        if os.path.isdir(output_path):
            return os.path.join(output_path, default_filename)
        elif output_path and os.path.splitext(output_path)[1].lower() == '.pdf':
            if os.path.exists(output_path):
                user_input = input(f"{output_path} 已存在，是否覆盖？ (y/n): ").strip().lower()
                if user_input != 'y':
                    print("操作已取消。")
                    return None
            return output_path
        else:
            return os.path.join('./', default_filename) if output_path == '' else output_path

    def _display_ignored_folders(self):
        """显示被忽略的文件夹信息"""
        if len(self.ignored_folders) > 0:
            print("\n以下文件夹被忽略：")
            table = PrettyTable()
            table.field_names = ["路径", "PDF", "图片"]
            
            for folder_data in self.ignored_folders:
                if len(folder_data) == 4:  # 包含PDF和图片计数的情况
                    folder_path, pdf_count, img_count, reason = folder_data
                    # PDF状态：需要恰好1个PDF
                    pdf_status = f"{Fore.GREEN}√{Style.RESET_ALL}" if pdf_count == 1 else f"{Fore.RED}X({pdf_count}){Style.RESET_ALL}"
                    
                    # 图片状态：需要至少2张图片
                    img_status = f"{Fore.GREEN}√{Style.RESET_ALL}" if img_count >= 2 else f"{Fore.RED}X({img_count}){Style.RESET_ALL}"
                    
                    table.add_row([folder_path, pdf_status, img_status])
                else:
                    # 处理异常情况
                    folder_path, error = folder_data
                    table.add_row([folder_path, f"{Fore.RED}错误{Style.RESET_ALL}", error])
            
            print(table)
            print(f"\n需求：每个文件夹应有1个PDF文件和至少2张图片文件")

    def rename_pdf_files(self, base_folder):
        """根据上级文件夹名称重命名PDF文件"""
        subfolders = [os.path.join(base_folder, subfolder) for subfolder in os.listdir(base_folder) 
                    if os.path.isdir(os.path.join(base_folder, subfolder))]

        # 按照Windows的排序规则（包括中文拼音）对子文件夹进行排序
        subfolders.sort(key=windows_sort_key)
        
        renamed_count = 0
        skipped_count = 0
        
        for subfolder_path in tqdm(subfolders, desc="正在重命名文件"):
            folder_name = os.path.basename(subfolder_path)
            pdf_files = [f for f in os.listdir(subfolder_path) if f.lower().endswith('.pdf')]
            
            if len(pdf_files) == 1:
                old_path = os.path.join(subfolder_path, pdf_files[0])
                new_name = f"{folder_name}.pdf"
                new_path = os.path.join(subfolder_path, new_name)
                
                try:
                    os.rename(old_path, new_path)
                    renamed_count += 1
                    log_debug(f"已将 {old_path} 重命名为 {new_path}", self.debug_mode)
                except Exception as e:
                    log_debug(f"重命名失败 {old_path}: {e}", self.debug_mode)
                    skipped_count += 1
            else:
                log_debug(f"跳过 {subfolder_path}: 找到 {len(pdf_files)} 个PDF文件(需要恰好1个)", self.debug_mode)
                skipped_count += 1
        
        print(f"\n完成! 已重命名 {renamed_count} 个文件, 跳过 {skipped_count} 个文件夹")