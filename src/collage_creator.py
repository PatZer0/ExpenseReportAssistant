from PIL import Image
from file_utils import log_debug

def create_collage_image(image_files, max_width, cell_height, debug_mode=False):
    """
    创建图片拼贴
    
    Args:
        image_files: 图片文件路径列表
        max_width: 拼贴最大宽度
        cell_height: 单元格高度
        debug_mode: 是否启用调试模式
    
    Returns:
        拼贴好的PIL图像对象，如果没有图片则返回None
    """
    log_debug(f'创建拼图 {image_files} (最大宽度 {max_width}, 单元高度 {cell_height})', debug_mode)
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
        return _create_row_collage(images, max_width, cell_height)
    else:
        return _create_grid_collage(images, max_width, cell_height)

def _create_row_collage(images, max_width, cell_height):
    """创建按行排列的拼贴图像"""
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
        
    return collage_image

def _create_grid_collage(images, max_width, cell_height):
    """创建网格排列的拼贴图像"""
    num_images = len(images)
    images_per_row = 4
    rows_needed = (num_images + images_per_row - 1) // images_per_row
    target_height_per_image = cell_height // rows_needed
    
    # 缩放图像以适应网格
    scaled_images = []
    for img in images:
        scale_factor = target_height_per_image / img.height
        new_width = int(img.width * scale_factor)
        new_height = target_height_per_image
        
        # 确保每列宽度不超过最大宽度的1/4
        max_col_width = max_width // images_per_row
        if new_width > max_col_width:
            new_width = max_col_width
            scale_factor = new_width / img.width
            new_height = int(img.height * scale_factor)
            
        scaled_images.append(img.resize((new_width, new_height), Image.LANCZOS))
    
    # 计算拼图的总高度
    collage_height = rows_needed * target_height_per_image
    collage_image = Image.new('RGB', (max_width, collage_height), (255, 255, 255))
    
    # 按照网格排布图像
    for idx, img in enumerate(scaled_images):
        row = idx // images_per_row
        col = idx % images_per_row
        
        # 计算每张图片在网格中的位置
        cell_width = max_width // images_per_row
        x_offset = col * cell_width + (cell_width - img.width) // 2  # 居中放置
        y_offset = row * target_height_per_image + (target_height_per_image - img.height) // 2  # 居中放置
        
        collage_image.paste(img, (x_offset, y_offset))
        
    return collage_image