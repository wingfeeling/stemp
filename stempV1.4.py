import os
import sys
import traceback
from tkinter import *
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import PyPDF2
from pdf2image import convert_from_path
import logging
from logging.handlers import RotatingFileHandler

# 常量定义
A4_WIDTH_POINTS = 595.0  # A4宽度（点）
A4_HEIGHT_POINTS = 842.0  # A4高度（点）
TARGET_STAMP_WIDTH_POINTS = 100.0  # 目标电子章宽度（点）
supported_pdf_formats = ('.pdf',)
supported_image_formats = ('.png', '.jpg', '.jpeg', '.bmp', '.gif')

# 初始化全局变量
selected_files = []
stamp = None
stamp_width = 0
stamp_height = 0
offset_x = 0
offset_y = 0
offset_x_points = 0
offset_y_points = 0
drag_window = None
preview_image = None
preview_canvas = None
stamp_image_on_canvas = None

# 设置日志
def setup_logging():
    log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    log_file = "pdf_stamp.log"
    
    handler = RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=2)
    handler.setFormatter(log_formatter)
    
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    logger.addHandler(handler)

def log_message(message, level="info"):
    if level == "debug":
        logging.debug(message)
    elif level == "warning":
        logging.warning(message)
    elif level == "error":
        logging.error(message)
    else:
        logging.info(message)

def show_error_message(title, message):
    log_message(f"{title}: {message}", level="error")
    messagebox.showerror(title, message)

def show_info_message(title, message):
    log_message(f"{title}: {message}", level="info")
    messagebox.showinfo(title, message)

def select_files():
    global selected_files
    files = filedialog.askopenfilenames(
        title="选择PDF或图片文件",
        filetypes=[("PDF文件", "*.pdf"), ("图片文件", "*.png *.jpg *.jpeg *.bmp *.gif"), ("所有文件", "*.*")]
    )
    if files:
        selected_files = list(files)
        log_message(f"已选择文件: {selected_files}", level="debug")
        show_info_message("选择成功", f"已选择 {len(selected_files)} 个文件")
    else:
        log_message("未选择任何文件", level="warning")

def select_stamp():
    global stamp, stamp_width, stamp_height
    file = filedialog.askopenfilename(
        title="选择电子章图片",
        filetypes=[("图片文件", "*.png *.jpg *.jpeg *.bmp *.gif")]
    )
    if file:
        try:
            stamp = Image.open(file).convert("RGBA")
            stamp_width, stamp_height = stamp.size
            log_message(f"已选择电子章: {file}, 尺寸: {stamp_width}x{stamp_height}", level="debug")
            show_info_message("选择成功", "电子章已载入")
        except Exception as e:
            log_message(f"载入电子章失败: {traceback.format_exc()}", level="error")
            show_error_message("错误", f"无法载入电子章: {e}")

def create_drag_window():
    global offset_x_points, offset_y_points, drag_window, preview_image, preview_canvas, stamp_image_on_canvas
    log_message("开始创建拖曳视窗", level="debug")
    
    if not selected_files:
        show_error_message("错误", "请先选择PDF文件")
        return None
    
    if not stamp:
        show_error_message("错误", "请先选择电子章")
        return None
    
    try:
        sample_file = selected_files[0]
        log_message(f"范例档案：{sample_file}", level="debug")

        # 取得PDF的实际尺寸（以点为单位）
        if sample_file.lower().endswith(supported_pdf_formats):
            try:
                pdf_reader = PyPDF2.PdfReader(sample_file)
                page = pdf_reader.pages[0]
                pdf_width_points = float(page.mediabox.width)
                pdf_height_points = float(page.mediabox.height)
                
                # 标准A4尺寸（以点为单位）
                standard_a4_width = A4_WIDTH_POINTS
                standard_a4_height = A4_HEIGHT_POINTS
                
                # 计算比例因子
                scale_x = pdf_width_points / standard_a4_width
                scale_y = pdf_height_points / standard_a4_height
                scale_factor = (scale_x + scale_y) / 2
                log_message(f"PDF比例因子：{scale_factor}", level="debug")
                
                # 调整电子章大小使其在物理尺寸上保持一致
                target_stamp_width_points = TARGET_STAMP_WIDTH_POINTS * scale_factor
                stamp_scale = target_stamp_width_points / stamp_width
                new_stamp_width = int(stamp_width * stamp_scale)
                new_stamp_height = int(stamp_height * stamp_scale)
                
                # 载入PDF页面为图片，保持原比例
                images = convert_from_path(sample_file, dpi=72, first_page=1, last_page=1)
                sample_image = images[0].convert("RGBA")
                
            except Exception as e:
                log_message(f"PDF处理错误：{traceback.format_exc()}", level="error")
                show_error_message("错误", f"PDF处理错误：{e}")
                return None
        else:
            # 处理图片档案的部分
            try:
                sample_image = Image.open(sample_file).convert("RGBA")
                pdf_width_points = A4_WIDTH_POINTS
                pdf_height_points = A4_HEIGHT_POINTS
                scale_factor = 1.0
                
                # 调整电子章大小
                target_stamp_width_points = TARGET_STAMP_WIDTH_POINTS
                stamp_scale = target_stamp_width_points / stamp_width
                new_stamp_width = int(stamp_width * stamp_scale)
                new_stamp_height = int(stamp_height * stamp_scale)
                
            except Exception as e:
                log_message(f"无法载入图片：{traceback.format_exc()}", level="error")
                show_error_message("错误", f"无法载入图片：{e}")
                return None

        # 调整电子章大小
        resized_stamp = stamp.resize((new_stamp_width, new_stamp_height), Image.Resampling.LANCZOS)
        
        # 创建预览视窗
        try:
            sample_image.thumbnail((800, 600), Image.Resampling.LANCZOS)
            sample_width, sample_height = sample_image.size
            
            drag_window = Toplevel()
            drag_window.title("电子章位置调整")
            drag_window.geometry(f"{sample_width+20}x{sample_height+20}")
            
            preview_image = ImageTk.PhotoImage(sample_image)
            preview_canvas = Canvas(drag_window, width=sample_width, height=sample_height)
            preview_canvas.create_image(0, 0, anchor=NW, image=preview_image)
            preview_canvas.pack()
            
            # 计算电子章在预览图上的比例
            preview_scale_x = sample_width / pdf_width_points
            preview_scale_y = sample_height / pdf_height_points
            preview_scale = (preview_scale_x + preview_scale_y) / 2
            
            # 放置电子章
            stamp_preview_width = int(new_stamp_width * preview_scale)
            stamp_preview_height = int(new_stamp_height * preview_scale)
            resized_stamp_preview = resized_stamp.resize((stamp_preview_width, stamp_preview_height), Image.Resampling.LANCZOS)
            stamp_tk_image = ImageTk.PhotoImage(resized_stamp_preview)
            
            # 初始位置设为中心
            initial_x = (sample_width - stamp_preview_width) // 2
            initial_y = (sample_height - stamp_preview_height) // 2
            
            stamp_image_on_canvas = preview_canvas.create_image(
                initial_x, initial_y, 
                anchor=NW, 
                image=stamp_tk_image,
                tags="stamp"
            )
            
            # 计算初始偏移量（点）
            offset_x_points = (initial_x / preview_scale)
            offset_y_points = (initial_y / preview_scale)
            
            # 绑定拖曳事件
            preview_canvas.tag_bind("stamp", "<B1-Motion>", drag_stamp)
            
            # 保持图片引用
            drag_window.stamp_tk_image = stamp_tk_image
            drag_window.preview_image = preview_image
            
        except Exception as e:
            log_message(f"创建拖曳视窗失败：{traceback.format_exc()}", level="error")
            show_error_message("错误", f"创建拖曳视窗失败：{e}")
            return None

    except Exception as e:
        log_message(f"拖曳视窗处理失败：{traceback.format_exc()}", level="error")
        show_error_message("错误", f"拖曳视窗处理失败：{e}")
        return None

def drag_stamp(event):
    global offset_x_points, offset_y_points
    if not drag_window or not preview_canvas:
        return
    
    try:
        # 获取当前电子章位置
        coords = preview_canvas.coords(stamp_image_on_canvas)
        if not coords:
            return
            
        # 计算新位置
        new_x = event.x
        new_y = event.y
        
        # 更新电子章位置
        preview_canvas.coords(stamp_image_on_canvas, new_x, new_y)
        
        # 获取PDF页面尺寸
        sample_file = selected_files[0]
        if sample_file.lower().endswith(supported_pdf_formats):
            pdf_reader = PyPDF2.PdfReader(sample_file)
            page = pdf_reader.pages[0]
            pdf_width_points = float(page.mediabox.width)
            pdf_height_points = float(page.mediabox.height)
        else:
            pdf_width_points = A4_WIDTH_POINTS
            pdf_height_points = A4_HEIGHT_POINTS
        
        # 计算预览图比例
        preview_width = preview_canvas.winfo_width()
        preview_height = preview_canvas.winfo_height()
        preview_scale_x = preview_width / pdf_width_points
        preview_scale_y = preview_height / pdf_height_points
        preview_scale = (preview_scale_x + preview_scale_y) / 2
        
        # 更新偏移量（点）
        offset_x_points = new_x / preview_scale
        offset_y_points = new_y / preview_scale
        
    except Exception as e:
        log_message(f"拖曳电子章失败：{traceback.format_exc()}", level="error")
        show_error_message("错误", f"拖曳电子章失败：{e}")

def process_pdf(pdf_path, output_path):
    log_message(f"开始处理 PDF：{pdf_path}", level="debug")
    try:
        pdf_reader = PyPDF2.PdfReader(pdf_path)
        page = pdf_reader.pages[0]
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)
        
        # 标准A4尺寸
        standard_a4_width = A4_WIDTH_POINTS
        standard_a4_height = A4_HEIGHT_POINTS
        
        # 计算比例因子
        scale_x = page_width / standard_a4_width
        scale_y = page_height / standard_a4_height
        scale_factor = (scale_x + scale_y) / 2
        
        # 调整电子章大小使其在物理尺寸上保持一致
        target_stamp_width_points = TARGET_STAMP_WIDTH_POINTS * scale_factor
        stamp_scale = target_stamp_width_points / stamp_width
        new_stamp_width = int(stamp_width * stamp_scale)
        new_stamp_height = int(stamp_height * stamp_scale)
        
        # 调整电子章图片大小
        resized_stamp = stamp.resize((new_stamp_width, new_stamp_height), Image.Resampling.LANCZOS)
        
        # 创建PDF页面的图片
        images = convert_from_path(pdf_path, dpi=300, first_page=1, last_page=1)
        pdf_image = images[0].convert("RGBA")
        
        # 计算电子章位置
        stamp_x = int(offset_x_points * (page_width / A4_WIDTH_POINTS))
        stamp_y = int(offset_y_points * (page_height / A4_HEIGHT_POINTS))
        
        # 将电子章盖在PDF图片上
        pdf_image.paste(resized_stamp, (stamp_x, stamp_y), resized_stamp)
        
        # 保存为临时图片
        temp_image_path = "temp_processed.png"
        pdf_image.save(temp_image_path, "PNG")
        
        # 将图片转换回PDF
        pdf_writer = PyPDF2.PdfWriter()
        pdf_writer.add_page(page)
        
        with open(output_path, "wb") as output_file:
            pdf_writer.write(output_file)
        
        # 删除临时文件
        if os.path.exists(temp_image_path):
            os.remove(temp_image_path)
            
        log_message(f"PDF处理完成：{output_path}", level="info")
        
    except Exception as e:
        log_message(f"处理 PDF 时出错：{traceback.format_exc()}", level="error")
        show_error_message("错误", f"处理 PDF 时出错：{e}")
        raise

def process_files():
    if not selected_files:
        show_error_message("错误", "请先选择PDF文件")
        return
    
    if not stamp:
        show_error_message("错误", "请先选择电子章")
        return
    
    if not drag_window:
        show_error_message("错误", "请先调整电子章位置")
        return
    
    output_dir = filedialog.askdirectory(title="选择输出目录")
    if not output_dir:
        return
    
    progress = ttk.Progressbar(root, orient=HORIZONTAL, length=300, mode='determinate')
    progress.pack(pady=10)
    progress['maximum'] = len(selected_files)
    
    processed_count = 0
    for i, file_path in enumerate(selected_files):
        try:
            file_name = os.path.basename(file_path)
            output_path = os.path.join(output_dir, f"stamped_{file_name}")
            
            if file_path.lower().endswith(supported_pdf_formats):
                process_pdf(file_path, output_path)
            else:
                # 处理图片文件（略）
                pass
            
            processed_count += 1
            progress['value'] = i + 1
            root.update_idletasks()
            
        except Exception as e:
            log_message(f"处理文件 {file_path} 失败：{traceback.format_exc()}", level="error")
            show_error_message("错误", f"处理文件 {file_path} 失败：{e}")
            continue
    
    progress.destroy()
    show_info_message("完成", f"已成功处理 {processed_count}/{len(selected_files)} 个文件")

# 主视窗
def create_main_window():
    global root
    root = Tk()
    root.title("PDF电子章工具")
    root.geometry("400x300")
    
    # 设置日志
    setup_logging()
    
    # 创建UI元素
    frame = Frame(root)
    frame.pack(pady=20)
    
    btn_select_files = Button(frame, text="选择PDF文件", command=select_files)
    btn_select_files.pack(side=LEFT, padx=10)
    
    btn_select_stamp = Button(frame, text="选择电子章", command=select_stamp)
    btn_select_stamp.pack(side=LEFT, padx=10)
    
    btn_adjust_position = Button(root, text="调整电子章位置", command=create_drag_window)
    btn_adjust_position.pack(pady=10)
    
    btn_process = Button(root, text="处理文件", command=process_files)
    btn_process.pack(pady=10)
    
    root.mainloop()

if __name__ == "__main__":
    create_main_window()
