import os
from PIL import Image, ImageTk
import PyPDF2
import img2pdf
import io
import tkinter as tk
from tkinter import Toplevel, Text, Scrollbar

# 定義路徑
base_dir = os.path.dirname(os.path.abspath(__file__))
source_folder = os.path.join(base_dir, "原始檔")
stamp_path = os.path.join(base_dir, "樣章.png")
target_folder = os.path.join(base_dir, "批次蓋章完成")

# 日誌函數：將訊息打印到控制台
def log_message(message):
    print(f"[LOG] {message}")

# 自訂錯誤訊息視窗（不自動關閉）
def show_error_message(title, message):
    log_message(f"錯誤訊息：{message}")
    error_window = Toplevel()
    error_window.title(title)
    error_window.geometry("500x300")
    
    text_area = Text(error_window, wrap="word")
    scrollbar = Scrollbar(error_window, orient="vertical", command=text_area.yview)
    text_area.configure(yscrollcommand=scrollbar.set)
    
    text_area.insert("end", message)
    text_area.config(state="disabled")
    
    text_area.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    error_window.protocol("WM_DELETE_WINDOW", lambda: error_window.destroy())
    error_window.mainloop()

# 檢查目標資料夾是否存在並有寫入權限
log_message(f"檢查目標資料夾：{target_folder}")
try:
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
    test_file = os.path.join(target_folder, "test.txt")
    with open(test_file, "w") as f:
        f.write("test")
    os.remove(test_file)
    log_message("目標資料夾可寫入")
except Exception as e:
    log_message(f"目標資料夾無法寫入：{e}")
    show_error_message("錯誤", f"目標資料夾無法寫入：{e}")
    raise

# 載入電子章並轉為 RGBA 模式
log_message(f"載入電子章：{stamp_path}")
try:
    stamp = Image.open(stamp_path).convert("RGBA")
    stamp_width, stamp_height = stamp.size
    log_message("電子章載入成功")
except Exception as e:
    log_message(f"無法載入電子章圖片：{e}")
    show_error_message("錯誤", f"無法載入電子章圖片：{e}")
    raise

# A4 紙尺寸（單位：毫米）
A4_WIDTH_MM = 210
A4_HEIGHT_MM = 297
POINTS_PER_MM = 72 / 25.4
A4_WIDTH_POINTS = A4_WIDTH_MM * POINTS_PER_MM
A4_HEIGHT_POINTS = A4_HEIGHT_MM * POINTS_PER_MM

# 目標電子章寬度：50 毫米
TARGET_STAMP_WIDTH_MM = 50
TARGET_STAMP_WIDTH_POINTS = TARGET_STAMP_WIDTH_MM * POINTS_PER_MM

# 支援的檔案格式
supported_image_formats = (".jpg", ".jpeg", ".png")
supported_pdf_formats = (".pdf",)

# 儲存拖曳後的偏移量（單位：點）
offset_x_points = 60  # 距離右邊的偏移量（點）
offset_y_points = 50  # 距離底部的偏移量（點）

# GUI 拖曳介面
def create_drag_window():
    global offset_x_points, offset_y_points
    log_message("開始創建拖曳視窗")

    # 尋找第一個支援的檔案作為範例
    sample_file = None
    for filename in os.listdir(source_folder):
        if filename.lower().endswith(supported_image_formats + supported_pdf_formats):
            sample_file = os.path.join(source_folder, filename)
            break
    if not sample_file:
        log_message("來源資料夾中沒有支援的檔案")
        show_error_message("錯誤", "來源資料夾中沒有支援的檔案！")
        return None

    log_message(f"範例檔案：{sample_file}")

    # 載入範例圖片
    if sample_file.lower().endswith(supported_image_formats):
        try:
            sample_image = Image.open(sample_file).convert("RGBA")
            page_width = A4_WIDTH_POINTS
            page_height = A4_HEIGHT_POINTS
            log_message("圖片載入成功")
        except Exception as e:
            log_message(f"無法載入圖片：{e}")
            show_error_message("錯誤", f"無法載入圖片：{e}")
            return None
    else:  # PDF
        try:
            images = convert_from_path(sample_file, dpi=72, first_page=1, last_page=1)
            sample_image = images[0].convert("RGBA")
            log_message("PDF 轉圖片成功")
        except Exception as e:
            log_message(f"無法轉換 PDF 為圖片：{e}")
            show_error_message("錯誤", f"無法轉換 PDF 為圖片：{e}\n請確保已安裝 poppler 並加入環境變數。")
            return None
        pdf_reader = PyPDF2.PdfReader(sample_file)
        page = pdf_reader.pages[0]
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)
        log_message(f"PDF 頁面尺寸：{page_width}x{page_height}")

    # 縮放圖片以適應視窗（假設視窗最大 800x600）
    sample_image.thumbnail((800, 600), Image.Resampling.LANCZOS)
    sample_width, sample_height = sample_image.size
    log_message(f"預覽圖片尺寸：{sample_width}x{sample_height}")

    # 計算縮放比例
    scale_x = sample_width / page_width
    scale_y = sample_height / page_height
    log_message(f"縮放比例：X={scale_x}, Y={scale_y}")

    # 調整電子章大小
    stamp_width_pixels = TARGET_STAMP_WIDTH_POINTS * scale_x
    stamp_scale = stamp_width_pixels / stamp_width
    new_stamp_width = int(stamp_width * stamp_scale)
    new_stamp_height = int(stamp_height * stamp_scale)
    resized_stamp = stamp.resize((new_stamp_width, new_stamp_height), Image.Resampling.LANCZOS)
    log_message(f"電子章縮放後尺寸：{new_stamp_width}x{new_stamp_height}")

    # 創建 Tkinter 視窗
    root = tk.Tk()
    root.title("拖曳電子章位置")
    canvas = tk.Canvas(root, width=sample_width, height=sample_height)
    canvas.pack()

    # 顯示範例圖片
    sample_photo = ImageTk.PhotoImage(sample_image)
    canvas.create_image(0, 0, anchor="nw", image=sample_photo)

    # 顯示電子章（初始位置：右下角）
    stamp_photo = ImageTk.PhotoImage(resized_stamp)
    stamp_id = canvas.create_image(sample_width - new_stamp_width - (offset_x_points * scale_x), 
                                  sample_height - new_stamp_height - (offset_y_points * scale_y), 
                                  anchor="nw", image=stamp_photo)

    # 添加「完成」按鈕
    def on_finish():
        stamp_x = canvas.coords(stamp_id)[0]
        stamp_y = canvas.coords(stamp_id)[1]
        offset_x_pixels = sample_width - stamp_x - new_stamp_width
        offset_y_pixels = sample_height - stamp_y - new_stamp_height
        global offset_x_points, offset_y_points
        offset_x_points = offset_x_pixels / scale_x
        offset_y_points = offset_y_pixels / scale_y
        log_message(f"拖曳完成，偏移量：X={offset_x_points}, Y={offset_y_points}")
        root.destroy()

    finish_button = tk.Button(root, text="完成拖曳", command=on_finish)
    finish_button.pack()

    # 拖曳邏輯
    def start_drag(event):
        canvas.data = {"x": event.x, "y": event.y}

    def dragging(event):
        dx = event.x - canvas.data["x"]
        dy = event.y - canvas.data["y"]
        canvas.move(stamp_id, dx, dy)
        canvas.data["x"] = event.x
        canvas.data["y"] = event.y

    canvas.bind("<Button-1>", start_drag)
    canvas.bind("<B1-Motion>", dragging)

    root.mainloop()
    return offset_x_points, offset_y_points

# 處理圖片檔案的函數
def process_image(image_path, output_path):
    log_message(f"開始處理圖片：{image_path}")
    try:
        image = Image.open(image_path).convert("RGBA")
        image_width, image_height = image.size

        try:
            dpi = image.info.get("dpi", (96, 96))[0]
        except:
            dpi = 96

        image_width_mm = (image_width / dpi) * 25.4
        scale_factor = A4_WIDTH_MM / image_width_mm
        stamp_width_mm = TARGET_STAMP_WIDTH_MM / scale_factor
        stamp_width_pixels = (stamp_width_mm / 25.4) * dpi

        stamp_scale = stamp_width_pixels / stamp_width
        new_stamp_width = int(stamp_width * stamp_scale)
        new_stamp_height = int(stamp_height * stamp_scale)
        resized_stamp = stamp.resize((new_stamp_width, new_stamp_height), Image.Resampling.LANCZOS)
        log_message(f"圖片電子章調整大小：{new_stamp_width}x{new_stamp_height}")

        # 使用拖曳後的偏移量（轉為圖片像素）
        scale_x = image_width / A4_WIDTH_POINTS
        scale_y = image_height / A4_HEIGHT_POINTS
        position_x = image_width - new_stamp_width - (offset_x_points * scale_x)
        position_y = image_height - new_stamp_height - (offset_y_points * scale_y)
        # 將浮點數轉為整數
        position_x = int(position_x)
        position_y = int(position_y)
        position = (position_x, position_y)
        log_message(f"圖片電子章位置：X={position_x}, Y={position_y}")

        image.paste(resized_stamp, position, resized_stamp)

        if image_path.lower().endswith((".jpg", ".jpeg")):
            image = image.convert("RGB")

        image.save(output_path)
        log_message(f"已處理並儲存圖片：{output_path}")

    except Exception as e:
        log_message(f"處理圖片時出錯：{e}")
        show_error_message("錯誤", f"處理圖片時出錯：{e}")
        raise

# 處理 PDF 檔案的函數
def process_pdf(pdf_path, output_path):
    log_message(f"開始處理 PDF：{pdf_path}")
    try:
        pdf_dpi = 72
        stamp_width_pixels = TARGET_STAMP_WIDTH_POINTS * (pdf_dpi / 72)
        stamp_scale = stamp_width_pixels / stamp_width
        new_stamp_width = int(stamp_width * stamp_scale)
        new_stamp_height = int(stamp_height * stamp_scale)
        resized_stamp = stamp.resize((new_stamp_width, new_stamp_height), Image.Resampling.LANCZOS)
        log_message(f"PDF 電子章調整大小：{new_stamp_width}x{new_stamp_height}")

        if resized_stamp.mode == "RGBA":
            resized_stamp = resized_stamp.convert("RGB")

        temp_stamp_path = os.path.join(target_folder, "resized_stamp.png")
        resized_stamp.save(temp_stamp_path, "PNG")
        log_message(f"已生成臨時電子章圖片：{temp_stamp_path}")

        stamp_pdf_path = os.path.join(target_folder, "stamp_temp.pdf")
        with open(stamp_pdf_path, "wb") as f:
            f.write(img2pdf.convert(temp_stamp_path, dpi=pdf_dpi))
        log_message(f"已將電子章轉為 PDF：{stamp_pdf_path}")

        pdf_reader = PyPDF2.PdfReader(pdf_path)
        stamp_pdf_reader = PyPDF2.PdfReader(stamp_pdf_path)
        stamp_page = stamp_pdf_reader.pages[0]
        log_message("PDF 檔案載入成功")

        pdf_writer = PyPDF2.PdfWriter()

        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            page_width = float(page.mediabox.width)
            page_height = float(page.mediabox.height)

            stamp_width_points = TARGET_STAMP_WIDTH_POINTS
            stamp_height_points = stamp_height * (stamp_width_points / stamp_width)
            log_message(f"PDF 電子章最終尺寸：{stamp_width_points}x{stamp_height_points}")

            position_x = page_width - stamp_width_points - offset_x_points
            position_y = offset_y_points
            log_message(f"PDF 電子章位置：X={position_x}, Y={position_y}")

            page.merge_translated_page(stamp_page, position_x, position_y)
            pdf_writer.add_page(page)

        with open(output_path, "wb") as f:
            pdf_writer.write(f)
        log_message(f"已處理並儲存 PDF：{output_path}")

        os.remove(temp_stamp_path)
        os.remove(stamp_pdf_path)
        log_message("已清理臨時檔案")

    except Exception as e:
        log_message(f"處理 PDF 時出錯：{e}")
        show_error_message("錯誤", f"處理 PDF 時出錯：{e}")
        raise

# 主程式
log_message("程式開始執行")
try:
    result = create_drag_window()
    if result is None:
        log_message("拖曳視窗未返回有效偏移量，程式中止")
        show_error_message("錯誤", "拖曳視窗未返回有效偏移量，程式中止")
        raise Exception("拖曳視窗未正確執行")
    offset_x_points, offset_y_points = result
    log_message(f"拖曳完成，偏移量：X={offset_x_points}, Y={offset_y_points}")
except Exception as e:
    log_message(f"拖曳視窗出錯：{e}")
    show_error_message("錯誤", f"拖曳視窗出錯：{e}")
    raise

# 遍歷來源資料夾中的檔案
for filename in os.listdir(source_folder):
    file_path = os.path.join(source_folder, filename)
    output_path = os.path.join(target_folder, filename)

    try:
        if filename.lower().endswith(supported_image_formats):
            log_message(f"開始處理圖片檔案：{file_path}")
            process_image(file_path, output_path)
        elif filename.lower().endswith(supported_pdf_formats):
            log_message(f"開始處理 PDF 檔案：{file_path}")
            process_pdf(file_path, output_path)
        else:
            log_message(f"跳過不支援的檔案：{file_path}")
    except Exception as e:
        log_message(f"處理檔案 {file_path} 時出錯：{e}")
        show_error_message("錯誤", f"處理檔案 {file_path} 時出錯：{e}")
        continue  # 繼續處理下一個檔案

log_message("批次處理完成！")
