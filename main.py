import os
import random
import platform
import tkinter as tk
from tkinter import ttk
from barcode import Code128
from barcode.writer import ImageWriter
from fpdf import FPDF
from PIL import Image

# ========== 默认寄件人信息 ==========
DEFAULT_SENDER_NAME = "张三"
DEFAULT_SENDER_PHONE = "13800001111"
DEFAULT_SENDER_ADDRESS = "北京市朝阳区XX路88号"

# ========== 字体配置 ==========
FONT_NAME = "CNFont"
FONT_SIZE = 8
ZIP_FONT_SIZE = 16

# ========== 排版设置 ==========
LINE_H = 4
PAGE_SIZE = (105, 148)  # 横向 A6：高105mm，宽148mm

def get_system_chinese_font_path():
    system = platform.system()
    possible_paths = []
    if system == "Darwin":
        possible_paths = [
            "/Library/Fonts/NotoSansCJKsc-Regular.otf",
            "/System/Library/Fonts/STHeiti Medium.ttc",
            "/System/Library/Fonts/PingFang.ttc",
        ]
    elif system == "Windows":
        possible_paths = [
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/simhei.ttf",
        ]
    elif system == "Linux":
        possible_paths = [
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
            "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.otf"
        ]
    for path in possible_paths:
        if os.path.exists(path):
            return path
    return None

CHINESE_FONT_PATH = get_system_chinese_font_path()

# ========== 条码 ==========
def generate_barcode_string():
    # 生成前10位数字
    digits = [random.randint(0, 9) for _ in range(10)]
    full_number = ''.join(map(str, digits)) + str(calculate_mod11_check_digit(digits))
    return f"XE{full_number}"

def calculate_mod11_check_digit(digits):
    weights = list(range(2, 12))  # 2~11 共10个
    total = sum(d * w for d, w in zip(reversed(digits), weights))
    remainder = total % 11
    check_digit = (11 - remainder) % 10  # 避免出现10，统一取余数
    return check_digit


def generate_barcode_image(barcode_value):
    filename = "barcode_temp"
    barcode_obj = Code128(barcode_value, writer=ImageWriter())
    full_path = barcode_obj.save(filename)
    return full_path

# ========== PDF ==========
def generate_pdf(recipient, sender, barcode_value):
    pdf = FPDF(orientation="L", unit="mm", format=PAGE_SIZE)
    pdf.add_page()

    # 设置正文字体（小号）
    try:
        if CHINESE_FONT_PATH:
            pdf.add_font(FONT_NAME, "", CHINESE_FONT_PATH)
            pdf.set_font(FONT_NAME, size=FONT_SIZE)
        else:
            raise FileNotFoundError("系统未找到合适的中文字体路径")
    except Exception as e:
        print(f"⚠️ 中文字体加载失败，使用默认字体：{e}")
        pdf.set_font("Helvetica", size=FONT_SIZE)

    # 收件人信息（左中）
    recipient_text = (
        f"收件人：{recipient['name']}\n"
        f"{recipient['phone']}\n"
        f"{recipient['address']}"
    )
    pdf.set_xy(8, 22)
    pdf.multi_cell(w=70, h=LINE_H, text=recipient_text)

    # 寄件人信息（右下）
    sender_text = (
        f"寄件人：{sender['name']}\n"
        f"{sender['phone']}\n"
        f"{sender['address']}"
    )
    pdf.set_xy(86, 72)
    pdf.multi_cell(w=50, h=LINE_H, text=sender_text)

    # 条码图像（左下，缩小一半，保持比例）
    barcode_path = generate_barcode_image(barcode_value)
    try:
        with Image.open(barcode_path) as img:
            orig_width, orig_height = img.size
            dpi = 300
            width_mm = orig_width / dpi * 25.4 * 0.5
            height_mm = orig_height / dpi * 25.4 * 0.5
        pdf.image(barcode_path, x=8, y=74, w=width_mm, h=height_mm)
    except Exception as e:
        print("⚠️ 插入条码失败：", e)

    # 最后插入邮编（左上角），使用独立字体大小
    if recipient.get("zip"):
        try:
            if CHINESE_FONT_PATH:
                pdf.set_font(FONT_NAME, size=ZIP_FONT_SIZE)
            else:
                raise FileNotFoundError()
        except:
            pdf.set_font("Helvetica", size=ZIP_FONT_SIZE)
        pdf.set_xy(8, 8)
        pdf.cell(w=30, h=LINE_H, text=recipient['zip'])

    pdf.output("delivery_info.pdf")
    print("✅ PDF 已生成：delivery_info.pdf")

    if os.path.exists(barcode_path):
        os.remove(barcode_path)

# ========== GUI ==========
def on_generate_barcode():
    code = generate_barcode_string()
    barcode_entry.delete(0, tk.END)
    barcode_entry.insert(0, code)

def on_generate_pdf():
    recipient = {
        "name": recipient_name.get(),
        "phone": recipient_phone.get(),
        "address": recipient_address.get(),
        "zip": recipient_zip.get()
    }
    sender = {
        "name": sender_name.get(),
        "phone": sender_phone.get(),
        "address": sender_address.get()
    }
    barcode_value = barcode_entry.get().strip()
    if not barcode_value:
        barcode_value = generate_barcode_string()
        barcode_entry.insert(0, barcode_value)

    generate_pdf(recipient, sender, barcode_value)

# ========== UI 布局 ==========
root = tk.Tk()
root.title("邮递标签打印系统（China Post）")

frame = ttk.Frame(root, padding=12)
frame.grid(row=0, column=0)

# 收件人输入区域
left_frame = ttk.LabelFrame(frame, text="收件人信息", padding=8)
left_frame.grid(row=0, column=0, padx=10)

ttk.Label(left_frame, text="姓名").grid(row=0, column=0, sticky="w")
recipient_name = ttk.Entry(left_frame, width=25)
recipient_name.grid(row=0, column=1)

ttk.Label(left_frame, text="电话").grid(row=1, column=0, sticky="w")
recipient_phone = ttk.Entry(left_frame, width=25)
recipient_phone.grid(row=1, column=1)

ttk.Label(left_frame, text="地址").grid(row=2, column=0, sticky="w")
recipient_address = ttk.Entry(left_frame, width=25)
recipient_address.grid(row=2, column=1)

ttk.Label(left_frame, text="邮编").grid(row=3, column=0, sticky="w")
recipient_zip = ttk.Entry(left_frame, width=25)
recipient_zip.grid(row=3, column=1)

# 寄件人输入区域
right_frame = ttk.LabelFrame(frame, text="寄件人信息", padding=8)
right_frame.grid(row=0, column=1, padx=10)

ttk.Label(right_frame, text="姓名").grid(row=0, column=0, sticky="w")
sender_name = ttk.Entry(right_frame, width=25)
sender_name.grid(row=0, column=1)
sender_name.insert(0, DEFAULT_SENDER_NAME)

ttk.Label(right_frame, text="电话").grid(row=1, column=0, sticky="w")
sender_phone = ttk.Entry(right_frame, width=25)
sender_phone.grid(row=1, column=1)
sender_phone.insert(0, DEFAULT_SENDER_PHONE)

ttk.Label(right_frame, text="地址").grid(row=2, column=0, sticky="w")
sender_address = ttk.Entry(right_frame, width=25)
sender_address.grid(row=2, column=1)
sender_address.insert(0, DEFAULT_SENDER_ADDRESS)


# 条码区域
barcode_frame = ttk.Frame(frame)
barcode_frame.grid(row=1, column=0, columnspan=2, pady=10)

ttk.Label(barcode_frame, text="条码数据").grid(row=0, column=0, padx=5)
barcode_entry = ttk.Entry(barcode_frame, width=30)
barcode_entry.grid(row=0, column=1, padx=5)
ttk.Button(barcode_frame, text="生成条码数据", command=on_generate_barcode).grid(row=0, column=2, padx=5)

# 打印按钮
ttk.Button(frame, text="生成 PDF", command=on_generate_pdf).grid(row=2, column=0, columnspan=2, pady=10)

root.mainloop()
