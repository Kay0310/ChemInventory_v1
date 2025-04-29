import pytesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

from PIL import Image, ImageEnhance, ImageFilter, ImageGrab
import pandas as pd
import re
import io
import tkinter as tk
from tkinter import messagebox
import tempfile

# === OCR 전처리 ===
def preprocess_image(image):
    image = image.convert('L')
    image = image.filter(ImageFilter.MedianFilter())
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2)
    return image

# === 제품명 추출 ===
def extract_product_name(image):
    image = preprocess_image(image)
    text = pytesseract.image_to_string(image)
    match = re.search(r"Product identifier.*?:\s*(.*)", text)
    return match.group(1).strip() if match else "UNKNOWN"

# === 화학물질 데이터 추출 ===
def extract_chemical_data(image):
    image = preprocess_image(image)
    text = pytesseract.image_to_string(image)
    lines = text.split("\n")
    data = []
    for line in lines:
        if re.search(r"\d{2,5}-\d{2}-\d", line):
            parts = re.split(r"\s{2,}|\t+", line.strip())
            if len(parts) >= 4:
                data.append(parts[:4])
            elif len(parts) == 3:
                data.append([parts[0], "None", parts[1], parts[2]])
    return data

# === 붙여넣기 처리 ===
def handle_paste():
    try:
        image = ImageGrab.grabclipboard()
        if not image:
            messagebox.showerror("오류", "클립보드에 이미지가 없습니다.")
            return

        product_name = extract_product_name(image)
        chemical_data = extract_chemical_data(image)

        if not chemical_data:
            messagebox.showwarning("결과 없음", "화학물질 데이터를 찾을 수 없습니다.")
            return

        df = pd.DataFrame(chemical_data, columns=[
            "3.화학물질명(필수)", "4.관용명및이명", "5.CAS번호", "21.함유량"
        ])
        df.insert(0, "2.제품명", product_name)

        file_path = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False).name
        df.to_excel(file_path, index=False)
        messagebox.showinfo("완료", f"제품명: {product_name}\n엑셀 저장 완료:\n{file_path}")
    except Exception as e:
        messagebox.showerror("에러 발생", str(e))

# === GUI 구성 ===
root = tk.Tk()
root.title("📋 MSDS 캡처 붙여넣기 인식기")
root.geometry("400x200")

label = tk.Label(root, text="MSDS 화면을 캡처한 뒤 Ctrl+V 를 눌러 붙여넣기 하세요", wraplength=300)
label.pack(pady=30)

paste_btn = tk.Button(root, text="📥 클립보드 이미지 붙여넣기", command=handle_paste)
paste_btn.pack(pady=10)

root.mainloop()
