import tkinter as tk
from tkinter import filedialog
import pymupdf  # PyMuPDF
import pytesseract
from PIL import Image
import ollama
import re



def extract_text_from_pdf(file_path):
    """
    从指定路径的PDF文件中提取所有文本内容，包括纯文本和图片中的文字。

    参数:
    file_path (str): PDF文件的路径

    返回:
    str: 提取的文本内容
    """
    try:
        doc = pymupdf.open(file_path)
        text = ""
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            
            # 获取页面的纯文本
            page_text = page.get_text("text")
            if page_text.strip():
                text += page_text + "\n"
            else:
                # 如果没有纯文本，则尝试从图片中提取文字
                pix = page.get_pixmap()
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                ocr_text = pytesseract.image_to_string(img, lang='chi_sim')
                text += ocr_text + "\n"
        
        return text
    except Exception as e:
        print(f"无法读取PDF文件: {e}")
        return ""
    
def ollama_pdf_summary(text):
    prompt = "下列文本是通过扫描公文pdf生成的，" \
    "请从中提取文件的文号、标题、责任人、发布时间，除此之外不要输出任何其他内容，" \
    "注意项：" \
    "1、某些文件在标题附近可能会出现发布文件的机构名，这时需要将机构名也加入标题中" \
    "2、但文本中没有明确具体责任人时，责任方也可能是对应政府部门" \
    "输出样例：" \
    "文号：xxx" \
    "标题：xxx" \
    "责任人：xxx" \
    "发布时间：xxx"

    response = ollama.chat(
    model='deepseek-r1:14b',
        messages=[
            {'role': 'user', 'content': prompt+text}
        ]
    )
    return response['message']['content']



def main():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])

    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

    if pdf_path:
        extracted_text = extract_text_from_pdf(pdf_path)
        summary = ollama_pdf_summary(extracted_text)
        cleaned_response = re.sub(r'<think>.*?</think>', '', summary, flags=re.DOTALL)
        print(cleaned_response)

if __name__ == "__main__":
    main()