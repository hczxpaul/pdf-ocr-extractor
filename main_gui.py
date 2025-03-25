import tkinter as tk
from tkinter import filedialog, ttk
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import ollama
import re
import threading
import os
import pandas as pd  # 用于导出 Excel

def extract_text_from_pdf(file_path):
    try:
        doc = fitz.open(file_path)
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
             "2、但文本中没有明确具体责任人时，责任方也可能是对应政府部门，" \
             "3、文件编号中可能会出现各种括号，统一为【】" \
             "输出样例：" \
             "文号：xxx【20xx】 xxx 号" \
             "标题：xxx政府关于xxx" \
             "责任人：xxx政府" \
             "发布时间：2xxx年x月x日"

    response = ollama.chat(
        model='deepseek-r1:14b',
        messages=[{'role': 'user', 'content': prompt + text}]
    )
    return response['message']['content']

def select_pdfs(root, result_table, processed_label):
    global pdf_paths, total_files, processed_files
    pdf_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    if pdf_paths:
        total_files = len(pdf_paths)  # 更新总文件数
        processed_files = 0
        update_progress(processed_files, total_files, processed_label)  # 更新计数器标签
        threading.Thread(target=process_pdfs, args=(root, result_table, processed_label)).start()

def update_progress(processed_files, total_files, processed_label):
    processed_label.config(text=f"已处理文件数量: {processed_files}/{total_files}")

def process_pdfs(root, result_table, processed_label):
    global processed_files
    for pdf_path in pdf_paths:
        extracted_text = extract_text_from_pdf(pdf_path)
        summary = ollama_pdf_summary(extracted_text)
        cleaned_response = re.sub(r'<think>.*?</think>', '', summary, flags=re.DOTALL)
        
        # 提取文号、标题、责任人、发布时间
        doc_number = re.search(r"文号：(.*)", cleaned_response)
        title = re.search(r"标题：(.*)", cleaned_response)
        responsible_person = re.search(r"责任人：(.*)", cleaned_response)
        publish_date = re.search(r"发布时间：(.*)", cleaned_response)

        doc_number = doc_number.group(1) if doc_number else "未提取"
        title = title.group(1) if title else "未提取"
        responsible_person = responsible_person.group(1) if responsible_person else "未提取"
        publish_date = publish_date.group(1) if publish_date else "未提取"

        # 获取文件名而不是完整路径
        file_name = os.path.basename(pdf_path)

        # 在表格中插入结果
        root.after(0, lambda: result_table.insert("", "end", values=(file_name, doc_number, title, responsible_person, publish_date)))
        
        # 更新进度
        processed_files += 1
        root.after(0, lambda: update_progress(processed_files, total_files, processed_label))  # 确保线程安全

def export_to_excel(result_table):
    # 获取表格中的所有数据
    rows = result_table.get_children()
    data = []
    for row in rows:
        data.append(result_table.item(row, "values"))
    
    # 将数据导出为 Excel 文件
    if data:
        # 获取当前时间戳
        timestamp = pd.Timestamp.now().strftime("%Y%m%d%H%M%S")
        default_filename = f"摘要导出-{timestamp}.xlsx"
        
        # 打开保存对话框并预填文件名
        file_path = filedialog.asksaveasfilename(
            initialfile=default_filename,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            df = pd.DataFrame(data, columns=["文件名", "文号", "标题", "责任人", "发布时间"])
            df.to_excel(file_path, index=False)
            print(f"结果已导出到 {file_path}")

def main():
    global result_table, processed_label, pdf_paths, total_files, processed_files

    root = tk.Tk()
    root.title("PDF OCR Extractor")

    # 初始化计数器变量
    pdf_paths = []
    total_files = 0
    processed_files = 0

    # 左侧选择文件按钮
    left_frame = tk.Frame(root)
    left_frame.pack(side=tk.LEFT, padx=10, pady=10)

    select_button = tk.Button(left_frame, text="选择PDF文件", command=lambda: select_pdfs(root, result_table, processed_label))
    select_button.pack(pady=5)

    # 新增导出结果按钮
    export_button = tk.Button(left_frame, text="导出结果为Excel", command=lambda: export_to_excel(result_table))
    export_button.pack(pady=5)

    # 新增计数器标签
    processed_label = tk.Label(left_frame, text="已处理文件数量: 0/0")
    processed_label.pack(pady=5)

    # 右侧结果显示表格
    right_frame = tk.Frame(root)
    right_frame.pack(side=tk.RIGHT, padx=10, pady=10)

    columns = ("文件路径", "文号", "标题", "责任人", "发布时间")
    result_table = ttk.Treeview(right_frame, columns=columns, show="headings", height=20)
    for col in columns:
        result_table.heading(col, text=col)
        result_table.column(col, width=200, anchor="w")
    result_table.pack()

    root.mainloop()

if __name__ == "__main__":
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Windows
    # pytesseract.pytesseract.tesseract_cmd = '/usr/local/bin/tesseract'  # macOS/Linux
    main()
