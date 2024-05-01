#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @PyCharm   :2023.3.4.PY-233.14475.56
# @Python    :3.12
# @FileName  :ui.py
# @Time      :2024/5/2 下午11:30
# @Author    :997
# @E-mail    :A997sama@outlook.com
# --------------------------------

import tkinter as tk
from tkinter import ttk, filedialog, font
from PIL import ImageGrab
import pyodbc, io, base64, threading, Baidu_Api_ocr, hashlib, time, clipboard

event = threading.Event()  # 检测状态
monitoring_active = False
last_image_hash = None


def get_clipboard_content_type():
    if ImageGrab.grabclipboard():
        return 'image'
    elif clipboard.paste():
        return 'text'
    else:
        return None


def get_image_hash():
    img = ImageGrab.grabclipboard()
    if img:
        img_hash = hashlib.sha256(img.tobytes()).hexdigest()
        return img_hash
    return None


def create_ui():
    window = tk.Tk()
    window.title('A Tool')
    window.geometry("280x240+50+50")
    window.attributes('-topmost', True)
    global lab_file_path, combo, btn3, lab4, text_box, text_box_question, text_box_answer

    lab1 = tk.Label(window, text="第一步", width=5, height=1)
    lab2 = tk.Label(window, text="第二步", width=5, height=1)
    lab3 = tk.Label(window, text="第三步", width=5, height=1)
    lab1.grid(row=0, column=0, padx=10, pady=10)
    lab2.grid(row=1, column=0, padx=10, pady=5)
    lab3.grid(row=2, column=0, padx=10, pady=5)

    btn1 = ttk.Button(window, text="选择文件", width=10, command=choose_file)
    btn2 = ttk.Button(window, text="选择表", width=10, state="disabled")
    btn3 = ttk.Button(window, text="监控开关", width=10, command=switch_monitoring_status, state="disabled")
    btn1.grid(row=0, column=1, padx=0, pady=10)
    btn2.grid(row=1, column=1, padx=0, pady=5)
    btn3.grid(row=2, column=1, padx=0, pady=5)

    lab_file_path = tk.Label(window, text="", width=15, bg="white")
    lab_file_path.grid(row=0, column=2, padx=10, pady=10)

    combo_frame = ttk.Frame(window)
    combo_frame.grid(row=1, column=2, padx=10, pady=5, sticky="w")

    combo = ttk.Combobox(combo_frame, width=13, state="readonly")
    combo['values'] = ()
    combo.grid(row=3, column=1)

    lab4 = tk.Label(window, text="状态：关闭", width=13, height=1)
    lab4.grid(row=2, column=2, padx=10, pady=10)

    text_font = font.Font(size=10)

    text_box_question = tk.Text(window, width=36, height=3, font=text_font)
    text_box_question.grid(row=3, column=0, padx=10, columnspan=3, rowspan=1)

    text_box_answer = tk.Text(window, width=36, height=3, font=text_font)
    text_box_answer.grid(row=4, column=0, padx=10, pady=5, columnspan=3, rowspan=1)

    text_box = tk.Text(window, width=36, height=16)
    text_box.grid(row=0, column=3, padx=10, pady=10, columnspan=1, rowspan=5)

    window.mainloop()


def choose_file():
    file_path = filedialog.askopenfilename(title="选择Access数据库文件",
                                           filetypes=[("Access数据库文件", "*.mdb;*.accdb")])
    if file_path:
        lab_file_path.config(text=file_path)
        conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                    f'DBQ={lab_file_path.cget("text")};')
        with pyodbc.connect(conn_str) as conn:
            tables = conn.cursor().tables(tableType='TABLE')
            table_list = [table.table_name for table in tables]
            for i, table_name in enumerate(table_list, 1):
                add_option(table_name)
    btn3.config(state="normal")


def add_option(table_name):
    current_values = combo['values']
    new_values = list(current_values)
    new_values.append(table_name)
    combo['values'] = tuple(new_values)


def switch_monitoring_status():
    global monitoring_active
    global event
    if not monitoring_active:
        monitoring_active = True
        lab4.config(text="状态：开启", foreground="green")
        output_message("\n开启监控", text_box)
        event.clear()  # 清除事件，确保线程可以正常运行
        threading.Thread(target=start_monitoring_clipboard, args=(event, text_box)).start()
    else:
        monitoring_active = False
        lab4.config(text="状态：关闭", foreground="red")
        output_message("\n关闭监控", text_box)
        event.set()  # 设置事件，通知监控程序停止运行


def output_question(question):
    text_box_question.delete(1.0, tk.END)  # 清空文本框内容
    text_box_question.insert(tk.END, question)  # 插入新内容


def output_answer(answer):
    text_box_answer.delete(1.0, tk.END)  # 清空文本框内容
    text_box_answer.insert(tk.END, answer)


def output_message(message, text_box):
    current_time = time.strftime("[%Y-%m-%d %H:%M:%S] ", time.localtime())
    text_box.insert(tk.END, current_time + message + "\n")
    text_box.see(tk.END)


def start_monitoring_clipboard(event=None, text_box=None):
    if (clipboard.copy("") != False):
        output_message("\n清空剪切板成功", text_box)
    output_message("\n开始检测...", text_box)
    # output_message(f"选择的是{combo.get()}", text_box)
    monitor_and_perform_inspection_tasks(event, text_box)
    output_message("\n停止检测...", text_box)


def monitor_and_perform_inspection_tasks(event=None, text_box=None):
    while not event.is_set():  # 检查事件对象的状态，决定是否退出监控
        check_clipboard(text_box)
        event.wait(1)  # 每隔0.5秒检测一次剪贴板内容
        # output_message("next",text_box)


def check_clipboard(text_box=None):
    global last_image_hash
    current_clipboard_content_type = get_clipboard_content_type()
    if current_clipboard_content_type == 'image':
        current_image_hash = get_image_hash()
        if current_image_hash and current_image_hash != last_image_hash:
            last_image_hash = current_image_hash
            base64_image = image_to_base64()  # 读取内存图片并转base64
            if base64_image:
                question = Baidu_Api_ocr.pic3text(base64_image)
                search_answer(question, text_box)
            else:
                output_message("\n转换失败。", text_box)


def image_to_base64():
    image = ImageGrab.grabclipboard()
    if image is not None:  # 检查图像是否存在
        img_byte_array = io.BytesIO()
        image.save(img_byte_array, format='PNG')
        base64_image = base64.b64encode(img_byte_array.getvalue()).decode('utf-8')  # 将字节流转换为Base64
        return base64_image  # 将图像转换为字节流


def search_answer(question, text_box=None):
    conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={lab_file_path.cget("text")};')  # 连接到 Access 数据库
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    table_name = combo.get()
    query = f'SELECT questions, answers FROM {table_name} WHERE questions LIKE ?'
    cursor.execute(query, ('%' + question + '%',))  # 执行 SQL 查询
    rows = cursor.fetchall()
    if rows:  # 检索结果
        for row in rows:
            output_question(row.questions)
            output_answer(row.answers)
            output_message(f"\n问题：{row.questions}\n答案：{row.answers}", text_box)
            break
    else:
        output_message("\n未找到匹配的答案。", text_box)
    cursor.close()  # 关闭连接
    conn.close()  # 关闭连接


if __name__ == "__main__":
    create_ui()
else:
    print(f"Import {__name__} Successful")
