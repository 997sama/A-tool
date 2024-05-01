#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @PyCharm   :2023.3.4.PY-233.14475.56
# @Python    :3.12
# @FileName  :Add_questions_and_answers_to_access.py
# @Time      :2024/4/30 下午14:10
# @Author    :997
# @E-mail    :A997sama@outlook.com
# --------------------------------

import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import pyodbc

def choose_excel_file():
    file_path = filedialog.askopenfilename(title="选择Excel文件", filetypes=[("Excel文件", "*.xlsx;*.xls")])
    if file_path:
        entry_excel_file.delete(0, tk.END)
        entry_excel_file.insert(0, file_path)
        load_excel_sheets(file_path)

def load_excel_sheets(file_path):
    sheets = pd.ExcelFile(file_path).sheet_names
    excel_sheet_combo['values'] = sheets

def choose_access_file():
    file_path = filedialog.askopenfilename(title="选择Access文件", filetypes=[("Access文件", "*.mdb;*.accdb")])
    if file_path:
        entry_access_file.delete(0, tk.END)
        entry_access_file.insert(0, file_path)
        load_access_tables(file_path)

def load_access_tables(file_path):
    # 连接到 Access 数据库
    conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={file_path};'
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    # 获取 Access 数据库中的表名
    tables = [table.table_name for table in cursor.tables(tableType='TABLE')]
    access_table_combo['values'] = tables

    # 关闭连接
    cursor.close()
    conn.close()

def import_data():
    excel_file_path = entry_excel_file.get()
    access_file_path = entry_access_file.get()
    access_table_name = access_table_combo.get()

    # 读取 Excel 数据
    selected_sheet = excel_sheet_combo.get()
    df = pd.read_excel(excel_file_path, sheet_name=selected_sheet, header=None)

    # 连接到 Access 数据库
    conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={access_file_path};'
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()

    # 检查数据库中已存在的题目
    existing_questions = set()
    cursor.execute(f"SELECT questions FROM {access_table_name}")
    for row in cursor.fetchall():
        existing_questions.add(row[0])

    # 插入数据到 Access 数据库
    for index, row in df.iterrows():
        question = row.iloc[0]  # 使用行列编号来获取题目（第一列）
        answer = row.iloc[1]  # 使用行列编号来获取答案（第二列）
        if question not in existing_questions:
            cursor.execute(f"INSERT INTO {access_table_name} (questions, answers) VALUES (?, ?)", (question, answer))
            conn.commit()
        else:
            print(f"题目 '{question}' 已存在于数据库中，跳过插入操作。")

    # 关闭连接
    cursor.close()
    conn.close()



# 创建 tkinter 窗口
window = tk.Tk()
window.title("导入数据")

# Excel 文件选择部件
tk.Label(window, text="Excel文件：").grid(row=0, column=0, padx=5, pady=5)
entry_excel_file = tk.Entry(window, width=40)
entry_excel_file.grid(row=0, column=1, padx=5, pady=5)
tk.Button(window, text="选择文件", command=choose_excel_file).grid(row=0, column=2, padx=5, pady=5)

# Excel 表选择部件
tk.Label(window, text="Excel表：").grid(row=1, column=0, padx=5, pady=5)
excel_sheet_combo = ttk.Combobox(window)
excel_sheet_combo.grid(row=1, column=1, padx=5, pady=5)

# Access 文件选择部件
tk.Label(window, text="Access文件：").grid(row=2, column=0, padx=5, pady=5)
entry_access_file = tk.Entry(window, width=40)
entry_access_file.grid(row=2, column=1, padx=5, pady=5)
tk.Button(window, text="选择文件", command=choose_access_file).grid(row=2, column=2, padx=5, pady=5)

# Access 表选择部件
tk.Label(window, text="Access表：").grid(row=3, column=0, padx=5, pady=5)
access_table_combo = ttk.Combobox(window)
access_table_combo.grid(row=3, column=1, padx=5, pady=5)

# 导入按钮
tk.Button(window, text="导入数据", command=import_data).grid(row=4, column=0, columnspan=3, padx=5, pady=10)

# 运行窗口主循环
window.mainloop()
