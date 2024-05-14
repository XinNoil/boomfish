# -*- coding: UTF-8 -*-
import os
import sys
import threading
import time
import chardet
import re
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

import pystray
# import win32api
# import win32con
from PIL import Image
from pymsgbox import prompt
from pystray import MenuItem
 
def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def save_to_xlsx(df, filepath, sheet_name):
    try:
        with pd.ExcelWriter(filepath, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        outfile = filepath
    except:
        postfix = filepath.split('.')[-1]
        outfile = filepath.replace(f'.{postfix}', 'out.xlsx')
        with pd.ExcelWriter(outfile, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return f'表格已保存至 “{outfile}”，表名为 “{sheet_name}”'

def export_chuchang(filepath):
    # 读入表格
    text_df = pd.read_excel(filepath, sheet_name="正文表")[['集数', '标题', '角色', '正文']] # 集数，标题，角色，正文
    characters_df = pd.read_excel(filepath, sheet_name="角色表")[['角色', '主播', '性别', '年龄', '人设']] # 角色，主播，性别，年龄，人设
    text_df.dropna(how='all', inplace=True)
    text_df.fillna(method='ffill', inplace=True)
    text_count_df = text_df.groupby(['集数', '标题', '角色']).count().reset_index()[['集数', '标题', '角色', '正文']]

    # 合并表格
    merged_df = pd.merge(characters_df, text_count_df, on='角色')
    merged_df.rename(columns={'正文': '句数'}, inplace=True)
    merged_df = merged_df.reindex(columns=['集数', '标题', '角色', '主播', '性别', '年龄', '人设', '句数'])

    # 自定义排序序列
    custom_order = dict(zip(characters_df['角色'], range(len(characters_df['角色']))))
    merged_df['角色'] = pd.Categorical(merged_df['角色'], categories=custom_order, ordered=True)
    merged_df = merged_df.sort_values(by=['集数', '角色'])

    # 保存表格
    xls = pd.ExcelFile(filepath)
    sheet_name = "出场表"
    new_id = 1
    while sheet_name in xls.sheet_names:
        sheet_name = f'出场表_{new_id}'
        new_id += 1
    return save_to_xlsx(merged_df, filepath, sheet_name)

def read_file(file_name):
        with open(file_name, 'rb') as f:
            result = chardet.detect(f.read())
            encoding = result['encoding']
        file_=open(file_name, 'r', encoding=encoding, errors='ignore')
        str_list = file_.read().splitlines()
        file_.close()
        return str_list

def export_zhengwen(filepath):
    lines = read_file(filepath)
    lines = list(filter(lambda x: len(x)>0, lines))

    # 找到标题
    global title_pattern
    pattern = title_pattern
    titles = list(filter(lambda x: len(re.findall(title_pattern, x))==1, lines))

    # 构建df
    df = pd.DataFrame(lines, columns=['正文'])
    df['集数'] = np.nan
    df['标题'] = np.nan
    df['角色'] = ''
    df.loc[df['正文'].isin(titles), '集数'] = df.loc[df['正文'].isin(titles), '正文'].apply(lambda x: f'第{re.findall(pattern, x)[0][1]}章')
    df.loc[df['正文'].isin(titles), '标题'] = df.loc[df['正文'].isin(titles), '正文'].apply(lambda x: re.findall(pattern, x)[0][2].strip())
    df.fillna(method='ffill', inplace=True)
    df = df.reindex(columns=['集数', '标题', '角色', '正文'])

    # 根据索引去除这些行
    df = df[~df['正文'].isin(titles)]
    return save_to_xlsx(df, filepath, '正文白表')

def browse_file():
    global process_type
    message = ''
    if process_type == 'chuchuang':
        filepath = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
        if filepath: message = export_chuchang(filepath)
    elif process_type == 'zhengwen':
        filepath = filedialog.askopenfilename(filetypes=[('text Files', '*.txt')])
        if filepath: message = export_zhengwen(filepath)
    if len(message):
        messagebox.showinfo("整好了", message)
    else:
        messagebox.showinfo("没整好", message)

def process(icon, item):
    # 创建主窗口
    window = tk.Tk()
    
    window.title("蹦蹦炸弹")
    window.geometry("300x100")

    global process_type
    if '养个鱼（正文白表）' in item.text:
        process_type = 'zhengwen'
        # # 创建按钮
        # browse_button = tk.Button(window, text=item.text, command=browse_file)
        # browse_button.pack(side=tk.LEFT, padx=10, pady=20)

        # set_button = tk.Button(window, text='标题模式', command=set_title_pattern)
        # set_button.pack(side=tk.RIGHT, padx=10, pady=20)
    elif '炸个鱼（出场表）' in item.text:
        process_type = 'chuchuang'
    # 创建按钮
    browse_button = tk.Button(window, text=item.text, command=browse_file)
    browse_button.pack(pady=20, padx=10)

    # 运行主循环
    window.mainloop()

# def set_title_pattern():
#     global title_pattern
#     return_value = prompt(text='输入', title='章节标题规则', default=title_pattern)
#     if return_value is not None:
#         title_pattern = return_value

def main():
    def on_exit(icon, item):
        global is_exit
        is_exit = True
        icon.stop()
 
    def start_icon_threading(icon):
        icon_threading = threading.Thread(target=icon.run)
        icon_threading.setDaemon(True)
        icon_threading.start()
        return icon_threading
 
    menu = (MenuItem(text='养个鱼（正文白表）', action=process),
            MenuItem(text='炸个鱼（出场表）', action=process, default=True),
            MenuItem(text='Exit', action=on_exit),
            )
    
    image = Image.open(get_resource_path("static/img.png"))
    icon = pystray.Icon("name", image, "炸鱼吗？点我点我！", menu)
    icon_threading = start_icon_threading(icon)

    while True:
        time.sleep(0.5)
        if is_exit:
            sys.exit(0)
        if not icon_threading.is_alive():
            icon_threading = start_icon_threading(icon)

if __name__ == "__main__":
    print('start')
    is_exit = False
    process_type = 'chuchang'
    title_pattern = r"^##(.*)第\s*(\d+)\s*章\s*(.*)$"
    main()