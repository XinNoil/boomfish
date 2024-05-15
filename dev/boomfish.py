# -*- coding: UTF-8 -*-
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pystray
from PIL import Image
from pymsgbox import prompt
from pystray import MenuItem
from tools import get_resource_path
from segment_text import segment_text
from export_session import export_session
import threading
from functools import partial

SEGMENT = 1
EXPORT  = 2

def update_progress(window, progress_bar, val, item_text, total_files, progress_value):
    progress_bar['value'] = progress_value  # 更新进度值
    val.set(f'{item_text}进度: {progress_value}/{total_files}')
    window.update()  # 更新 GUI
    
def browse_file(process_type, window, progress_bar, item_text, val, chekcbox_var): # , val, label
    message = '未选中文件'
    if process_type == EXPORT:
        file_type = 'Excel Files'
        file_ext = 'xlsx'
        func = export_session
    elif process_type == SEGMENT:
        file_type = 'txt Files'
        file_ext = 'txt'
        func = segment_text
        
    filepaths = filedialog.askopenfilenames(filetypes=[(file_type, '*.' + file_ext)])
    total_files = len(filepaths)
    progress_bar.config(maximum=total_files) 
    update_progress(window, progress_bar, val, item_text, total_files, 0)
    messages = []
    for i,filepath in enumerate(filepaths, start=1):
        messages.append(func(filepath, chekcbox_var=chekcbox_var.get()))
        update_progress(window, progress_bar, val, item_text, total_files, i)
    if total_files > 0:
        messagebox.showinfo("整好了", '\n'.join(messages))
    else:
        messagebox.showinfo("没整好", '未选中文件')
    
def process(icon, item):
    # 创建主窗口
    window = tk.Tk()
    window.title("蹦蹦炸弹")
    window.geometry("300x160")
    checkbox_var = tk.IntVar(window)
    if '养个鱼（正文白表）' == item.text:
        process_type = SEGMENT
        checkbox_var.set(0)
        checkbox = ttk.Checkbutton(window, text="颜色标记（处理速度很慢，建议在Excel里标记）", variable=checkbox_var)
        checkbox.pack()
        pro_text = '养鱼'
    elif '炸个鱼（出场表）' == item.text:
        process_type = EXPORT
        pro_text = '炸鱼'
    progress_bar = ttk.Progressbar(window, length=100, mode='determinate')
    progress_bar.pack()
    val = tk.StringVar(window)
    val.set(f'{pro_text}吗？')
    label = tk.Label(window, textvariable=val)
    label.pack(padx=10, pady=10)
    browse_button = tk.Button(window, text='选择文件', command=partial(browse_file,process_type, window, progress_bar, pro_text, val, checkbox_var)) # , label
    browse_button.pack(pady=20, padx=10)
    window.focus_force()
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
    title_pattern = r"^##(.*)第\s*(\d+)\s*章\s*(.*)$"
    main()