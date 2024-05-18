# -*- coding: UTF-8 -*-
import sys
import threading
import time
import pystray
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
from pymsgbox import prompt
from pystray import MenuItem
from tools import get_resource_path
from segment_text import segment_text
from export_session import export_session
from functools import partial

SEGMENT = 1
EXPORT  = 2

def update_progress(window, progress_bar, val, item_text, total_files, progress_value):
    progress_bar['value'] = progress_value  # 更新进度值
    val.set(f'{item_text}进度: {progress_value}/{total_files}')
    window.update()  # 更新 GUI
    
def browse_file(process_type, window, progress_bar, item_text, val, kwargs):
    if process_type == EXPORT:
        file_type = 'Excel Files'
        file_ext = 'xlsx'
    elif process_type == SEGMENT:
        file_type = 'txt Files'
        file_ext = 'txt'
        checkbox_var = kwargs['checkbox_var']
        entry = kwargs['entry']
        
    filepaths = filedialog.askopenfilenames(filetypes=[(file_type, '*.' + file_ext)])
    total_files = len(filepaths)
    progress_bar.config(maximum=total_files) 
    update_progress(window, progress_bar, val, item_text, total_files, 0)
    messages = []
    for i,filepath in enumerate(filepaths, start=1):
        if process_type == EXPORT:
            messages.append(export_session(filepath))
        elif process_type == SEGMENT:
            messages.append(segment_text(filepath, checkbox_var=checkbox_var.get(), title_pattern=entry.get()))
        update_progress(window, progress_bar, val, item_text, total_files, i)
    if total_files > 0:
        messagebox.showinfo("整好了", '\n'.join(messages))
    else:
        messagebox.showinfo("没整好", '未选中文件')
    
def process(icon, item):
    # 创建主窗口
    window = tk.Tk()
    window.title(item.text)
    # 创建进度条
    progress_bar = ttk.Progressbar(window, length=100, mode='determinate')
    # 创建进度标签
    val = tk.StringVar(window)
    label = tk.Label(window, textvariable=val)

    if '养个鱼（正文白表）' == item.text:
        window.geometry("300x240")
        process_type = SEGMENT
        pro_text = '养鱼'

        # 创建勾选框
        checkbox_var = tk.IntVar(window)
        checkbox = ttk.Checkbutton(window, text="颜色标记（处理速度很慢，建议在Excel里标记）", variable=checkbox_var)
        checkbox_var.set(0)

        # 创建输入框
        global title_pattern
        entry = tk.Entry(window, width=50)
        entry.insert(0, title_pattern)
        
        # pack all
        checkbox.pack(padx=10, pady=10)
        entry.pack(padx=10, pady=10)
        label.pack(padx=10, pady=10)

        # 创建文件按钮
        browse_button = tk.Button(window, text='选择文件', command=partial(browse_file,process_type, window, progress_bar, pro_text, val, {'checkbox_var':checkbox_var, 'entry':entry}))
    
    elif '炸个鱼（出场表）' == item.text:
        window.geometry("300x160")
        process_type = EXPORT
        pro_text = '炸鱼'

        # 创建文件按钮
        browse_button = tk.Button(window, text='选择文件', command=partial(browse_file,process_type, window, progress_bar, pro_text, val, None))
    
    val.set(f'{pro_text}进度: {0}/{0}')

    progress_bar.pack(padx=10, pady=10)
    label.pack(padx=10, pady=10)
    browse_button.pack(padx=10, pady=10)

    window.focus_force()
    window.mainloop()

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