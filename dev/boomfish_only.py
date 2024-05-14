import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def process_file(filepath):
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

    try:
        with pd.ExcelWriter(filepath, engine='openpyxl', mode='a') as writer:
            merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
        message = f'表格已保存至{filepath}，表名为{sheet_name}'
    except:
        postfix = filepath.split('.')[-1]
        outfile = filepath.replace(f'.{postfix}', f'out.{postfix}')
        with pd.ExcelWriter(outfile, engine='openpyxl') as writer:
            merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
        message = f'表格已保存至{outfile}，表名为{sheet_name}'

    messagebox.showinfo("处理结果", message)

def browse_file():
    filepath = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    if filepath:
        process_file(filepath)

# 创建主窗口
window = tk.Tk()
window.title("文件处理")
window.geometry("300x100")

# 创建按钮
browse_button = tk.Button(window, text="选择文件", command=browse_file)
browse_button.pack(pady=20)

# 运行主循环
window.mainloop()