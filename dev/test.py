import tkinter as tk
from tkinter import ttk
import time

root = tk.Tk()
root.title('oxxo.studio')
root.geometry('200x200')

bar = ttk.Progressbar(root, mode='determinate')
bar.pack(pady=10)

def show():
    button['state'] = 'disabled'    # 進度條開始時，不能點擊按鈕
    button2['state'] = 'disabled'   # 進度條開始時，不能點擊按鈕
    for i in range(101):
        bar['value'] = i            # 每次迴圈執行時改變進度條進度
        val.set(f'{i}%')            # 每次迴圈執行時改變顯示文字
        root.update()               # 更新視窗內容 ( 很重要！ )
        time.sleep(3)
    button2['state'] = 'normal'     # 進度條結束始時，可以點擊「重來」按鈕

def clear():
      button['state'] = 'normal'    # 點擊重來按鈕時，「開始」按鈕可以點擊
      button2['state'] = 'disabled' # 點擊重來按鈕時，不能點擊「重來」按鈕
      bar['value'] = 0              # 進度條設為 0
      val.set('0%')                 # 顯示文字為 0%

val = tk.StringVar()
val.set('0%')

label = tk.Label(root, textvariable=val)
label.pack()

button = tk.Button(root, text='開始', command=show)
button.pack()
button2 = tk.Button(root, text='重來', command=clear, state='disabled')
button2.pack()

root.mainloop()