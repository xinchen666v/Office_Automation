import tkinter as tk
from tkinter import filedialog
from tkinter.ttk import Progressbar
from tkinter import messagebox
import pandas as pd
from tkinter import ttk
from tkinter import *
import shutil
import time


# 获取待处理的文件路径
def browse_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        input_file.delete(0, tk.END)
        input_file.insert(0, file_path)


# 获取处理文件的存放路径
def output_browse_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        output_file.delete(0, tk.END)
        output_file.insert(0, file_path)


# 文件的处理
def process_file():
    print("开始拼命为您处理文件！")
    # 获取输入文件路径和输出文件路径
    input_file_path = input_file.get()
    output_file_path = output_file.get()

    # 如果输入文件路径和输出文件路径都有值
    if input_file_path and output_file_path:
        # 在这里添加你的处理代码
        # 读取Excel文件
        df = pd.read_excel(input_file_path)

        # 遍历数据帧，找到A列中值为0的行并删除
        df = df[(df['{}'.format(combobox.get())] != 0) & (df['{}'.format(combobox.get())].notnull())]

        # 将修改后的数据帧写入新的Excel文件
        df.to_excel(output_file_path, index=False, header=True)

        # tkinter进度条组件
        shutil.copy2(input_file_path, output_file_path, follow_symlinks=True)
        progress_bar['value'] = 0

        # 进度条组件
        for i in range(101):
            progress_bar['value'] = i
            root.update()
            time.sleep(0.05)

        # 处理完成
        print("处理完成")
        tk.messagebox.showinfo(title='提示', message='处理成功！', icon='info')
    else:
        print("请先选择文件")
        tk.messagebox.showinfo(title='提示', message='亲，是不是选错文件夹了呢！', icon='info')
        pass


def choose(event):
    # 选中事件
    print("选中的数据:{}".format(combobox.get()))
    print("value的值:{}".format(value.get()))


root = tk.Tk()
root.title("Excel处理工具，将选中月份的列中0和空的所在行删除")

input_file = tk.Entry(root, width=50)
input_file.pack(pady=10)

browse_button = tk.Button(root, text="来看看哪个文件需要处理", command=browse_file)
browse_button.pack(pady=5)

output_file = tk.Entry(root, width=50)
output_file.pack(pady=10)

output_browse_button = tk.Button(root, text="将文件存放在哪里呢？", command=output_browse_file)
output_browse_button.pack(pady=5)

screenwidth = root.winfo_screenwidth()  # 屏幕宽度
screenheight = root.winfo_screenheight()  # 屏幕高度
width = 600
height = 500
x = int((screenwidth - width) / 2)
y = int((screenheight - height) / 2)
root.geometry("{}x{}+{}+{}".format(width, height, x, y))  # 大小以及位置
value = StringVar()
value.set("请选择对应的月份")  # 默认选中CCC==combobox.current(2)

values = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月"]
combobox = ttk.Combobox(
    master=root,  # 父容器
    height=10,  # 高度,下拉显示的条目数量
    width=20,  # 宽度
    state="normal",  # 设置状态 normal(可选可输入)、readonly(只可选)、 disabled
    cursor="arrow",  # 鼠标移动时样式 arrow, circle, cross, plus...
    font=("", 16),  # 字体
    textvariable=value,  # 通过StringVar设置可改变的值
    values=values,  # 设置下拉框的选项
)
combobox.bind("<<ComboboxSelected>>", choose)
print(combobox.keys())  # 可以查看支持的参数
combobox.pack()

process_button = tk.Button(root, text="开始处理！", command=process_file)
process_button.pack(pady=5)

# 创建进度条控件
progress_bar = Progressbar(root, orient=HORIZONTAL, length=400, mode='determinate')
progress_bar.pack(pady=20)

root.mainloop()
