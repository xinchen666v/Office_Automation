import tkinter as tk
from tkinter import filedialog
import pandas as pd
import time


def browse_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        input_file.delete(0, tk.END)
        input_file.insert(0, file_path)


def output_browse_file():
    file_path = filedialog.askopenfilename()
    if file_path:
        output_file.delete(0, tk.END)
        output_file.insert(0, file_path)


def process_file():
    print("开始拼命为您处理文件！")
    input_file_path = input_file.get()
    output_file_path = output_file.get()

    if input_file_path and output_file_path:
        # 在这里添加你的处理代码
        # 读取Excel文件
        df = pd.read_excel(input_file_path)

        # 遍历数据帧，找到A列中值为0的行并删除
        df = df[df['10月'] != 0]

        # 将修改后的数据帧写入新的Excel文件
        df.to_excel(output_file_path, index=False, header=True)
        print("处理完成")
    else:
        print("请先选择文件")
        pass


root = tk.Tk()
root.title("Excel处理工具")

input_file = tk.Entry(root, width=50)
input_file.pack(pady=10)

browse_button = tk.Button(root, text="来看看哪个文件需要处理", command=browse_file)
browse_button.pack(pady=5)

output_file = tk.Entry(root, width=50)
output_file.pack(pady=10)

output_browse_button = tk.Button(root, text="将文件存放在哪里呢？", command=output_browse_file)
output_browse_button.pack(pady=5)

process_button = tk.Button(root, text="开始处理！", command=process_file)
process_button.pack(pady=5)

root.mainloop()
