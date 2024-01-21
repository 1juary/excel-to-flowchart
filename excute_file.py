import tkinter as tk
from Flowchart5 import create_graph_from_excel
import os
import sys

def execute_flowchart():
    file_name = entry.get()
    if getattr(sys, 'frozen', False):
        file_path = os.path.dirname(sys.executable)
    else:
        file_path = os.path.abspath(os.path.dirname(__file__))
    file_path_with_xls = os.path.join(file_path, file_name + '.xls')
    sheet_name = 'Sheet1'
    DOE_name = file_name
    create_graph_from_excel(file_path_with_xls, sheet_name, DOE_name)

# 创建主窗口
window = tk.Tk()
window.title("Flowchart Executor")

# 创建标签和输入框
label = tk.Label(window, text="请输入文件名（不包含路径和扩展名）：")
label.pack()
entry = tk.Entry(window)
entry.pack()

# 创建执行按钮
button = tk.Button(window, text="执行", command=execute_flowchart)
button.pack()

# 运行主循环
window.mainloop()