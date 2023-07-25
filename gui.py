import tkinter as tk
from tkinter import messagebox
import re
from cra_script import crawler
from decorator import decorator
import threading


def areaInsert(str):
    text_area.insert(tk.END, str + "\n")


# 处理学年选择，例如16,17，返回[16,17]
def extract_numbers(text):
    numbers = re.findall(r'\d+', text)
    return numbers


# 销毁窗口
def on_closing():
    window.destroy()


# 创建主窗口
window = tk.Tk()

# 设置窗口标题
window.title("Grade crawler")
# 设置窗口宽度和高度
window_width = 560
window_height = 400
window.geometry(f"{window_width}x{window_height}")
# 禁止窗口调整大小
window.resizable(False, False)

# 创建标签和输入框
label1 = tk.Label(window, text="cookie名:")
label1.place(x=5, y=15)  # 设置标签的绝对位置

entry1 = tk.Entry(window)
entry1.place(x=75, y=15)  # 设置输入框的绝对位置

label2 = tk.Label(window, text="cookie值:")
label2.place(x=225, y=15)  # 设置标签的绝对位置

entry2 = tk.Entry(window, width=35)
entry2.place(x=295, y=15)  # 设置输入框的绝对位置

label3 = tk.Label(window, text="学   年:")
label3.place(x=5, y=70)  # 设置标签的绝对位置

entry3 = tk.Entry(window)
entry3.place(x=75, y=70)  # 设置输入框的绝对位置

# 创建文本显示区域
text_area = tk.Text(window, width=79, height=21, bg="lightgray")
text_area.place(x=1, y=120)

# 创建爬虫器
cra = crawler(text_area)
# 创建修饰器
dec = decorator(text_area)


def Crawling():
    # 组装cookie
    cookie_name = entry1.get().strip()
    cookie_value = entry2.get().strip()
    year_ids = entry3.get().strip()
    if len(cookie_name) == 0 or len(cookie_value) == 0 or len(year_ids) == 0:
        messagebox.showwarning("警告", "输入框不能为空")
        return

    cookie_string = cookie_name + "=" + cookie_value
    classYearIds = extract_numbers(year_ids)
    t = threading.Thread(target=cra.Crawling_def, args=(cookie_string, classYearIds))
    t.start()


# 开始分析
def data_analysis():
    t1 = threading.Thread(target=dec.decoration, args=())
    t1.start()


button = tk.Button(window, text="开始爬取", command=Crawling)
button.place(x=300, y=68)  # 设置按钮的绝对位置
button = tk.Button(window, text="开始分析", command=data_analysis)
button.place(x=400, y=68)  # 设置按钮的绝对位置

# 关联窗口关闭事件处理程序
window.protocol("WM_DELETE_WINDOW", on_closing)

# 开始窗口的事件循环
window.mainloop()
