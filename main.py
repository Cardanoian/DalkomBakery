import tkinter as tk
import tkinter.ttk as ttk
import pandas as pd
from tkinter import font
from tkinter import filedialog
import tkinter.messagebox as msg
from datetime import date

from funcs import read_process

btn_width = 8
orders = []


def read_file():
    order_list.delete(0, tk.END)
    global orders
    file_names = filedialog.askopenfilenames(initialdir="./", title="Select File",
                                             filetypes=(
                                                 ("Data Files", ("*.xlsx", "*.xls", "*.csv")), ("All files", "*.*")))
    for file in file_names:
        # print(file[-3:])
        if file[-3:] == "csv":
            orders += read_process.make_shop(file)
        else:
            orders += read_process.naver_store(file)
    orders.sort(key=lambda x: x["수령희망일"])
    for row in orders:
        order_str = f'{row["출처"]} {row["수신자"]}, 수령희망일: {row["수령희망일"] if row["수령희망일"] else ""}, {row["상품명(수량)"]}, 휴대전화: {row["수신자휴대전화"] if row["수신자휴대전화"] else ""}, 집전화: {row["수신자전화번호"] if row["수신자전화번호"] else ""}, 우편번호: {row["수신자우편번호"]}, 주소: {row["수신자주소"]}'
        order_list.insert(tk.END, order_str)


def save_file():
    first = pd.DataFrame(orders)
    print(first.loc[:, ["출처", "수신자", "상품명(수량)", "수령희망일"]])
    # print(df)
    # df.to_excel(excel_writer="result.xlsx")


def del_order():
    if msg.askokcancel(title="확인", message="삭제하시겠습니까?"):
        for index in reversed(order_list.curselection()):
            order_list.delete(index)
            del orders[index]
            # print(len(orders), orders)


def set_font(font_size):
    order_list.config(font=font.Font(size=font_size))


# Main Window

window = tk.Tk()

window.minsize(800, 600)
window.title("Dalkom Bakery")
window.geometry("800x600+300+300")
window.resizable(True, True)

# Menu Bar

menubar = tk.Menu(window)
filemenu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="File", menu=filemenu)
filemenu.add_command(label="Open", command=read_file)
filemenu.add_separator()
filemenu.add_command(label="주문 삭제", command=del_order)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=window.quit)

fontmenu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="Font", menu=fontmenu)
fontmenu.add_command(label="12", command=lambda: set_font(12))
fontmenu.add_command(label="14", command=lambda: set_font(14))
fontmenu.add_command(label="16", command=lambda: set_font(16))
fontmenu.add_command(label="18", command=lambda: set_font(18))
fontmenu.add_command(label="20", command=lambda: set_font(20))
fontmenu.add_command(label="22", command=lambda: set_font(22))
fontmenu.add_command(label="24", command=lambda: set_font(24))

window.config(menu=menubar)

# Order Frame

notebook = ttk.Notebook(window, width=700, height=500)
notebook.pack(fill="both", expand=1)

order_frame = tk.Frame(window)

y_scrollbar = tk.Scrollbar(order_frame)
y_scrollbar.pack(side="right", fill="y")
x_scrollbar = tk.Scrollbar(order_frame)
x_scrollbar.pack(side="bottom", fill="x")

# Order List

order_list = tk.Listbox(order_frame, selectmode="multiple", height=15, yscrollcommand=y_scrollbar.set,
                        xscrollcommand=x_scrollbar.set)
order_list.pack(side="right", fill="both", expand=True)

# order_list.config(font=FONT)

y_scrollbar.config(command=order_list.yview())
x_scrollbar.config(command=order_list.xview())

notebook.add(order_frame, text="주문")

# Complete Page

complete_frame = tk.Frame(window)
label2 = tk.Label(complete_frame, text="Frame2")
label2.pack()
notebook.add(complete_frame, text="완료 주문")

# File Frame

file_frame = tk.Frame(window)
file_frame.pack(fill="x", padx=5, pady=5)

add_btn = tk.Button(
    file_frame, text="파일 열기", padx=5, pady=5, width=btn_width, command=read_file
)
add_btn.pack(side="left")

save_btn = tk.Button(
    file_frame, text="파일 저장", padx=5, pady=5, width=btn_width, command=save_file
)
save_btn.pack(side="left")

del_btn = tk.Button(
    file_frame, text="주문 지우기", padx=5, pady=5, width=btn_width, command=del_order
)
del_btn.pack(side="right")

# Main Loop

window.mainloop()
