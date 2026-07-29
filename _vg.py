import ttkbootstrap as ttkb
from tkinter import ttk
root = ttkb.Window(themename="sandstone-light")
root.geometry("620x240+30+30")
tv = ttk.Treeview(root, columns=("a", "b", "c"), show="headings", height=5)
for col, t in (("a", "项目"), ("b", "检测方法"), ("c", "描述")):
    tv.heading(col, text=t)
    tv.column(col, width=180, anchor="center")
tv.pack(fill="both", expand=True, padx=10, pady=10)
tv.insert("", "end", values=("固化剂", "GC-MS", "示例一"))
tv.insert("", "end", values=("水分", "卡尔费休", "示例二"))
root.update_idletasks(); root.update()

from PIL import ImageGrab
x, y = root.winfo_rootx(), root.winfo_rooty()
ImageGrab.grab(bbox=(x, y, x+root.winfo_width(), y+root.winfo_height())).save("_grid_def.png")

# 尝试通过 style 增加 fieldbackground 等，看是否有列分隔线
ttkb.Style().configure("Treeview", borderwidth=1)
root.update_idletasks(); root.update()
ImageGrab.grab(bbox=(x, y, x+root.winfo_width(), y+root.winfo_height())).save("_grid_styled.png")
root.destroy(); print("done")
