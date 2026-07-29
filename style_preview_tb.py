"""ttkbootstrap 风格预览 —— 顶部下拉切主题，整窗实时换皮。

用法: .venv/Scripts/python style_preview_tb.py
说明: 30+ 主题(各色 light/dark)，用 ttkbootstrap 2.0 现代控件渲染同一样例。
      底部彩色徽标(成功/失败)证明 ttk 也能做着色状态格，不只是 tk 行。
"""
import ttkbootstrap as ttk


def build(root):
    style = ttk.Style()

    # 主题切换条
    top = ttk.Frame(root, padding=(14, 10))
    top.pack(fill='x')
    ttk.Label(top, text="主题:").pack(side='left')
    themes = style.theme_names()
    cb = ttk.Combobox(top, values=themes, width=18, state='readonly')
    cb.set(style.theme_use())
    cb.pack(side='left', padx=8)
    ttk.Label(top, text="← 选主题，整窗实时变", bootstyle="secondary").pack(side='left')

    def on_change(_=None):
        style.theme_use(cb.get())
    cb.bind("<<ComboboxSelected>>", on_change)

    # 标题
    ttk.Label(root, text="检测数据登记系统",
              font=("Segoe UI", 16, "bold")).pack(pady=(2, 10))

    # 工具栏
    bar = ttk.Frame(root)
    bar.pack(fill='x', padx=14)
    for t in ["Add", "✕", "↓", "Clear", "方法"]:
        ttk.Button(bar, text=t, bootstyle="secondary").pack(side='left', padx=2)
    ttk.Button(bar, text="运行", bootstyle="primary").pack(side='right')

    # 表单卡片
    card = ttk.Labelframe(root, text=" 登录 ", padding=16)
    card.pack(fill='x', padx=14, pady=12)
    ttk.Label(card, text="用户").pack(anchor='w')
    ttk.Combobox(card, values=["马伊楠", "沈佳睿", "王君燕"],
                 state='readonly').pack(fill='x', pady=(2, 10))
    ttk.Label(card, text="验证码").pack(anchor='w')
    row = ttk.Frame(card); row.pack(fill='x', pady=(2, 4))
    ttk.Entry(row).pack(side='left', fill='x', expand=True)
    cap = ttk.Entry(row, width=6, font=("Consolas", 11, "bold"))
    cap.insert(0, "A1B2"); cap.configure(state='readonly')
    cap.pack(side='left', padx=(6, 0))
    ttk.Label(card, text="点击图片可刷新", bootstyle="secondary",
              font=("Segoe UI", 8)).pack(anchor='w')

    # 主按钮 + 彩色按钮展示
    bf = ttk.Frame(root); bf.pack(fill='x', padx=14, pady=(0, 8))
    ttk.Button(bf, text="登 录", bootstyle="primary").pack(side='left', padx=2)
    ttk.Button(bf, text="成功", bootstyle="success").pack(side='left', padx=2)
    ttk.Button(bf, text="信息", bootstyle="info").pack(side='left', padx=2)
    ttk.Button(bf, text="危险", bootstyle="danger").pack(side='left', padx=2)

    # 彩色状态徽标(替代绿/红着色单元格 —— 证明 ttk 能做)
    sf = ttk.Frame(root); sf.pack(fill='x', padx=14, pady=(0, 8))
    ttk.Label(sf, text="着色状态格:").pack(side='left', padx=(0, 6))
    for txt, bs in [("待运行", "inverse-secondary"), ("运行中", "inverse-info"),
                    ("成功", "inverse-success"), ("失败", "inverse-danger")]:
        ttk.Label(sf, text=txt, bootstyle=bs, padding=(8, 3)).pack(side='left', padx=3)

    # 状态条
    ttk.Label(root, text="状态: 就绪", bootstyle="secondary",
              anchor='w').pack(fill='x', padx=14, pady=(0, 12))


def main():
    root = ttk.Window(themename="bootstrap-light")
    root.title("ttkbootstrap 风格预览 —— 顶部下拉切主题")
    root.geometry("660x580")
    build(root)
    root.mainloop()


if __name__ == "__main__":
    main()
