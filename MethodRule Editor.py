import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk
from tkinter import messagebox, filedialog
import ttkbootstrap as ttkb  # 档1: 现代主题(sandstone-light)，ttk 控件自动套用
from method_file import load_method, save_method
import os
import paths
import random


def _make_large_cb_style(scale=1.3):
    """放大 ttk.Checkbutton 勾选框：复用 ttkbootstrap 内置位图按 scale 重建 indicator element。
    幂等标记挂在 ttkb.Style() 单例上(而非模块级字典)：SequenceMaster 经 importlib 每次重执行本模块，
    模块级缓存会被重置→重建样式→旧 PhotoImage 被 GC→已建勾选框的勾选标记消失。
    Style 单例依附 Tk root、跨模块重执行存活，故只建一次。"""
    s = ttkb.Style()
    _cached = getattr(s, '_big_cb_style_name', None)
    if _cached:
        return _cached
    try:
        from ttkbootstrap.style.layout import image_element, layout, El
        from ttkbootstrap.style.assets import RecolorRenderer
        from PIL import ImageTk
        c = s.colors
        size = int(20 * scale)
        def _img(name, ink):
            return ImageTk.PhotoImage(RecolorRenderer.render(name, (size, size), '#ffffff', ink, None, None))
        chk, unk = _img('checkbox_checked', c.primary), _img('checkbox_unchecked', c.fg)
        dchk, dunk = _img('checkbox_checked', c.border), _img('checkbox_unchecked', c.border)
        sn = 'Big.TCheckbutton'
        image_element(s, sn + '.indicator', default=chk,
                      states={'disabled selected': dchk, 'disabled': dunk,
                              '!selected': unk}, border=0, padding=0, sticky=tk.W)
        layout(s, sn, El('Checkbutton.padding', sticky=tk.NSEW, children=[
            El(sn + '.indicator', side=tk.LEFT, sticky=''),
            El('Checkbutton.focus', side=tk.LEFT, sticky='',
               children=[El('Checkbutton.label', sticky=tk.NSEW)])]))
        s.configure(sn, background=c.bg, foreground=c.fg, focuscolor=c.fg)
        s._big_cb_imgs = (chk, unk, dchk, dunk)  # 持有 PhotoImage 引用防 GC
        s._big_cb_style_name = sn  # 缓存在 Style 单例上，跨模块重执行存活
        return sn
    except Exception:
        return None


def _setup_cell_tooltip(tree, columns, wraplength=420):
    """为 tree 指定列悬停显示单元格全部内容(项目/检测方法/设备编号等长文本列)。
    wraplength: 提示框折行宽度(像素)，长文本列(如试验过程)可调大以免提示过高。"""
    tip = tk.Toplevel(tree)
    tip.withdraw()
    tip.overrideredirect(True)
    label = ttk.Label(tip, background="#ffffe0", relief="solid", borderwidth=1,
                      padding=(6, 3), wraplength=wraplength)
    label.pack()
    last = [None]  # (item, col) 避免同一格反复重绘

    def on_motion(event):
        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not item or not col:
            tip.withdraw(); last[0] = None; return
        try:
            idx = int(col.replace("#", "")) - 1
            name = str(tree["columns"][idx])
        except (ValueError, IndexError):
            tip.withdraw(); last[0] = None; return
        if name not in columns:
            tip.withdraw(); last[0] = None; return
        values = tree.item(item, "values")
        text = str(values[idx]).strip() if idx < len(values) else ""
        if not text:
            tip.withdraw(); last[0] = None; return
        if last[0] != (item, name):
            label.configure(text=text); last[0] = (item, name)
        tip.geometry(f"+{tree.winfo_rootx() + event.x + 14}+{tree.winfo_rooty() + event.y + 14}")
        tip.deiconify(); tip.lift()

    tree.bind("<Motion>", on_motion)
    tree.bind("<Leave>", lambda e: (tip.withdraw(), last.__setitem__(0, None)))


def _bind_entry_tooltip(entry):
    """鼠标悬停时显示 Entry 完整内容（用于窄输入框被截断的长文本）"""
    tip = tk.Toplevel(entry)
    tip.withdraw()
    tip.overrideredirect(True)
    label = ttk.Label(tip, background="#ffffe0", relief="solid", borderwidth=1,
                      font=("微软雅黑", 9), padding=(6, 3))
    label.pack()

    def show(_):
        text = entry.get().strip()
        if not text:
            return
        label.configure(text=text)
        tip.geometry(f"+{entry.winfo_rootx()+14}+{entry.winfo_rooty()+entry.winfo_height()+4}")
        tip.deiconify(); tip.lift()

    def hide(_):
        tip.withdraw()

    entry.bind('<Enter>', show)
    entry.bind('<Leave>', hide)


class QueryTab:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.query_rules = []
        self.exclusion_rules = []  # 新增：排除项目规则
        self.create_tab()

    def create_tab(self):
        # 创建查询条件标签页
        query_frame = ttk.Frame(self.parent)
        self.parent.add(query_frame, text="查询条件")

        # 查询条件规则管理区域
        self.create_query_rules_management(query_frame)

        # 查询策略(方法池 vs 按样品；日期窗口) —— 方法文件级，存 query_rules 外层
        strat = ttk.Frame(query_frame)
        strat.pack(side='bottom', fill='x', padx=5, pady=(2, 4))
        strat_row = ttk.Frame(strat)
        strat_row.pack(fill='x')
        ttk.Label(strat_row, text="按样品查询阈值:").pack(side='left', padx=(0, 3))
        self.sample_parallel_threshold_var = tk.StringVar(value="20")
        e_threshold = ttk.Entry(strat_row, width=6, textvariable=self.sample_parallel_threshold_var)
        e_threshold.pack(side='left', padx=(0, 12))
        ttk.Label(strat_row, text="日期窗口(天):").pack(side='left', padx=(0, 3))
        self.date_window_days_var = tk.StringVar(value="30")
        e_days = ttk.Entry(strat_row, width=6, textvariable=self.date_window_days_var)
        e_days.pack(side='left', padx=(0, 12))
        ttk.Label(strat, text="样品数≤阈值→按样品并行查；>阈值→方法池。日期窗口对方法池与逐样品查询均生效",
                  foreground="gray").pack(fill='x')
        for _e in (e_threshold, e_days):
            _e.bind("<KeyRelease>", lambda _ev: self.app.mark_modified())

    def create_query_rules_management(self, parent):
        """创建查询条件规则管理区域"""
        # 创建规则管理框架
        rules_frame = ttk.LabelFrame(parent, text="查询条件规则管理", padding=5)
        rules_frame.pack(fill='x', padx=5, pady=5)  # 不 expand：收紧到内容自然高度，treeview 按实际行数显示

        # 创建规则输入区域 - 分两行：第一行文本输入(描述/项目/检测方法)，第二行选项与按钮(Retest/最大选中数量/Mode/添加)
        input_frame = ttk.Frame(rules_frame)
        input_frame.pack(fill='x', pady=2)

        # 第一行：项目 / 检测方法
        row0 = ttk.Frame(input_frame)
        row0.pack(fill='x', pady=1)
        row0.columnconfigure(3, weight=1)  # 检测方法输入框列撑开填满

        ttk.Label(row0, text="项目:").grid(row=0, column=0, padx=(0, 5), pady=1, sticky='w')
        self.query_project_entry = ttk.Entry(row0, width=25)
        self.query_project_entry.grid(row=0, column=1, padx=(0, 10), pady=1, sticky='w')

        ttk.Label(row0, text="检测方法:").grid(row=0, column=2, padx=(0, 5), pady=1, sticky='w')
        self.query_method_entry = ttk.Entry(row0, width=20)
        self.query_method_entry.grid(row=0, column=3, padx=(0, 5), pady=1, sticky='ew')

        # 第二行：Retest / 最大选中数量 / Mode / 是否总和 / 添加
        row1 = ttk.Frame(input_frame)
        row1.pack(fill='x', pady=1)
        row1.columnconfigure(8, weight=1)  # 末尾空白撑开，把添加按钮推到右侧

        ttk.Label(row1, text="Retest:").grid(row=0, column=0, padx=(0, 5), pady=1, sticky='w')
        self.query_cancel_test_var = tk.BooleanVar()
        self.query_cancel_test_check = ttk.Checkbutton(
            row1, variable=self.query_cancel_test_var, style=self.app.large_cb_style)
        self.query_cancel_test_check.grid(row=0, column=1, padx=(0, 10), pady=1, sticky='w')

        ttk.Label(row1, text="最大选中数量:").grid(row=0, column=2, padx=(0, 5), pady=1, sticky='w')
        self.query_max_select_entry = ttk.Entry(row1, width=8)
        self.query_max_select_entry.grid(row=0, column=3, padx=(0, 10), pady=1, sticky='w')

        ttk.Label(row1, text="Mode:").grid(row=0, column=4, padx=(0, 5), pady=1, sticky='w')
        self.query_mode_var = tk.StringVar(value="方法")
        self.query_mode_combo = ttk.Combobox(row1, textvariable=self.query_mode_var,
                                             values=["方法", "样品"], state="readonly", width=8)
        self.query_mode_combo.grid(row=0, column=5, padx=(0, 10), pady=1, sticky='w')

        # 是否总和：勾选才走「总和录入」(读各组分值求和)，否则按报告解析
        ttk.Label(row1, text="是否总和:").grid(row=0, column=6, padx=(0, 5), pady=1, sticky='w')
        self.query_sum_entry_var = tk.BooleanVar()
        self.query_sum_entry_check = ttk.Checkbutton(
            row1, variable=self.query_sum_entry_var, style=self.app.large_cb_style)
        self.query_sum_entry_check.grid(row=0, column=7, padx=(0, 5), pady=1, sticky='w')

        self.add_query_rule_btn = ttkb.Button(row1, text="添加", command=self.add_query_rule, bootstyle="secondary")
        self.add_query_rule_btn.grid(row=0, column=9, padx=(5, 0), pady=1, sticky='e')

        # 创建规则显示区域
        display_frame = ttk.Frame(rules_frame)
        display_frame.pack(fill='both', expand=True, pady=5)

        # 创建规则表格 - 列顺序与输入框一致：项目/检测方法/Retest/最大选中数量；Mode 为非输入项置末
        columns = ("No", "项目", "检测方法", "Retest", "Max", "Mode", "是否总和")
        self.query_rules_tree = ttk.Treeview(display_frame, columns=columns, show="headings", height=10)

        # 设置列标题和宽度
        # 短列(No/Retest/Max/Mode/是否总和)固定窄列宽不随窗口拉伸；文本列(项目/检测方法)随窗口伸缩
        column_configs = {
            "No": {"width": 50, "anchor": "center", "stretch": False},  # 序号列居中对齐
            "项目": {"width": 200, "anchor": "center"},  # 项目列居中对齐
            "检测方法": {"width": 200, "anchor": "center"},  # 检测方法列居中对齐
            "Retest": {"width": 90, "anchor": "center", "stretch": False},  # 注销复测列居中对齐
            "Max": {"width": 80, "anchor": "center", "stretch": False},  # 最大选中数量列居中对齐
            "Mode": {"width": 90, "anchor": "center", "stretch": False},  # 录入方式列居中对齐
            "是否总和": {"width": 90, "anchor": "center", "stretch": False}  # 是否总和录入列居中对齐
        }

        for col in columns:
            self.query_rules_tree.heading(col, text=col, anchor="center")
            config = column_configs.get(col, {})
            self.query_rules_tree.column(col, width=config.get("width", 100),
                                         anchor=config.get("anchor", "center"),
                                         stretch=config.get("stretch", True))

        # 添加滚动条
        scrollbar = ttk.Scrollbar(display_frame, orient="vertical", command=self.query_rules_tree.yview)
        self.query_rules_tree.configure(yscrollcommand=scrollbar.set)

        self.query_rules_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 绑定双击事件，用于编辑规则
        self.query_rules_tree.bind("<Double-1>", self.on_query_rule_double_click)

        # 项目/检测方法 列悬停显示全部内容
        _setup_cell_tooltip(self.query_rules_tree, ("项目", "检测方法"))

        # 创建规则操作按钮 - 新增上移、下移按钮，以及排除按钮
        button_frame = ttk.Frame(rules_frame)
        button_frame.pack(fill='x', pady=3)

        # 左侧按钮
        ttkb.Button(button_frame, text="上移", command=self.move_query_rule_up, bootstyle="secondary").pack(side='left', padx=2)
        ttkb.Button(button_frame, text="下移", command=self.move_query_rule_down, bootstyle="secondary").pack(side='left', padx=2)
        ttkb.Button(button_frame, text="删除", command=self.delete_query_rule, bootstyle="danger").pack(side='left', padx=2)

        # 右侧按钮 - 排除按钮
        ttkb.Button(button_frame, text="排除", command=self.show_exclusion_dialog, bootstyle="secondary").pack(side='right', padx=2)

    def show_exclusion_dialog(self):
        """显示排除项目设置对话框"""
        dialog = ttkb.Toplevel(self.parent)
        dialog.title("排除项目设置")
        dialog.geometry("880x760")  # 增加宽度以容纳更多内容
        dialog.transient(self.parent)
        dialog.grab_set()

        # 允许对话框调整大小
        dialog.resizable(True, True)

        # 计算居中位置
        self.center_dialog(dialog, 880, 760)

        # 主容器
        main_frame = ttk.Frame(dialog, padding=10)
        main_frame.pack(fill='both', expand=True)
        main_frame.columnconfigure(0, weight=1)

        # 创建输入区域 - 所有控件在一行
        input_frame = ttk.LabelFrame(main_frame, text="添加排除项", padding=5)
        input_frame.pack(fill='x', pady=(0, 5))

        # 配置列权重，使所有输入框都能扩展
        input_frame.columnconfigure(0, weight=0)  # 项目标签
        input_frame.columnconfigure(1, weight=2)  # 项目输入框 - 增加权重
        input_frame.columnconfigure(2, weight=0)  # 方法标签
        input_frame.columnconfigure(3, weight=2)  # 方法输入框 - 增加权重
        input_frame.columnconfigure(4, weight=0)  # 添加按钮

        # 项目输入 - 增加初始宽度
        ttk.Label(input_frame, text="项目:").grid(row=0, column=0, padx=(0, 0), pady=3, sticky='w')
        exclusion_project_var = tk.StringVar()
        exclusion_project_entry = ttk.Entry(input_frame, textvariable=exclusion_project_var, width=20)
        exclusion_project_entry.grid(row=0, column=1, padx=(0, 10), pady=3, sticky='ew')

        # 检测方法输入 - 增加初始宽度
        ttk.Label(input_frame, text="检测方法:").grid(row=0, column=2, padx=(0, 0), pady=3, sticky='w')
        exclusion_method_var = tk.StringVar()
        exclusion_method_entry = ttk.Entry(input_frame, textvariable=exclusion_method_var, width=30)
        exclusion_method_entry.grid(row=0, column=3, padx=(0, 10), pady=3, sticky='ew')

        # 添加按钮
        def add_exclusion_rule():
            project = exclusion_project_var.get().strip()
            method = exclusion_method_var.get().strip()

            if not project and not method:
                messagebox.showwarning("输入错误", "请至少填写项目或检测方法")
                return

            # 检查是否已存在相同的规则
            new_rule = {
                "project": project,
                "method": method
            }

            # 检查重复
            for existing_rule in self.exclusion_rules:
                if (existing_rule.get("project", "") == project and
                        existing_rule.get("method", "") == method):
                    messagebox.showwarning("重复条目", "已存在相同的排除项目，无法重复添加")
                    return

            self.exclusion_rules.append(new_rule)
            refresh_exclusion_list()

            # 清空输入框
            exclusion_project_var.set("")
            exclusion_method_var.set("")

            # 标记已修改
            self.app.mark_modified()

        ttkb.Button(input_frame, text="添加", command=add_exclusion_rule, bootstyle="secondary").grid(row=0, column=4, padx=5, pady=3,
                                                                              sticky='e')

        # 排除项目列表
        list_frame = ttk.LabelFrame(main_frame, text="已添加的排除项目", padding=5)
        list_frame.pack(fill='both', expand=True, pady=(0, 5))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        # 创建列表
        columns = ("项目", "检测方法")
        exclusion_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)

        # 设置列宽 - 增加项目列宽度
        exclusion_tree.heading("项目", text="项目", anchor="w")
        exclusion_tree.column("项目", width=250, anchor="w")
        exclusion_tree.heading("检测方法", text="检测方法", anchor="w")
        exclusion_tree.column("检测方法", width=200, anchor="w")

        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=exclusion_tree.yview)
        exclusion_tree.configure(yscrollcommand=scrollbar.set)

        exclusion_tree.grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')

        # 绑定双击事件，用于编辑排除规则
        exclusion_tree.bind("<Double-1>", lambda event: self.on_exclusion_rule_double_click(event, exclusion_tree))

        def refresh_exclusion_list():
            # 清空现有列表
            for item in exclusion_tree.get_children():
                exclusion_tree.delete(item)

            # 添加排除规则到列表
            for rule in self.exclusion_rules:
                exclusion_tree.insert("", "end", values=(
                    rule.get("project", ""),
                    rule.get("method", "")
                ))

        # 初始刷新列表
        refresh_exclusion_list()

        # 底部按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=5)

        # 删除选中按钮
        def delete_selected_exclusion():
            selected_items = exclusion_tree.selection()
            if not selected_items:
                messagebox.showwarning("选择错误", "请先选择要删除的排除项")
                return

            if messagebox.askyesno("确认删除", "确定要删除选中的排除项吗？"):
                # 从后往前删除
                for item in reversed(selected_items):
                    index = exclusion_tree.index(item)
                    if 0 <= index < len(self.exclusion_rules):
                        del self.exclusion_rules[index]

                refresh_exclusion_list()
                self.app.mark_modified()

        ttkb.Button(button_frame, text="删除选中", command=delete_selected_exclusion, bootstyle="danger").pack(side='left', padx=5)

        # 关闭按钮
        ttkb.Button(button_frame, text="关闭", command=dialog.destroy, bootstyle="secondary").pack(side='right', padx=5)

    def on_exclusion_rule_double_click(self, event, tree):
        """双击排除规则进行编辑"""
        item = tree.selection()
        if not item:
            return

        item = item[0]
        column = tree.identify_column(event.x)
        column_index = int(column.replace('#', '')) - 1  # 列索引从0开始

        # 获取当前值
        current_values = tree.item(item, 'values')
        current_value = current_values[column_index]

        # 获取单元格坐标
        x, y, width, height = tree.bbox(item, column)

        # 创建编辑框
        entry = ttk.Entry(tree)
        entry.place(x=x, y=y-4, width=width, height=height+8)
        entry.insert(0, current_value)
        entry.focus_set()

        def save_edit(event=None):
            # 获取新值
            new_value = entry.get()

            # 更新显示
            new_values = list(current_values)
            new_values[column_index] = new_value
            tree.item(item, values=new_values)

            # 更新数据
            index = tree.index(item)
            if 0 <= index < len(self.exclusion_rules):
                rule = self.exclusion_rules[index]
                if column_index == 0:  # 项目列
                    rule["project"] = new_value
                elif column_index == 1:  # 检测方法列
                    rule["method"] = new_value

            # 标记已修改
            self.app.mark_modified()

            # 销毁编辑框
            entry.destroy()

        def cancel_edit(event=None):
            entry.destroy()

        # 绑定事件
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", cancel_edit)

    def center_dialog(self, dialog, width, height):
        """将对话框居中显示在主窗口中间"""
        # 更新对话框以确保获取正确的尺寸
        dialog.update_idletasks()

        # 获取主窗口位置和尺寸
        parent = self.parent.winfo_toplevel()
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()

        # 计算居中位置
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2

        # 设置对话框位置
        dialog.geometry(f"{width}x{height}+{x}+{y}")
    # 其他方法保持不变...
    def on_query_rule_double_click(self, event):
        """双击查询条件规则进行编辑"""
        item = self.query_rules_tree.selection()
        if not item:
            return

        item = item[0]
        column = self.query_rules_tree.identify_column(event.x)
        column_index = int(column.replace('#', '')) - 1  # 列索引从0开始

        # 序号列不允许编辑
        if column_index == 0:
            return

        # 获取当前值
        current_values = self.query_rules_tree.item(item, 'values')
        current_value = current_values[column_index]

        # 获取单元格坐标
        x, y, width, height = self.query_rules_tree.bbox(item, column)

        # 对于录入方式列，使用下拉框编辑
        column_name = self.query_rules_tree.heading(column_index)['text']
        if column_name == "Mode":
            # 创建下拉框
            combo = ttk.Combobox(self.query_rules_tree, values=["方法", "样品"], state="readonly")
            combo.place(x=x, y=y-4, width=width, height=height+8)
            combo.set(current_value)
            combo.focus_set()

            def save_combo_edit(event=None):
                if not combo.winfo_exists():
                    return
                new_value = combo.get()
                new_values = list(current_values)
                new_values[column_index] = new_value
                self.query_rules_tree.item(item, values=new_values)

                # 更新数据
                index = self.query_rules_tree.index(item)
                if 0 <= index < len(self.query_rules):
                    self.query_rules[index]["input_method"] = new_value

                # 标记已修改
                self.app.mark_modified()
                combo.destroy()

            def cancel_combo_edit(event=None):
                if combo.winfo_exists():
                    combo.destroy()

            def on_focus_out(event=None):
                # 下拉框展开会先触发 FocusOut；延迟检查焦点是否落在 popdown 上，
                # 若是则不关闭（避免一展开下拉就被销毁，导致 方法/样品 切换不顺）
                def _check():
                    if not combo.winfo_exists():
                        return
                    if "popdown" in str(self.query_rules_tree.tk.call('focus')):
                        return
                    combo.destroy()
                combo.after(10, _check)

            combo.bind("<Return>", save_combo_edit)
            combo.bind("<FocusOut>", on_focus_out)
            combo.bind("<Escape>", cancel_combo_edit)
            combo.bind("<<ComboboxSelected>>", save_combo_edit)
        else:
            # 其他列使用文本框编辑
            entry = ttk.Entry(self.query_rules_tree)
            entry.place(x=x, y=y-4, width=width, height=height+8)
            entry.insert(0, current_value)
            entry.focus_set()

            def save_edit(event=None):
                # 获取新值
                new_value = entry.get()

                # 最大选中数量校验：留空表示不限制，否则必须是非负整数
                if column_name == "Max":
                    new_value = new_value.strip()
                    if new_value and not new_value.isdigit():
                        messagebox.showwarning("输入错误", "最大选中数量必须为空或非负整数")
                        return

                # 更新显示
                new_values = list(current_values)
                new_values[column_index] = new_value
                self.query_rules_tree.item(item, values=new_values)

                # 更新数据
                index = self.query_rules_tree.index(item)
                if 0 <= index < len(self.query_rules):
                    rule = self.query_rules[index]

                    if column_name == "项目":
                        rule["project"] = new_value
                    elif column_name == "检测方法":
                        rule["method"] = new_value
                    elif column_name == "Retest":
                        rule["cancel_test"] = (new_value == "是")
                    elif column_name == "Max":
                        rule["max_select"] = new_value
                    elif column_name == "是否总和":
                        rule["sum_entry"] = (new_value == "是")

                # 标记已修改
                self.app.mark_modified()

                # 销毁编辑框
                entry.destroy()

            def cancel_edit(event=None):
                entry.destroy()

            # 绑定事件
            entry.bind("<Return>", save_edit)
            entry.bind("<FocusOut>", save_edit)
            entry.bind("<Escape>", cancel_edit)

    def move_query_rule_up(self):
        """上移选中的查询规则"""
        selected_items = self.query_rules_tree.selection()
        if not selected_items:
            messagebox.showwarning("选择错误", "请先选择要移动的规则")
            return

        if len(selected_items) > 1:
            messagebox.showwarning("选择错误", "只能选择一个规则进行移动")
            return

        item = selected_items[0]
        index = self.query_rules_tree.index(item)

        # 检查是否已经是第一个
        if index == 0:
            messagebox.showinfo("提示", "已经是第一个规则，无法上移")
            return

        # 交换规则位置
        self.query_rules[index], self.query_rules[index - 1] = self.query_rules[index - 1], self.query_rules[index]

        # 标记已修改
        self.app.mark_modified()

        # 更新显示
        self.refresh_query_rules_tree()

        # 重新选中移动后的规则
        self.query_rules_tree.selection_set(self.query_rules_tree.get_children()[index - 1])

    def move_query_rule_down(self):
        """下移选中的查询规则"""
        selected_items = self.query_rules_tree.selection()
        if not selected_items:
            messagebox.showwarning("选择错误", "请先选择要移动的规则")
            return

        if len(selected_items) > 1:
            messagebox.showwarning("选择错误", "只能选择一个规则进行移动")
            return

        item = selected_items[0]
        index = self.query_rules_tree.index(item)

        # 检查是否已经是最后一个
        if index == len(self.query_rules) - 1:
            messagebox.showinfo("提示", "已经是最后一个规则，无法下移")
            return

        # 交换规则位置
        self.query_rules[index], self.query_rules[index + 1] = self.query_rules[index + 1], self.query_rules[index]

        # 标记已修改
        self.app.mark_modified()

        # 更新显示
        self.refresh_query_rules_tree()

        # 重新选中移动后的规则
        self.query_rules_tree.selection_set(self.query_rules_tree.get_children()[index + 1])

    def add_query_rule(self):
        """添加查询条件规则"""
        # 获取输入值
        project = self.query_project_entry.get().strip()
        method = self.query_method_entry.get().strip()
        cancel_test = self.query_cancel_test_var.get()
        max_select = self.query_max_select_entry.get().strip()

        # 录入方式（来自下拉框，默认"方法"）
        input_method = self.query_mode_var.get() or "方法"
        sum_entry = self.query_sum_entry_var.get()  # 是否总和录入(勾选才走总和录入，否则按报告解析)

        # 验证输入 - 至少需要项目或检测方法中的一个
        if not project and not method:
            messagebox.showwarning("输入错误", "请至少填写项目或检测方法")
            return

        # 验证最大选中数量 - 留空表示不限制，否则必须是非负整数
        if max_select and not max_select.isdigit():
            messagebox.showwarning("输入错误", "最大选中数量必须为空或非负整数")
            return

        # 添加规则
        rule = {
            "project": project,
            "method": method,
            "cancel_test": cancel_test,
            "input_method": input_method,
            "max_select": max_select,
            "sum_entry": sum_entry
        }
        self.query_rules.append(rule)

        # 标记已修改
        self.app.mark_modified()

        # 清空输入框
        self.query_project_entry.delete(0, tk.END)
        self.query_method_entry.delete(0, tk.END)
        self.query_max_select_entry.delete(0, tk.END)
        self.query_cancel_test_var.set(False)
        self.query_mode_var.set("方法")
        self.query_sum_entry_var.set(False)

        # 更新显示
        self.refresh_query_rules_tree()

        messagebox.showinfo("成功", "查询规则已添加")

    def delete_query_rule(self):
        """删除选中的查询规则"""
        selected_items = self.query_rules_tree.selection()
        if not selected_items:
            messagebox.showwarning("选择错误", "请先选择要删除的规则")
            return

        # 确认删除
        if not messagebox.askyesno("确认删除", "确定要删除选中的查询规则吗？"):
            return

        # 从后往前删除，避免索引变化
        for item in reversed(selected_items):
            index = self.query_rules_tree.index(item)
            if 0 <= index < len(self.query_rules):
                del self.query_rules[index]

        # 标记已修改
        self.app.mark_modified()

        # 更新显示
        self.refresh_query_rules_tree()

        messagebox.showinfo("成功", "查询规则已删除")

    def refresh_query_rules_tree(self):
        """刷新查询规则显示"""
        # 清空现有规则显示
        for item in self.query_rules_tree.get_children():
            self.query_rules_tree.delete(item)

        # 添加规则到表格，包含序号
        for i, rule in enumerate(self.query_rules, 1):
            self.query_rules_tree.insert("", "end", values=(
                i,  # 序号
                rule.get("project", ""),
                rule.get("method", ""),
                "是" if rule.get("cancel_test", False) else "否",
                rule.get("max_select", ""),  # 最大选中数量
                rule.get("input_method", "方法"),  # 默认值为"方法"
                "是" if rule.get("sum_entry", False) else "否"  # 是否总和
            ))

    def get_rules(self):
        """获取查询规则"""
        def _iv(var, default):
            try:
                n = int((var.get() or "").strip())
                return n if n > 0 else default
            except (TypeError, ValueError, tk.TclError):
                return default
        return {
            "query_rules": self.query_rules,
            "exclusion_rules": self.exclusion_rules,
            "sample_parallel_threshold": _iv(self.sample_parallel_threshold_var, 20),
            "date_window_days": _iv(self.date_window_days_var, 30),
        }

    def set_rules(self, rules_data):
        """设置查询规则"""
        # 为旧数据设置默认的录入方式
        query_rules = rules_data.get("query_rules", [])
        for rule in query_rules:
            if "input_method" not in rule:
                rule["input_method"] = "方法"
            if "max_select" not in rule:
                rule["max_select"] = ""
            if "sum_entry" not in rule:
                rule["sum_entry"] = False
        self.query_rules = query_rules

        # 设置排除规则
        self.exclusion_rules = rules_data.get("exclusion_rules", [])

        # 查询策略(方法文件级，默认值兼容旧文件)
        self.sample_parallel_threshold_var.set(str(rules_data.get("sample_parallel_threshold", 20)))
        self.date_window_days_var.set(str(rules_data.get("date_window_days", 30)))

        self.refresh_query_rules_tree()


class MethodTab:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.switch_rules = []
        self.create_tab()
        self.update_rules_display()

    def create_tab(self):
        # 创建方法切换标签页
        method_frame = ttk.Frame(self.parent)
        self.parent.add(method_frame, text="方法切换")

        # 创建规则管理区域
        self.create_rules_management(method_frame)

        # 初始状态
        self.update_rules_display()

    def create_rules_management(self, parent):
        # 创建规则管理框架
        rules_frame = ttk.LabelFrame(parent, text="切换规则管理", padding=5)
        rules_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # 创建规则输入区域 - 分两行：第一行(项目名/文件名/试样描述)，第二行(原ID/目标ID/添加)
        input_frame = ttk.Frame(rules_frame)
        input_frame.pack(fill='x', pady=2)

        # 第一行：项目名 / 文件名 / 试样描述
        row0 = ttk.Frame(input_frame)
        row0.pack(fill='x', pady=1)
        row0.columnconfigure(5, weight=1)  # 末尾空白撑开

        ttk.Label(row0, text="项目名:").grid(row=0, column=0, padx=(0, 5), pady=1, sticky='w')
        self.project_name_entry = ttk.Entry(row0, width=12)
        self.project_name_entry.grid(row=0, column=1, padx=(0, 10), pady=1, sticky='w')
        _bind_entry_tooltip(self.project_name_entry)

        ttk.Label(row0, text="文件名:").grid(row=0, column=2, padx=(0, 5), pady=1, sticky='w')
        self.filename_entry = ttk.Entry(row0, width=12)
        self.filename_entry.grid(row=0, column=3, padx=(0, 10), pady=1, sticky='w')

        ttk.Label(row0, text="试样描述:").grid(row=0, column=4, padx=(0, 5), pady=1, sticky='w')
        self.desc_entry = ttk.Entry(row0, width=12)
        self.desc_entry.grid(row=0, column=5, padx=(0, 5), pady=1, sticky='ew')

        # 第二行：原ID / 目标ID / 添加按钮
        row1 = ttk.Frame(input_frame)
        row1.pack(fill='x', pady=1)
        row1.columnconfigure(4, weight=1)  # 末尾空白撑开，把添加按钮推到右侧

        ttk.Label(row1, text="原ID:").grid(row=0, column=0, padx=(0, 5), pady=1, sticky='w')
        self.from_id_entry = ttk.Entry(row1, width=10)
        self.from_id_entry.grid(row=0, column=1, padx=(0, 10), pady=1, sticky='w')

        ttk.Label(row1, text="目标ID:").grid(row=0, column=2, padx=(0, 5), pady=1, sticky='w')
        self.to_id_entry = ttk.Entry(row1, width=15)
        self.to_id_entry.grid(row=0, column=3, padx=(0, 5), pady=1, sticky='w')

        # 添加规则按钮
        self.add_rule_btn = ttkb.Button(row1, text="添加", command=self.add_rule, bootstyle="secondary")
        self.add_rule_btn.grid(row=0, column=5, padx=(5, 0), pady=1, sticky='e')

        # 创建规则显示区域
        display_frame = ttk.Frame(rules_frame)
        display_frame.pack(fill='both', expand=True, pady=5)

        # 创建规则表格（#0 列=序号, show="tree headings" 使其可见）
        columns = ("项目名", "文件名", "试样描述", "原ID", "目标ID")
        self.rules_tree = ttk.Treeview(display_frame, columns=columns, show="tree headings", height=6)

        # 设置 #0 树列（序号）
        self.rules_tree.heading("#0", text="NO", anchor="center")
        self.rules_tree.column("#0", width=45, anchor="center", stretch=False)

        # 设置列标题和宽度
        column_configs = {
            "项目名":  {"width": 180, "anchor": "w"},
            "文件名":  {"width": 80, "anchor": "w"},
            "试样描述": {"width": 80, "anchor": "w"},
            "原ID":    {"width": 40,  "anchor": "center"},
            "目标ID":  {"width": 40,  "anchor": "center"},
        }

        for col in columns:
            cfg = column_configs[col]
            self.rules_tree.heading(col, text=col, anchor=cfg["anchor"])
            self.rules_tree.column(col, width=cfg["width"], anchor=cfg["anchor"], stretch=True)

        # 添加滚动条（竖向 + 横向），用 grid 统一布局
        vsb = ttk.Scrollbar(display_frame, orient="vertical", command=self.rules_tree.yview)
        hsb = ttk.Scrollbar(display_frame, orient="horizontal", command=self.rules_tree.xview)
        self.rules_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        display_frame.columnconfigure(0, weight=1)
        display_frame.rowconfigure(0, weight=1)
        self.rules_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        # 绑定双击事件，用于编辑规则
        self.rules_tree.bind("<Double-1>", self.on_rule_double_click)

        # 项目名 列悬停显示全部内容
        _setup_cell_tooltip(self.rules_tree, ("项目名",))

        # 创建规则操作按钮
        button_frame = ttk.Frame(rules_frame)
        button_frame.pack(fill='x', pady=3)

        self.move_up_btn = ttkb.Button(button_frame, text="上移", command=self.move_rule_up, bootstyle="secondary")
        self.move_up_btn.pack(side='left', padx=2)

        self.move_down_btn = ttkb.Button(button_frame, text="下移", command=self.move_rule_down, bootstyle="secondary")
        self.move_down_btn.pack(side='left', padx=2)

        ttkb.Button(button_frame, text="删除", command=self.delete_rule, bootstyle="danger").pack(side='left', padx=2)

    def on_rule_double_click(self, event):
        """双击方法切换规则进行编辑"""
        item = self.rules_tree.selection()
        if not item:
            return

        item = item[0]
        column = self.rules_tree.identify_column(event.x)
        column_index = int(column.replace('#', '')) - 1  # 列索引从0开始

        # 序号列(#0)不可编辑
        if column_index < 0:
            return

        # 获取当前值
        current_values = self.rules_tree.item(item, 'values')
        current_value = current_values[column_index]

        # 获取单元格坐标
        x, y, width, height = self.rules_tree.bbox(item, column)

        # 创建编辑框
        entry = ttk.Entry(self.rules_tree)
        entry.place(x=x, y=y-4, width=width, height=height+8)
        entry.insert(0, current_value)
        entry.focus_set()

        # 列名映射到规则字段
        col_to_field = {
            "项目名": "project_name",
            "文件名": "filename",
            "试样描述": "desc",
            "原ID": "from_id",
            "目标ID": "to_id",
        }

        def save_edit(event=None):
            new_value = entry.get()

            # 更新显示
            new_values = list(current_values)
            new_values[column_index] = new_value
            self.rules_tree.item(item, values=new_values)

            # 更新数据
            index = self.rules_tree.index(item)
            column_name = self.rules_tree.heading(column_index)['text']
            field = col_to_field.get(column_name)

            if field and 0 <= index < len(self.switch_rules):
                self.switch_rules[index][field] = new_value

            # 标记已修改
            self.app.mark_modified()

            # 销毁编辑框
            entry.destroy()

        def cancel_edit(event=None):
            entry.destroy()

        # 绑定事件
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", cancel_edit)

    def move_rule_up(self):
        """上移选中的规则"""
        selected_items = self.rules_tree.selection()
        if not selected_items:
            messagebox.showwarning("选择错误", "请先选择要移动的规则")
            return

        if len(selected_items) > 1:
            messagebox.showwarning("选择错误", "只能选择一个规则进行移动")
            return

        item = selected_items[0]
        index = self.rules_tree.index(item)

        if index == 0:
            messagebox.showinfo("提示", "已经是第一个规则，无法上移")
            return

        self.switch_rules[index], self.switch_rules[index - 1] = self.switch_rules[index - 1], self.switch_rules[index]

        self.app.mark_modified()
        self.refresh_rules_tree()
        self.rules_tree.selection_set(self.rules_tree.get_children()[index - 1])

    def move_rule_down(self):
        """下移选中的规则"""
        selected_items = self.rules_tree.selection()
        if not selected_items:
            messagebox.showwarning("选择错误", "请先选择要移动的规则")
            return

        if len(selected_items) > 1:
            messagebox.showwarning("选择错误", "只能选择一个规则进行移动")
            return

        item = selected_items[0]
        index = self.rules_tree.index(item)

        if index == len(self.switch_rules) - 1:
            messagebox.showinfo("提示", "已经是最后一个规则，无法下移")
            return

        self.switch_rules[index], self.switch_rules[index + 1] = self.switch_rules[index + 1], self.switch_rules[index]

        self.app.mark_modified()
        self.refresh_rules_tree()
        self.rules_tree.selection_set(self.rules_tree.get_children()[index + 1])

    def update_rules_display(self):
        # 输入控件已在 create_rules_management 中布局(两行)，此处仅刷新规则表格
        self.refresh_rules_tree()

    def refresh_rules_tree(self):
        # 确保规则表格已创建
        if not hasattr(self, 'rules_tree'):
            return

        # 清空现有规则显示
        for item in self.rules_tree.get_children():
            self.rules_tree.delete(item)

        for i, rule in enumerate(self.switch_rules):
            self.rules_tree.insert("", "end", text=str(i + 1), values=(
                rule.get("project_name", ""),
                rule.get("filename", ""),
                rule.get("desc", ""),
                rule["from_id"],
                rule["to_id"]
            ))

    def add_rule(self):
        # 获取输入值
        from_id = self.from_id_entry.get().strip()
        to_id = self.to_id_entry.get().strip()
        project_name = self.project_name_entry.get().strip()
        filename = self.filename_entry.get().strip()
        desc = self.desc_entry.get().strip()

        if not from_id or not to_id:
            messagebox.showwarning("输入错误", "请填写原方法ID和目标方法ID")
            return

        rule = {
            "from_id": from_id,
            "to_id": to_id,
            "project_name": project_name,
            "filename": filename,
            "desc": desc
        }

        self.switch_rules.append(rule)

        # 标记已修改
        self.app.mark_modified()

        # 清空输入框
        self.from_id_entry.delete(0, tk.END)
        self.to_id_entry.delete(0, tk.END)
        self.project_name_entry.delete(0, tk.END)
        self.filename_entry.delete(0, tk.END)
        self.desc_entry.delete(0, tk.END)

        # 更新显示
        self.refresh_rules_tree()

    def delete_rule(self):
        # 确保规则表格已创建
        if not hasattr(self, 'rules_tree'):
            return

        # 获取选中的规则
        selected_items = self.rules_tree.selection()
        if not selected_items:
            messagebox.showwarning("选择错误", "请先选择要删除的规则")
            return

        # 确认删除
        if not messagebox.askyesno("确认删除", "确定要删除选中的规则吗？"):
            return

        # 收集要删除的规则索引（从后往前删避免索引偏移）
        indices_to_delete = []
        for item in selected_items:
            idx = self.rules_tree.index(item)
            indices_to_delete.append(idx)

        for idx in sorted(indices_to_delete, reverse=True):
            if 0 <= idx < len(self.switch_rules):
                del self.switch_rules[idx]

        # 标记已修改
        self.app.mark_modified()

        # 更新显示
        self.refresh_rules_tree()

    def get_rules(self):
        """获取方法切换规则"""
        return {"switch_rules": self.switch_rules}

    def set_rules(self, switch_rules):
        """设置方法切换规则"""
        self.switch_rules = switch_rules
        self.refresh_rules_tree()


class WeighingTab:
    # 称样方式码 ↔ 编辑器显示标签（条件称样规则表用）
    _WEIGHING_MODE_OPTIONS = [
        ("record", "称量记录"),
        ("random", "随机数生成"),
        ("process", "称量记录处理"),
        ("none", "无需称样量"),
        ("pdf", "PDF报告"),
    ]
    _LABEL_BY_MODE = {code: lbl for code, lbl in _WEIGHING_MODE_OPTIONS}
    _MODE_BY_LABEL = {lbl: code for code, lbl in _WEIGHING_MODE_OPTIONS}

    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.processing_rules = []
        self.weighing_rules = []  # 条件称样规则：[{filename,project_name,desc,weighing_mode}]
        self.weighing_params = {
            "decimal_places": "",
            "min_value": "",
            "max_value": "",
            "conversion_factor": "",
            "result_decimal_places": "",
            "weighing_mode": "random",
            "non_parallel_suffixes": [],
            "weighing_share_group": "",
            "single_weighing": False,
            "writeback_excel": False
        }
        self.create_tab()

    def create_tab(self):
        """创建称样量标签页"""
        weighing_frame = ttk.Frame(self.parent)
        self.parent.add(weighing_frame, text="称样量")

        # 创建称样量模式选择
        mode_frame = ttk.LabelFrame(weighing_frame, text="称样量模式", padding=5)
        mode_frame.pack(fill='x', padx=5, pady=5)

        self.weighing_mode = tk.StringVar(value="random")

        ttkb.Radiobutton(mode_frame, text="称量记录",
                        variable=self.weighing_mode, value="record",
                        command=self.on_weighing_mode_change, bootstyle="primary").pack(side='left', padx=10)
        ttkb.Radiobutton(mode_frame, text="随机数生成",
                        variable=self.weighing_mode, value="random",
                        command=self.on_weighing_mode_change, bootstyle="primary").pack(side='left', padx=10)
        ttkb.Radiobutton(mode_frame, text="称量记录处理",
                        variable=self.weighing_mode, value="process",
                        command=self.on_weighing_mode_change, bootstyle="primary").pack(side='left', padx=10)
        ttkb.Radiobutton(mode_frame, text="无需称样量",
                        variable=self.weighing_mode, value="none",
                        command=self.on_weighing_mode_change, bootstyle="primary").pack(side='left', padx=10)
        ttkb.Radiobutton(mode_frame, text="PDF报告",
                        variable=self.weighing_mode, value="pdf",
                        command=self.on_weighing_mode_change, bootstyle="primary").pack(side='left', padx=10)
        ttkb.Radiobutton(mode_frame, text="条件称样",
                        variable=self.weighing_mode, value="conditional",
                        command=self.on_weighing_mode_change, bootstyle="primary").pack(side='left', padx=10)

        # 平行样合并 + 称样量共享组 左右并排
        merge_share_row = ttk.Frame(weighing_frame)
        merge_share_row.pack(fill='x', padx=5, pady=(0, 5))
        # 平行样合并设置：不并入平行样的标记后缀字母(如 M=基体加标)
        parallel_frame = ttk.LabelFrame(merge_share_row, text="平行样合并", padding=5)
        parallel_frame.pack(side='left', fill='both', expand=True, padx=(0, 8))
        nps_input = ttk.Frame(parallel_frame)
        nps_input.pack(fill='x')
        ttk.Label(nps_input, text="非平行样标记后缀:").pack(side='left', padx=(0, 5))
        self.non_parallel_suffixes = tk.StringVar(value="")
        ttk.Entry(nps_input, textvariable=self.non_parallel_suffixes, width=16).pack(side='left', padx=(0, 8))
        # 单称样多次进样：一个称样量复用于各平行槽(如 TDI 称1次进样2次，LIMS 多平行槽时不再报"称量记录平行不足")
        self.single_weighing = tk.BooleanVar(value=False)
        ttk.Checkbutton(parallel_frame, text=" 单称样多次进样",
                        variable=self.single_weighing, style=self.app.large_cb_style).pack(fill='x', pady=(4, 0))

        # 称样量共享组：同名组的方法跨行共享称样缓存(同一样品首方法生成、后方法复用)
        share_frame = ttk.LabelFrame(merge_share_row, text="称样量共享组", padding=5)
        share_frame.pack(side='left', fill='both', expand=True)
        sf = ttk.Frame(share_frame)
        sf.pack(fill='x')
        ttk.Label(sf, text="共享组名:").pack(side='left', padx=(0, 5))
        self.weighing_share_group = tk.StringVar(value="")
        ttk.Entry(sf, textvariable=self.weighing_share_group, width=24).pack(side='left', padx=(0, 8))

        # 创建内容区域 - 所有模式的内容都显示
        self.weighing_content_frame = ttk.Frame(weighing_frame)
        self.weighing_content_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # 创建所有模式的内容区域
        self.create_all_weighing_modes()

        # 条件称样规则表（初始隐藏，仅 conditional 时显示）
        self.create_weighing_rules_section(weighing_frame)

        # 初始状态设置
        self.on_weighing_mode_change()

    def create_all_weighing_modes(self):
        """创建所有称样量模式的内容区域"""
        content_frame = self.weighing_content_frame

        # 随机数生成模式
        self.random_frame = ttk.LabelFrame(content_frame, text="随机数生成模式", padding=5)
        self.random_frame.pack(fill='x', pady=5)

        # 参数设置区域
        param_frame = ttk.Frame(self.random_frame)
        param_frame.pack(fill='x', pady=2)

        # 小数位数 - 默认为空
        decimal_frame = ttk.Frame(param_frame)
        decimal_frame.pack(fill='x', pady=2)

        ttk.Label(decimal_frame, text="小数位数:").pack(side='left', padx=(0, 5))
        self.decimal_places = tk.StringVar(value="")
        self.decimal_entry = ttk.Entry(decimal_frame, textvariable=self.decimal_places, width=5)
        self.decimal_entry.pack(side='left', padx=(0, 5))  # 添加这行

        # 最小值 - 默认为空
        ttk.Label(decimal_frame, text="最小值:").pack(side='left', padx=(0, 5))
        self.min_value = tk.StringVar(value="")
        self.min_entry = ttk.Entry(decimal_frame, textvariable=self.min_value, width=8)
        self.min_entry.pack(side='left', padx=(0, 5))  # 添加这行

        # 最大值 - 默认为空
        ttk.Label(decimal_frame, text="最大值:").pack(side='left', padx=(0, 5))
        self.max_value = tk.StringVar(value="")
        self.max_entry = ttk.Entry(decimal_frame, textvariable=self.max_value, width=8)
        self.max_entry.pack(side='left', padx=(0, 5))  # 添加这行

        # 回写称量记录Excel：random 模式生成的称样量回填到称样量空格（序列运行时），与小数位同行
        self.writeback_excel = tk.BooleanVar(value=False)
        ttk.Checkbutton(decimal_frame, text=" 回写称量记录Excel",
                        variable=self.writeback_excel, style=self.app.large_cb_style).pack(side='left', padx=(10, 0))

        # 称量记录处理模式
        self.process_frame = ttk.LabelFrame(content_frame, text="称量记录处理模式", padding=5)
        self.process_frame.pack(fill='both', expand=True, pady=5)

        # 处理规则配置区域 - 使用grid布局
        config_frame = ttk.Frame(self.process_frame)
        config_frame.pack(fill='x', pady=5)

        # 配置列的权重，使添加按钮可以右对齐
        config_frame.columnconfigure(0, weight=0)  # 检测方法标签
        config_frame.columnconfigure(1, weight=0)  # 检测方法输入框
        config_frame.columnconfigure(2, weight=0)  # 处理类型标签
        config_frame.columnconfigure(3, weight=0)  # 处理类型下拉框
        config_frame.columnconfigure(4, weight=0)  # 换算因子标签
        config_frame.columnconfigure(5, weight=0)  # 换算因子输入框
        config_frame.columnconfigure(6, weight=0)  # 位数标签
        config_frame.columnconfigure(7, weight=0)  # 位数输入框
        config_frame.columnconfigure(8, weight=1)  # 空白区域（用于右对齐按钮）
        config_frame.columnconfigure(9, weight=0)  # 添加按钮

        # 检测方法
        ttk.Label(config_frame, text="检测方法:").grid(row=0, column=0, padx=(0, 5), pady=2, sticky='w')
        self.processing_method = tk.StringVar()
        self.processing_method_entry = ttk.Entry(config_frame, textvariable=self.processing_method, width=12)
        self.processing_method_entry.grid(row=0, column=1, padx=(0, 10), pady=2, sticky='w')

        # 处理类型 - 添加"换算加补充"选项
        ttk.Label(config_frame, text="处理类型:").grid(row=0, column=2, padx=(0, 5), pady=2, sticky='w')
        self.processing_type = tk.StringVar(value="直接读取")
        processing_type_combo = ttk.Combobox(config_frame, textvariable=self.processing_type,
                                             values=["直接读取", "小数位补充", "换算处理", "换算加补充"],
                                             state="readonly", width=10)
        processing_type_combo.grid(row=0, column=3, padx=(0, 10), pady=2, sticky='w')

        # 添加这行 - 绑定处理类型改变事件
        processing_type_combo.bind("<<ComboboxSelected>>", self.on_processing_type_change)

        # 换算因子标签和输入框 - 创建但不立即放置到网格中
        self.conversion_factor_label = ttk.Label(config_frame, text="换算因子:")
        self.conversion_factor = tk.StringVar(value="")
        self.conversion_factor_entry = ttk.Entry(config_frame, textvariable=self.conversion_factor, width=6)

        # 位数标签和输入框 - 改为普通输入框，允许自由输入
        self.result_decimal_label = ttk.Label(config_frame, text="位数:")
        self.result_decimal_places = tk.StringVar(value="")
        self.result_decimal_entry = ttk.Entry(config_frame, textvariable=self.result_decimal_places, width=5)

        # 添加按钮 - 右对齐
        self.add_processing_rule_btn = ttkb.Button(config_frame, text="添加", command=self.add_processing_rule, bootstyle="secondary")
        self.add_processing_rule_btn.grid(row=0, column=9, padx=(10, 0), pady=2, sticky='e')

        # 初始显示参数配置
        self.show_processing_params()

        # 处理规则列表
        rules_list_frame = ttk.Frame(self.process_frame)
        rules_list_frame.pack(fill='both', expand=True, pady=5)

        columns = ("检测方法", "处理类型", "换算因子", "结果小数位数")
        self.processing_rules_tree = ttk.Treeview(rules_list_frame, columns=columns, show="headings", height=3)

        column_configs = {
            "检测方法": {"width": 150, "anchor": "w"},
            "处理类型": {"width": 100, "anchor": "center"},
            "换算因子": {"width": 80, "anchor": "center"},
            "结果小数位数": {"width": 100, "anchor": "center"}
        }

        for col in columns:
            config = column_configs.get(col, {})
            anchor = config.get("anchor", "w")
            self.processing_rules_tree.heading(col, text=col, anchor=anchor)
            self.processing_rules_tree.column(col, width=config.get("width", 100),
                                              anchor=anchor)

        rules_scrollbar = ttk.Scrollbar(rules_list_frame, orient="vertical", command=self.processing_rules_tree.yview)
        self.processing_rules_tree.configure(yscrollcommand=rules_scrollbar.set)

        self.processing_rules_tree.pack(side="left", fill="both", expand=True)
        rules_scrollbar.pack(side="right", fill="y")

        # 绑定双击事件，用于编辑处理规则
        self.processing_rules_tree.bind("<Double-1>", self.on_processing_rule_double_click)

        # 处理规则操作按钮 - 只保留删除按钮，移除执行处理按钮
        rules_button_frame = ttk.Frame(self.process_frame)
        rules_button_frame.pack(fill='x', pady=2)

        self.delete_processing_rule_btn = ttkb.Button(rules_button_frame, text="删除选中",
                                                     command=self.delete_processing_rule, bootstyle="danger")
        self.delete_processing_rule_btn.pack(side='left', padx=5)

    def show_processing_params(self):
        """显示处理类型对应的参数配置"""
        processing_type = self.processing_type.get()

        # 先移除所有参数控件（如果已经存在）
        try:
            self.conversion_factor_label.grid_remove()
            self.conversion_factor_entry.grid_remove()
            self.result_decimal_label.grid_remove()
            self.result_decimal_entry.grid_remove()
        except:
            pass  # 忽略可能的异常

        if processing_type == "换算处理" or processing_type == "换算加补充":
            # 换算处理和换算加补充需要显示换算因子和位数
            self.conversion_factor_label.grid(row=0, column=4, padx=(0, 5), pady=2, sticky='w')
            self.conversion_factor_entry.grid(row=0, column=5, padx=(0, 10), pady=2, sticky='w')
            self.result_decimal_label.grid(row=0, column=6, padx=(0, 5), pady=2, sticky='w')
            self.result_decimal_entry.grid(row=0, column=7, padx=(0, 10), pady=2, sticky='w')

    def on_processing_rule_double_click(self, event):
        """双击处理规则进行编辑"""
        item = self.processing_rules_tree.selection()
        if not item:
            return

        item = item[0]
        column = self.processing_rules_tree.identify_column(event.x)
        column_index = int(column.replace('#', '')) - 1  # 列索引从0开始

        # 获取当前值
        current_values = self.processing_rules_tree.item(item, 'values')
        current_value = current_values[column_index]

        # 获取单元格坐标
        x, y, width, height = self.processing_rules_tree.bbox(item, column)

        # 创建编辑框
        entry = ttk.Entry(self.processing_rules_tree)
        entry.place(x=x, y=y-4, width=width, height=height+8)
        entry.insert(0, current_value)
        entry.focus_set()

        def save_edit(event=None):
            # 获取新值
            new_value = entry.get()

            # 更新显示
            new_values = list(current_values)
            new_values[column_index] = new_value
            self.processing_rules_tree.item(item, values=new_values)

            # 更新数据
            index = self.processing_rules_tree.index(item)
            if 0 <= index < len(self.processing_rules):
                rule = self.processing_rules[index]
                column_name = self.processing_rules_tree.heading(column_index)['text']

                if column_name == "检测方法":
                    rule["method"] = new_value
                elif column_name == "处理类型":
                    rule["type"] = new_value
                elif column_name == "换算因子":
                    rule["factor"] = new_value
                elif column_name == "结果小数位数":
                    rule["decimal_places"] = new_value

            # 标记已修改
            self.app.mark_modified()

            # 销毁编辑框
            entry.destroy()

        def cancel_edit(event=None):
            entry.destroy()

        # 绑定事件
        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", cancel_edit)

    def on_processing_type_change(self, event=None):
        """处理类型改变时的回调"""
        self.show_processing_params()

    def on_weighing_mode_change(self):
        """称样量模式改变时的处理"""
        # 设置所有框架的状态
        mode = self.weighing_mode.get()
        # 条件称样规则表：仅 conditional 时显示
        self.weighing_rules_frame.pack_forget()
        if mode == "conditional":
            self.weighing_rules_frame.pack(fill='both', expand=True, pady=5)
        if mode == "record":
            # 称量记录模式 - 两个区域都禁用
            self.set_frame_state(self.random_frame, "disabled")
            self.set_frame_state(self.process_frame, "disabled")
        elif mode == "random":
            # 随机数生成模式 - 随机数区域启用，处理区域禁用
            self.set_frame_state(self.random_frame, "normal")
            self.set_frame_state(self.process_frame, "disabled")
            # 自动生成随机数
            self.generate_random_numbers()
        elif mode == "process":
            # 处理模式 - 随机数区域禁用，处理区域启用
            self.set_frame_state(self.random_frame, "disabled")
            self.set_frame_state(self.process_frame, "normal")
        elif mode == "none":
            # 无需称样量 - 两个区域都禁用
            self.set_frame_state(self.random_frame, "disabled")
            self.set_frame_state(self.process_frame, "disabled")
        elif mode == "pdf":
            # PDF报告 - 称样量取自报告(样品初始质量)；仅「小数位数」可编辑(按其保留末尾0，如0.3100)，min/max 与处理区禁用
            self.set_frame_state(self.random_frame, "disabled")
            self.set_frame_state(self.process_frame, "disabled")
            self.decimal_entry.configure(state="normal")
        elif mode == "conditional":
            # 条件称样：不同样品按规则命中不同方式，随机/处理参数都可能被引用 → 两者都启用
            self.set_frame_state(self.random_frame, "normal")
            self.set_frame_state(self.process_frame, "normal")

    def set_frame_state(self, frame, state):
        """设置框架及其子组件的状态"""
        # 递归设置所有子组件的状态
        for child in frame.winfo_children():
            self.set_widget_state(child, state)

    def set_widget_state(self, widget, state):
        """设置单个组件的状态"""
        if isinstance(widget, (ttk.Entry, ttk.Spinbox, ttk.Button, ttk.Combobox)):
            widget.configure(state=state)
        elif isinstance(widget, ttk.Treeview):
            # Treeview没有直接的state属性，但可以禁用选择
            if state == "disabled":
                widget.configure(selectmode="none")
            else:
                widget.configure(selectmode="browse")
        elif isinstance(widget, ttk.Frame) or isinstance(widget, ttk.LabelFrame):
            # 递归设置子组件状态
            for child in widget.winfo_children():
                self.set_widget_state(child, state)

    def generate_random_numbers(self):
        """生成随机数"""
        try:
            # 处理空值，使用默认值
            decimal_places_str = self.decimal_places.get()
            min_val_str = self.min_value.get()
            max_val_str = self.max_value.get()

            decimal_places = int(decimal_places_str) if decimal_places_str else 2
            min_val = float(min_val_str) if min_val_str else 0.1
            max_val = float(max_val_str) if max_val_str else 1.0

            if min_val >= max_val:
                return  # 不显示错误消息，只是不生成

            # 生成10个随机数
            random_numbers = []
            for _ in range(10):
                num = random.uniform(min_val, max_val)
                formatted_num = f"{num:.{decimal_places}f}"
                random_numbers.append(formatted_num)

            # 显示随机数
            result_text = "\n".join(random_numbers)
            # 这里可以添加显示逻辑，例如更新标签或文本框

        except Exception:
            pass  # 静默处理错误

    def add_processing_rule(self):
        """添加或更新处理规则"""
        method = self.processing_method.get().strip()
        processing_type = self.processing_type.get()

        # 构建规则数据
        rule = {
            "method": method,
            "type": processing_type
        }

        # 根据处理类型添加相应参数
        if processing_type == "换算处理" or processing_type == "换算加补充":
            factor_str = self.conversion_factor.get().strip()
            decimal_str = self.result_decimal_places.get().strip()

            if not factor_str:
                messagebox.showwarning("输入错误", "请输入换算因子")
                return

            if not decimal_str:
                messagebox.showwarning("输入错误", "请输入位数")
                return

            try:
                factor = float(factor_str)
                decimal_places = int(decimal_str)
            except ValueError:
                messagebox.showwarning("输入错误", "换算因子和位数必须是数字")
                return

            rule["factor"] = factor
            rule["decimal_places"] = decimal_places

        # 检查是否已存在该检测方法
        for i, existing_rule in enumerate(self.processing_rules):
            if existing_rule["method"] == method:
                # 更新现有记录
                self.processing_rules[i] = rule
                self.refresh_processing_rules_tree()
                # 标记已修改
                self.app.mark_modified()
                messagebox.showinfo("成功", "处理规则已更新")
                return

        # 添加新记录
        self.processing_rules.append(rule)

        # 标记已修改
        self.app.mark_modified()

        # 清空输入框
        self.processing_method.set("")
        self.conversion_factor.set("")
        self.result_decimal_places.set("")

        # 更新显示
        self.refresh_processing_rules_tree()
        messagebox.showinfo("成功", "处理规则已添加")

    def delete_processing_rule(self):
        """删除选中的处理规则"""
        selected_items = self.processing_rules_tree.selection()
        if not selected_items:
            messagebox.showwarning("选择错误", "请先选择要删除的处理规则")
            return

        if messagebox.askyesno("确认删除", "确定要删除选中的处理规则吗？"):
            for item in selected_items:
                index = self.processing_rules_tree.index(item)
                if 0 <= index < len(self.processing_rules):
                    del self.processing_rules[index]

            # 标记已修改
            self.app.mark_modified()

            # 更新显示
            self.refresh_processing_rules_tree()

    def refresh_processing_rules_tree(self):
        """刷新处理规则显示"""
        # 清空现有规则显示
        for item in self.processing_rules_tree.get_children():
            self.processing_rules_tree.delete(item)

        # 添加规则到表格
        for rule in self.processing_rules:
            # 根据处理类型设置显示值
            processing_type_display = rule["type"]

            # 设置换算因子和位数的显示值
            factor_display = rule.get("factor", "")
            decimal_places_display = rule.get("decimal_places", "")

            self.processing_rules_tree.insert("", "end", values=(
                rule["method"],
                processing_type_display,
                factor_display,
                decimal_places_display
            ))

    def create_weighing_rules_section(self, parent):
        """条件称样规则表（仅 weighing_mode=conditional 时显示）：
        按 文件名/项目名/试样描述 关键词命中 → 用对应称样方式。首条命中即用，末条留空三条件=默认方式。"""
        self.weighing_rules_frame = ttk.LabelFrame(
            parent, text="条件称样规则", padding=5)
        # 初始不 pack，由 on_weighing_mode_change 控制显隐
        wr_input = ttk.Frame(self.weighing_rules_frame)
        wr_input.pack(fill='x', pady=2)
        ttk.Label(wr_input, text="文件名:").pack(side='left', padx=(0, 3))
        self.wr_filename_var = tk.StringVar()
        ttk.Entry(wr_input, textvariable=self.wr_filename_var, width=8).pack(side='left', padx=(0, 6))
        ttk.Label(wr_input, text="项目名:").pack(side='left', padx=(0, 3))
        self.wr_project_var = tk.StringVar()
        ttk.Entry(wr_input, textvariable=self.wr_project_var, width=12).pack(side='left', padx=(0, 6))
        ttk.Label(wr_input, text="试样描述:").pack(side='left', padx=(0, 3))
        self.wr_desc_var = tk.StringVar()
        ttk.Entry(wr_input, textvariable=self.wr_desc_var, width=12).pack(side='left', padx=(0, 6))
        ttk.Label(wr_input, text="称样方式:").pack(side='left', padx=(0, 3))
        self.wr_mode_var = tk.StringVar()
        ttk.Combobox(wr_input, textvariable=self.wr_mode_var,
                     values=[lbl for _, lbl in self._WEIGHING_MODE_OPTIONS],
                     state="readonly", width=12).pack(side='left', padx=(0, 6))
        ttkb.Button(wr_input, text="添加", command=self.add_weighing_rule, bootstyle="secondary").pack(side='left')
        # 表格
        wr_list = ttk.Frame(self.weighing_rules_frame)
        wr_list.pack(fill='both', expand=True, pady=3)
        wcols = ("序号", "文件名关键词", "项目名关键词", "试样描述关键词", "称样方式")
        self.wr_tree = ttk.Treeview(wr_list, columns=wcols, show="headings", height=3)
        for c in wcols:
            self.wr_tree.heading(c, text=c, anchor='w')
        self.wr_tree.column("序号", width=40, anchor='w')
        self.wr_tree.column("文件名关键词", width=90, anchor='w')
        self.wr_tree.column("项目名关键词", width=110, anchor='w')
        self.wr_tree.column("试样描述关键词", width=110, anchor='w')
        self.wr_tree.column("称样方式", width=120, anchor='w')
        self.wr_tree.pack(side='left', fill='both', expand=True)
        wr_scroll = ttk.Scrollbar(wr_list, orient="vertical", command=self.wr_tree.yview)
        wr_scroll.pack(side='right', fill='y')
        self.wr_tree.configure(yscrollcommand=wr_scroll.set)
        self.wr_tree.bind("<Double-1>", self.on_weighing_rule_double_click)
        # 操作按钮
        wr_btn = ttk.Frame(self.weighing_rules_frame)
        wr_btn.pack(fill='x', pady=2)
        ttkb.Button(wr_btn, text="删除选中", command=self.delete_weighing_rule, bootstyle="danger").pack(side='left', padx=2)
        ttkb.Button(wr_btn, text="上移", command=lambda: self.move_weighing_rule(-1), bootstyle="secondary").pack(side='left', padx=2)
        ttkb.Button(wr_btn, text="下移", command=lambda: self.move_weighing_rule(1), bootstyle="secondary").pack(side='left', padx=2)

    # ---------- 条件称样规则 CRUD ----------

    def add_weighing_rule(self):
        """添加一条条件称样规则"""
        mode_code = self._MODE_BY_LABEL.get(self.wr_mode_var.get().strip(), "")
        if not mode_code:
            messagebox.showwarning("输入错误", "请选择称样方式")
            return
        self.weighing_rules.append({
            "filename": self.wr_filename_var.get().strip(),
            "project_name": self.wr_project_var.get().strip(),
            "desc": self.wr_desc_var.get().strip(),
            "weighing_mode": mode_code
        })
        self.refresh_weighing_rules_tree()
        self.wr_filename_var.set(""); self.wr_project_var.set(""); self.wr_desc_var.set(""); self.wr_mode_var.set("")
        self.app.mark_modified()

    def delete_weighing_rule(self):
        """删除选中的条件称样规则"""
        sel = self.wr_tree.selection()
        if not sel:
            messagebox.showwarning("选择错误", "请先选择要删除的规则")
            return
        idx = int(sel[0]) - 1
        if 0 <= idx < len(self.weighing_rules):
            del self.weighing_rules[idx]
            self.refresh_weighing_rules_tree()
            self.app.mark_modified()

    def move_weighing_rule(self, delta):
        """上移(delta=-1)或下移(delta=1)选中的规则（顺序即优先级）"""
        sel = self.wr_tree.selection()
        if not sel:
            return
        i = int(sel[0]) - 1
        if not (0 <= i < len(self.weighing_rules)):
            return
        j = i + delta
        if not (0 <= j < len(self.weighing_rules)):
            return
        self.weighing_rules[i], self.weighing_rules[j] = self.weighing_rules[j], self.weighing_rules[i]
        self.refresh_weighing_rules_tree()
        for item in self.wr_tree.get_children():
            if int(item) == j + 1:
                self.wr_tree.selection_set(item)
                break
        self.app.mark_modified()

    def refresh_weighing_rules_tree(self):
        """重建条件称样规则表格（称样方式列显示标签）"""
        for item in self.wr_tree.get_children():
            self.wr_tree.delete(item)
        for i, r in enumerate(self.weighing_rules):
            self.wr_tree.insert("", "end", iid=str(i + 1),
                                values=(str(i + 1), r.get("filename", ""),
                                        r.get("project_name", ""), r.get("desc", ""),
                                        self._LABEL_BY_MODE.get(r.get("weighing_mode", ""),
                                                                 r.get("weighing_mode", ""))))

    def on_weighing_rule_double_click(self, event):
        """双击条件称样规则单元格就地编辑（称样方式列用 Combobox，其余用 Entry）"""
        sel = self.wr_tree.selection()
        if not sel:
            return
        item = sel[0]
        column_index = int(self.wr_tree.identify_column(event.x).replace('#', '')) - 1
        col_key = {1: "filename", 2: "project_name", 3: "desc", 4: "weighing_mode"}.get(column_index)
        if col_key is None:  # 序号列或越界
            return
        current_values = self.wr_tree.item(item, 'values')
        current_value = current_values[column_index]
        x, y, width, height = self.wr_tree.bbox(item, self.wr_tree.identify_column(event.x))

        if col_key == "weighing_mode":
            combo = ttk.Combobox(self.wr_tree, values=[lbl for _, lbl in self._WEIGHING_MODE_OPTIONS],
                                 state="readonly")
            combo.place(x=x, y=y - 4, width=width, height=height + 8)
            combo.set(current_value)
            combo.focus_set()

            def save_mode(event=None):
                if not combo.winfo_exists():
                    return
                lbl = combo.get()
                code = self._MODE_BY_LABEL.get(lbl, "")
                if code:
                    new_values = list(current_values)
                    new_values[column_index] = lbl
                    self.wr_tree.item(item, values=new_values)
                    idx = int(item) - 1
                    if 0 <= idx < len(self.weighing_rules):
                        self.weighing_rules[idx]["weighing_mode"] = code
                        self.app.mark_modified()
                combo.destroy()

            def cancel_mode(event=None):
                if combo.winfo_exists():
                    combo.destroy()

            combo.bind("<<ComboboxSelected>>", save_mode)
            combo.bind("<FocusOut>", save_mode)
            combo.bind("<Escape>", cancel_mode)
            return

        entry = ttk.Entry(self.wr_tree)
        entry.place(x=x, y=y - 4, width=width, height=height + 8)
        entry.insert(0, current_value)
        entry.focus_set()

        def save_edit(event=None):
            if not entry.winfo_exists():
                return
            new_value = entry.get()
            new_values = list(current_values)
            new_values[column_index] = new_value
            self.wr_tree.item(item, values=new_values)
            idx = int(item) - 1
            if 0 <= idx < len(self.weighing_rules):
                self.weighing_rules[idx][col_key] = new_value
            self.app.mark_modified()
            entry.destroy()

        def cancel_edit(event=None):
            if entry.winfo_exists():
                entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", cancel_edit)

    def get_rules_and_params(self):
        """获取处理规则和称样量参数"""
        # 非平行样标记后缀：逗号分隔(中英文) → 大写字母列表
        non_par = [p.strip().upper() for p in self.non_parallel_suffixes.get().replace("，", ",").split(",") if p.strip()]
        # 更新称样量参数
        self.weighing_params = {
            "decimal_places": self.decimal_places.get(),
            "min_value": self.min_value.get(),
            "max_value": self.max_value.get(),
            "conversion_factor": self.conversion_factor.get(),
            "result_decimal_places": self.result_decimal_places.get(),
            "weighing_mode": self.weighing_mode.get(),
            "non_parallel_suffixes": non_par,
            "weighing_share_group": self.weighing_share_group.get().strip(),
            "single_weighing": self.single_weighing.get(),
            "writeback_excel": self.writeback_excel.get(),
            "weighing_rules": self.weighing_rules
        }

        return {
            "processing_rules": self.processing_rules,
            "weighing_params": self.weighing_params
        }

    def set_rules_and_params(self, processing_rules, weighing_params):
        """设置处理规则和称样量参数"""
        self.processing_rules = processing_rules
        self.weighing_params = weighing_params

        # 设置称样量参数
        self.decimal_places.set(weighing_params.get("decimal_places", ""))
        self.min_value.set(weighing_params.get("min_value", ""))
        self.max_value.set(weighing_params.get("max_value", ""))
        self.conversion_factor.set(weighing_params.get("conversion_factor", ""))
        self.result_decimal_places.set(weighing_params.get("result_decimal_places", ""))
        self.non_parallel_suffixes.set(
            ", ".join(str(s).strip().upper() for s in (weighing_params.get("non_parallel_suffixes") or [])))
        self.weighing_share_group.set(weighing_params.get("weighing_share_group", ""))
        self.single_weighing.set(bool(weighing_params.get("single_weighing", False)))
        self.writeback_excel.set(bool(weighing_params.get("writeback_excel", False)))

        # 条件称样规则
        self.weighing_rules = list(weighing_params.get("weighing_rules") or [])
        self.refresh_weighing_rules_tree()

        # 设置称样量模式
        saved_weighing_mode = weighing_params.get("weighing_mode", "random")
        self.weighing_mode.set(saved_weighing_mode)

        # 更新显示
        self.refresh_processing_rules_tree()
        self.on_weighing_mode_change()


class SpectrumUploadTab:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.upload_mode = "local_upload"  # 默认模式为本地上传
        # 新增谱图检查参数
        self.spectrum_check_params = {
            "blank": {"enabled": False, "count": "", "keyword": ""},
            "standard": {"enabled": False, "count": "", "keyword": ""},
            "linearity": {"enabled": False, "count": "", "keyword": ""},
            "sample": {"enabled": False, "count": "", "keyword": ""}
        }
        # 新增参数
        self.instrument_type = ""  # 仪器类型(GCMS/GC/LC/LCMSMS/XRF/ICP)
        self.report_parse_enabled = False  # 报告解析启用开关（仪器类型）
        self.clear_spectrum = False  # 新增：录入前是否清空谱图
        self.spectrum_filter_rules = []  # 按项目筛选谱图文件名 [{project, filter}]
        self.create_tab()

    def create_tab(self):
        """创建谱图上传标签页"""
        spectrum_frame = ttk.Frame(self.parent)
        self.parent.add(spectrum_frame, text="谱图获取")

        # 创建模式选择区域
        self.create_mode_selection(spectrum_frame)

        # 创建内容区域
        self.create_content_area(spectrum_frame)

    def create_mode_selection(self, parent):
        """创建上传模式选择区域"""
        mode_frame = ttk.LabelFrame(parent, text="获取方式", padding=0)
        mode_frame.pack(fill='x', padx=5, pady=5)

        # 创建单选按钮容器，使用水平布局
        mode_selection_frame = ttk.Frame(mode_frame)
        mode_selection_frame.pack(fill='x', pady=2)

        # 模式选择变量 - 使用同一个变量确保互斥
        self.mode_var = tk.StringVar(value="local_upload")

        # 本地上传单选按钮
        self.local_upload_radio = ttkb.Radiobutton(
            mode_selection_frame,
            text="本地上传",
            variable=self.mode_var,
            value="local_upload",
            command=self.on_mode_change, bootstyle="primary")
        self.local_upload_radio.pack(side='left', padx=(0, 20))

        # 无需谱图（与"本地上传"互斥）
        self.no_spectrum_radio = ttkb.Radiobutton(
            mode_selection_frame,
            text="无需谱图",
            variable=self.mode_var,
            value="no_spectrum",
            command=self.on_mode_change, bootstyle="primary")
        self.no_spectrum_radio.pack(side='left', padx=(20, 0))

        # 是否清空谱图：选中后序列运行录入数据前调 deleteSpectrumByProjectIds 删除谱图再录入
        self.clear_spectrum_var = tk.BooleanVar(value=self.clear_spectrum)
        self.clear_spectrum_check = ttk.Checkbutton(
            mode_selection_frame,
            text="录入前清空谱图",
            variable=self.clear_spectrum_var,
            command=self.on_clear_spectrum_change, style=self.app.large_cb_style)
        self.clear_spectrum_check.pack(side='left', padx=(30, 0))

    def create_content_area(self, parent):
        """创建内容区域"""
        self.content_frame = ttk.Frame(parent)
        self.content_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # 创建本地上传模式内容
        self.create_local_upload_content()

        # 初始显示
        self.on_mode_change()

    def create_local_upload_content(self):
        """创建本地上传模式内容"""
        self.local_upload_frame = ttk.LabelFrame(self.content_frame, text="本地参数设置", padding=0)
        self.local_upload_frame.pack(fill='both', expand=True)  # 始终显示

        # 创建垂直布局，确保所有内容左侧对齐
        main_content_frame = ttk.Frame(self.local_upload_frame)
        main_content_frame.pack(fill='both', expand=True, pady=2, padx=4)

        # 谱图检查区域 - 放在顶部
        self.create_spectrum_check_section(main_content_frame)

        # 谱图分流区域 - 按项目筛选谱图(同编号不同谱图)，复选框默认隐藏
        sf_toggle_frame = ttk.Frame(main_content_frame)
        sf_toggle_frame.pack(fill='x', padx=5, pady=(2, 0))
        self.spf_enable_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(sf_toggle_frame, text=" 启用谱图分流（按项目）",
                        variable=self.spf_enable_var,
                        style=self.app.large_cb_style,
                        command=self.toggle_spectrum_filter).pack(side='left')
        self.create_spectrum_filter_section(main_content_frame)
        self.spf_frame.pack_forget()

        # 新增参数设置区域 - 放在底部
        self.create_additional_params_section(main_content_frame)

    def create_spectrum_check_section(self, parent):
        """创建谱图检查设置区域（双列紧凑布局：空白+线性 / 标液+样品）"""
        check_frame = ttk.LabelFrame(parent, text="谱图检查", padding=3)
        check_frame.pack(fill='x', pady=(0, 5))

        left = ttk.Frame(check_frame)
        left.pack(side='left', fill='x', expand=True)
        right = ttk.Frame(check_frame)
        right.pack(side='left', fill='x', expand=True, padx=(10, 0))

        self._make_spec_row(left, "blank", "空白")
        self._make_spec_row(left, "linearity", "线性")
        self._make_spec_row(right, "standard", "标液")
        self._make_spec_row(right, "sample", "样品")

    def create_spectrum_filter_section(self, parent):
        """谱图分流：同一样品编号下不同项目用不同谱图时，按项目名筛选谱图文件名。"""
        self.spf_frame = ttk.LabelFrame(parent, text="谱图分流（按项目）", padding=5)
        self.spf_frame.pack(fill='x', pady=(0, 5))
        frame = self.spf_frame

        # 添加行：项目 + 谱图筛选 + 添加
        row = ttk.Frame(frame)
        row.pack(fill='x', pady=(0, 4))
        row.columnconfigure(1, weight=1)
        ttk.Label(row, text="项目:").grid(row=0, column=0, padx=(0, 4))
        self.spf_project_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.spf_project_var).grid(row=0, column=1, sticky='ew', padx=(0, 8))
        ttk.Label(row, text="谱图筛选:").grid(row=0, column=2, padx=(0, 4))
        self.spf_filter_var = tk.StringVar()
        ttk.Entry(row, textvariable=self.spf_filter_var, width=12).grid(row=0, column=3, padx=(0, 8))
        ttkb.Button(row, text="添加", command=self.add_spectrum_filter_rule, bootstyle="secondary").grid(row=0, column=4)

        # 列表
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill='x')
        self.spf_tree = ttk.Treeview(list_frame, columns=("项目", "谱图筛选"), show="headings", height=3)
        self.spf_tree.heading("项目", text="项目", anchor="w")
        self.spf_tree.column("项目", width=340, anchor="w")
        self.spf_tree.heading("谱图筛选", text="谱图筛选", anchor="w")
        self.spf_tree.column("谱图筛选", width=140, anchor="w")
        sb = ttk.Scrollbar(list_frame, orient="vertical", command=self.spf_tree.yview)
        self.spf_tree.configure(yscrollcommand=sb.set)
        self.spf_tree.pack(side="left", fill="x", expand=True)
        sb.pack(side="right", fill="y")
        self.spf_tree.bind("<Double-1>", self.on_spf_double_click)

        # 操作按钮 + 说明
        btn = ttk.Frame(frame)
        btn.pack(fill='x', pady=(4, 0))
        ttkb.Button(btn, text="删除", command=self.delete_spectrum_filter_rule, bootstyle="danger").pack(side='left')
        ttk.Label(btn, text="筛选按文件名子串(大小写不敏感)：M=只传含M的；!M=排除含M的；留空=全传",
                  foreground="#888").pack(side='left', padx=12)

        self.refresh_spf_tree()

    def refresh_spf_tree(self):
        """刷新谱图分流列表"""
        for item in self.spf_tree.get_children():
            self.spf_tree.delete(item)
        for r in self.spectrum_filter_rules:
            self.spf_tree.insert("", "end", values=(r.get("project", ""), r.get("filter", "")))

    def toggle_spectrum_filter(self):
        """复选框控制谱图分流区域的显示/隐藏"""
        if self.spf_enable_var.get():
            self.spf_frame.pack(fill='x', pady=(0, 5), before=self.report_parse_toggle_frame)
        else:
            self.spf_frame.pack_forget()

    def add_spectrum_filter_rule(self):
        """添加一条谱图分流规则"""
        project = self.spf_project_var.get().strip()
        if not project:
            messagebox.showwarning("输入错误", "请填写项目名")
            return
        self.spectrum_filter_rules.append({"project": project, "filter": self.spf_filter_var.get().strip()})
        self.spf_project_var.set("")
        self.spf_filter_var.set("")
        self.refresh_spf_tree()
        self.app.mark_modified()

    def delete_spectrum_filter_rule(self):
        """删除选中的谱图分流规则"""
        selected = self.spf_tree.selection()
        if not selected:
            messagebox.showwarning("选择错误", "请先选择要删除的规则")
            return
        for item in reversed(selected):
            index = self.spf_tree.index(item)
            if 0 <= index < len(self.spectrum_filter_rules):
                del self.spectrum_filter_rules[index]
        self.refresh_spf_tree()
        self.app.mark_modified()

    def on_spf_double_click(self, event):
        """双击编辑谱图分流规则"""
        selected = self.spf_tree.selection()
        if not selected:
            return
        item = selected[0]
        column = self.spf_tree.identify_column(event.x)
        column_index = int(column.replace('#', '')) - 1
        if column_index < 0:
            return
        current_values = self.spf_tree.item(item, 'values')
        current_value = current_values[column_index] if column_index < len(current_values) else ""
        x, y, width, height = self.spf_tree.bbox(item, column)

        entry = ttk.Entry(self.spf_tree)
        entry.place(x=x, y=y - 4, width=width, height=height + 8)
        entry.insert(0, current_value)
        entry.focus_set()

        def save_edit(event=None):
            if not entry.winfo_exists():
                return
            new_value = entry.get()
            new_values = list(current_values)
            new_values[column_index] = new_value
            self.spf_tree.item(item, values=new_values)
            index = self.spf_tree.index(item)
            if 0 <= index < len(self.spectrum_filter_rules):
                key = "project" if column_index == 0 else "filter"
                self.spectrum_filter_rules[index][key] = new_value.strip() if key == "filter" else new_value
            self.app.mark_modified()
            entry.destroy()

        def cancel_edit(event=None):
            entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", cancel_edit)

    def create_additional_params_section(self, parent):
        """创建报告解析设置区域"""
        # 启用复选框 - 单独一行（样式同「启用谱图分流」），选中后下方三项才可编辑
        self.report_parse_toggle_frame = ttk.Frame(parent)
        self.report_parse_toggle_frame.pack(fill='x', padx=5, pady=(14, 0))
        self.report_parse_enabled_var = tk.BooleanVar(value=self.report_parse_enabled)
        ttk.Checkbutton(self.report_parse_toggle_frame, text=" 启用报告解析",
                        variable=self.report_parse_enabled_var,
                        style=self.app.large_cb_style,
                        command=self.on_report_parse_enabled_change).pack(side='left')

        # 报告解析参数（无单独标题框，置于「启用报告解析」下方）
        params_frame = ttk.Frame(parent)
        params_frame.pack(fill='x', pady=0)

        # 使用网格布局实现输入框两端对齐
        params_frame.columnconfigure(0, weight=0)  # 标签列，不扩展
        params_frame.columnconfigure(1, weight=1)  # 输入框列，占据剩余空间

        # 仪器类型 - 第一行
        ttk.Label(params_frame, text="仪器类型:").grid(row=0, column=0, padx=(5, 5), pady=2, sticky='w')
        self.instrument_type_var = tk.StringVar(value=self.instrument_type)
        self.instrument_type_combo = ttk.Combobox(params_frame, textvariable=self.instrument_type_var,
                                                  values=["GCMS", "GC", "LC", "LCMSMS", "XRF", "ICP"],
                                                  state="readonly")
        self.instrument_type_combo.grid(row=0, column=1, padx=(0, 10), pady=2, sticky='w')
        self.instrument_type_combo.bind('<<ComboboxSelected>>', self.on_instrument_type_change)

    def _make_spec_row(self, parent, param_type, label):
        """单行谱图检查：复选框 + 数+数量框 + 键+关键字框（水平排列）"""
        row = ttk.Frame(parent)
        row.pack(fill='x', pady=1)

        check_var = tk.BooleanVar(value=self.spectrum_check_params[param_type]["enabled"])
        cb = ttk.Checkbutton(row, text=label, variable=check_var,
                              command=lambda pt=param_type, cv=check_var: self.on_spectrum_check_change(pt, cv),
                              style=self.app.large_cb_style)
        cb.pack(side='left')

        ttk.Label(row, text="数").pack(side='left', padx=(4, 0))
        count_var = tk.StringVar(value=self.spectrum_check_params[param_type]["count"])
        ce = ttk.Entry(row, textvariable=count_var, width=4)
        ce.pack(side='left', padx=(2, 0))
        ce.bind('<KeyRelease>', lambda e, pt=param_type, cv=count_var:
                self.on_spectrum_param_change(pt, 'count', cv.get()))

        ttk.Label(row, text="键").pack(side='left', padx=(4, 0))
        keyword_var = tk.StringVar(value=self.spectrum_check_params[param_type]["keyword"])
        ke = ttk.Entry(row, textvariable=keyword_var, width=14)
        ke.pack(side='left', padx=(2, 0), fill='x', expand=True)
        ke.bind('<KeyRelease>', lambda e, pt=param_type, kv=keyword_var:
                self.on_spectrum_param_change(pt, 'keyword', kv.get()))

        setattr(self, f"{param_type}_check_var", check_var)
        setattr(self, f"{param_type}_count_var", count_var)
        setattr(self, f"{param_type}_keyword_var", keyword_var)

    def on_spectrum_check_change(self, param_type, check_var):
        """谱图检查复选框状态改变"""
        self.spectrum_check_params[param_type]["enabled"] = check_var.get()
        self.app.mark_modified()

    def on_spectrum_param_change(self, param_type, param_name, value):
        """谱图检查参数改变"""
        self.spectrum_check_params[param_type][param_name] = value
        self.app.mark_modified()

    def on_instrument_type_change(self, event=None):
        """仪器类型改变"""
        self.instrument_type = self.instrument_type_var.get()
        self.app.mark_modified()

    def on_report_parse_enabled_change(self):
        """报告解析启用复选框状态改变"""
        self.report_parse_enabled = self.report_parse_enabled_var.get()
        self._update_report_parse_inputs_state()
        self.app.mark_modified()

    def _update_report_parse_inputs_state(self):
        """仪器类型 仅在「非无需谱图 且 勾选启用」时可编辑。"""
        active = (self.mode_var.get() != "no_spectrum") and self.report_parse_enabled_var.get()
        self.instrument_type_combo.config(state="readonly" if active else "disabled")

    def on_clear_spectrum_change(self):
        """录入前清空谱图复选框状态改变"""
        self.clear_spectrum = self.clear_spectrum_var.get()
        self.app.mark_modified()

    def validate_spectrum_check_params(self):
        """验证谱图检查参数"""
        errors = []

        # 样品类型按样品编号前缀匹配，关键字可空(blank/standard/linearity 仍需关键字)
        for param_type, display_name in [("blank", "空白"), ("standard", "标液"),
                                         ("linearity", "线性")]:
            if self.spectrum_check_params[param_type]["enabled"]:
                count = self.spectrum_check_params[param_type]["count"].strip()
                keyword = self.spectrum_check_params[param_type]["keyword"].strip()

                if not count and not keyword:
                    errors.append(f"{display_name}谱图检查已启用，但数量和关键字都为空")

        return errors

    def on_mode_change(self):
        """模式改变时的处理"""
        mode = self.mode_var.get()

        if mode == "no_spectrum":
            # 无需谱图：禁用全部参数
            self.set_local_upload_frame_state("disabled")
            self.local_upload_frame.configure(text="本地参数设置(禁用)")
        else:
            # 本地上传：启用全部参数（含报告解析：仪器类型）
            self.set_local_upload_frame_state("normal")
            self.local_upload_frame.configure(text="本地参数设置")

        # 报告解析三项再按复选框门控
        self._update_report_parse_inputs_state()

        # 标记已修改
        self.app.mark_modified()

    def set_local_upload_frame_state(self, state):
        """设置本地上传模式区域内所有组件的状态"""

        # 递归设置框架内所有组件的状态
        def set_widget_state(widget, state):
            if isinstance(widget, (ttk.Entry, ttk.Checkbutton)):
                widget.configure(state=state)
            elif isinstance(widget, ttk.Frame) or isinstance(widget, ttk.LabelFrame):
                for child in widget.winfo_children():
                    set_widget_state(child, state)

        # 设置本地上传模式区域内所有组件的状态
        for child in self.local_upload_frame.winfo_children():
            set_widget_state(child, state)

    def get_settings(self):
        """获取谱图上传设置"""
        return {
            "upload_mode": self.mode_var.get(),
            "spectrum_check_params": self.spectrum_check_params,
            "instrument_type": self.instrument_type,  # 仪器类型(GCMS/GC/LC/LCMSMS/XRF/ICP)
            "clear_spectrum": self.clear_spectrum,  # 新增：录入前是否清空谱图
            "spectrum_filter_rules": self.spectrum_filter_rules,  # 按项目筛选谱图
            "report_parse_enabled": self.report_parse_enabled  # 报告解析启用开关
        }

    def set_settings(self, settings):
        """设置谱图上传配置"""
        if not settings:
            return

        # 设置模式
        if "upload_mode" in settings:
            self.mode_var.set(settings["upload_mode"])

        # 设置谱图检查参数
        if "spectrum_check_params" in settings:
            check_params = settings["spectrum_check_params"]
            for param_type in ["blank", "standard", "linearity", "sample"]:
                if param_type in check_params:
                    params = check_params[param_type]
                    # 设置复选框状态
                    check_var = getattr(self, f"{param_type}_check_var", None)
                    if check_var:
                        en = params.get("enabled", False)
                        check_var.set(en)
                        self.spectrum_check_params[param_type]["enabled"] = en

                    # 设置数量
                    count_var = getattr(self, f"{param_type}_count_var", None)
                    if count_var:
                        cv = params.get("count", "")
                        count_var.set(cv)
                        self.spectrum_check_params[param_type]["count"] = cv

                    # 设置关键字
                    keyword_var = getattr(self, f"{param_type}_keyword_var", None)
                    if keyword_var:
                        kw = params.get("keyword", "")
                        keyword_var.set(kw)
                        self.spectrum_check_params[param_type]["keyword"] = kw

        # 设置仪器类型
        if "instrument_type" in settings:
            self.instrument_type = settings["instrument_type"]
            self.instrument_type_var.set(self.instrument_type)

        # 设置报告解析启用开关
        if "report_parse_enabled" in settings:
            self.report_parse_enabled = bool(settings["report_parse_enabled"])
            self.report_parse_enabled_var.set(self.report_parse_enabled)

        # 新增：设置录入前清空谱图
        if "clear_spectrum" in settings:
            self.clear_spectrum = bool(settings["clear_spectrum"])
            self.clear_spectrum_var.set(self.clear_spectrum)

        # 新增：设置谱图分流规则（按项目筛选谱图）
        if "spectrum_filter_rules" in settings:
            self.spectrum_filter_rules = [
                {"project": str(r.get("project", "")), "filter": str(r.get("filter", ""))}
                for r in (settings["spectrum_filter_rules"] or []) if isinstance(r, dict)
            ]
            self.refresh_spf_tree()
            if self.spectrum_filter_rules:
                self.spf_enable_var.set(True)
                self.spf_frame.pack(fill='x', pady=(0, 5), before=self.report_parse_toggle_frame)

        # 更新显示
        self.on_mode_change()


class QueryAppFixed:
    def __init__(self, root):
        self.root = root
        self.root.title("录入方法编辑器_未加载配置文件")
        self.root.geometry("1400x1100")  # 增大窗口以完整显示各标签页内容

        # 统一控件风格（与 SequenceMaster 保持一致）
        style = ttk.Style()
        # 字号统一为菜单栏（文件/操作）字体大小
        _mf = tkfont.nametofont("TkMenuFont").actual()
        _font = (_mf["family"], _mf["size"])
        for _s in ("TLabel", "TButton", "TEntry", "TCombobox", "TCheckbutton",
                   "TRadiobutton", "Treeview.Heading", "TLabelframe.Label"):
            style.configure(_s, font=_font)
        style.configure("Treeview", font=_font, rowheight=36)  # 行高足够、字号统一
        # 选中行蓝底白字：默认主题下 selected 配色不明显，看不出选了哪行
        style.map("Treeview",
                  background=[("selected", "#2563eb")],
                  foreground=[("selected", "#ffffff")])
        # 标签页配色：未选中灰字、选中蓝字白底，更醒目
        style.configure("TNotebook.Tab", padding=(24, 10),
                        font=(_font[0], _font[1], "bold"), foreground="#6b7280")
        style.map("TNotebook.Tab",
                  foreground=[("selected", "#2563eb")],
                  background=[("selected", "#ffffff")])
        style.configure("TButton", padding=(8, 2))          # 收紧按钮，避免偏高偏大
        self.large_cb_style = _make_large_cb_style(1.3) or "TCheckbutton"  # 放大勾选框样式
        # 禁用态更明显：灰底 + 浅灰字（默认仅变字色、底色仍白，不易区分）
        for _e in ("TEntry", "TCombobox"):
            style.map(_e,
                      fieldbackground=[("disabled", "#e8e8e6")],
                      foreground=[("disabled", "#9aa0a6")])

        # 当前配置文件路径和名称
        self.current_config_file = None
        self.config_file_name = "未加载配置文件"

        # 标记是否有未保存的修改
        self.modified = False

        # 记录上次打开的文件路径
        self.last_config_file = os.path.join(paths.data_dir(), "last_config.txt")

        # 标记初始化是否完成
        self.initialization_complete = False

        # 创建标签页控件
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)

        # 创建五个标签页
        self.query_tab = QueryTab(self.notebook, self)
        self.method_tab = MethodTab(self.notebook, self)
        self.other_params_tab = OtherParamsTab(self.notebook, self)
        self.other_params_tab.create_standard_tab()          # 标液设备
        self.weighing_tab = WeighingTab(self.notebook, self)  # 称样量
        self.spectrum_tab = SpectrumUploadTab(self.notebook, self)  # 谱图获取
        self.other_params_tab.create_submit_tab()            # 其他参数(固定参数+提交签名)

        # 创建菜单栏（置于标签页上方）
        self.create_menubar()

        # 绑定事件处理
        self.bind_events()

        # 静默加载保存的规则（不显示弹窗）
        self.silent_load_rules()

        # 标记初始化完成
        self.initialization_complete = True

    def bind_events(self):
        """绑定事件处理，修复焦点问题"""
        # 绑定笔记本切换事件
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        # 绑定鼠标点击事件，确保输入框可以获得焦点
        self.root.bind("<Button-1>", self.ensure_focus)

    def on_tab_changed(self, event):
        """当标签页切换时，确保当前标签页的输入框可以获得焦点"""
        current_tab = self.notebook.select()
        if current_tab:
            # 强制更新焦点
            self.root.update_idletasks()

    def ensure_focus(self, event):
        """确保点击的widget可以获得焦点"""
        widget = event.widget
        if isinstance(widget, (ttk.Entry, tk.Entry)):
            widget.focus_set()

    def mark_modified(self):
        """标记配置已修改"""
        # 只有在初始化完成后才标记为已修改
        if hasattr(self, 'initialization_complete') and self.initialization_complete:
            if not self.modified:
                self.modified = True
                self.update_window_title()

    def update_window_title(self):
        """更新窗口标题显示当前配置文件（不显示文件后缀）"""
        if self.current_config_file:
            # 去掉文件后缀显示
            base_name = os.path.splitext(self.config_file_name)[0]
            title = f"录入方法编辑器_{base_name}"
            if self.modified:
                title += "*"
            self.root.title(title)
        else:
            self.root.title("录入方法编辑器_未加载配置文件")

    def save_last_config_path(self):
        """保存最后打开的配置文件路径"""
        if self.current_config_file:
            try:
                with open(self.last_config_file, 'w', encoding='utf-8') as f:
                    f.write(self.current_config_file)
            except Exception as e:
                print(f"保存最后配置文件路径失败: {str(e)}")

    def load_last_config_path(self):
        """加载最后打开的配置文件路径"""
        try:
            if os.path.exists(self.last_config_file):
                with open(self.last_config_file, 'r', encoding='utf-8') as f:
                    path = f.read().strip()
                    if os.path.exists(path):
                        return path
        except Exception as e:
            print(f"加载最后配置文件路径失败: {str(e)}")
        return None

    def convert_old_rules(self, default_rules, filename_rules):
        """将旧的 default_rules + filename_rules 合并为统一 switch_rules 格式"""
        switch_rules = []
        # 旧 default_rules（仅 project_name 匹配）
        for rule in (default_rules or []):
            switch_rules.append({
                "from_id": rule.get("from_id", ""),
                "to_id": rule.get("to_id", ""),
                "project_name": rule.get("project_name", ""),
                "filename": "",
                "desc": ""
            })
        # 旧 filename_rules（filename + desc + project_name 匹配）
        for rule in (filename_rules or []):
            new_rule = {
                "from_id": rule.get("from_id", ""),
                "to_id": rule.get("to_id", ""),
                "project_name": rule.get("project_name", ""),
                "filename": rule.get("filename", rule.get("name", "")),
                "desc": rule.get("desc", "")
            }
            # 旧格式 name_type=="项目名" 时 name 实际是项目名
            if rule.get("name_type") == "项目名":
                new_rule["project_name"] = rule.get("name", "")
                new_rule["filename"] = ""
            switch_rules.append(new_rule)
        return switch_rules

    def silent_load_rules(self):
        """静默加载规则，不显示任何弹窗"""
        try:
            # 首先尝试加载上次打开的配置文件
            last_config = self.load_last_config_path()
            if last_config and os.path.exists(last_config):
                self.current_config_file = last_config
                self.config_file_name = os.path.basename(last_config)

                rules_data = load_method(last_config)

                # 加载方法切换规则（优先新格式，回退旧格式）
                switch_rules = rules_data.get("switch_rules")
                if switch_rules is None:
                    default_rules = rules_data.get("default_rules", [])
                    filename_rules = rules_data.get("filename_rules", [])
                    switch_rules = self.convert_old_rules(default_rules, filename_rules)

                self.method_tab.set_rules(switch_rules)

                # 加载查询规则
                query_rules = rules_data.get("query_rules", [])
                self.query_tab.set_rules(query_rules)

                # 加载称样量规则和参数
                processing_rules = rules_data.get("processing_rules", [])
                weighing_params = rules_data.get("weighing_params", {})
                self.weighing_tab.set_rules_and_params(processing_rules, weighing_params)

                # 加载谱图上传设置
                spectrum_settings = rules_data.get("spectrum_upload_settings", {})
                self.spectrum_tab.set_settings(spectrum_settings)

                # 加载其他参数设置
                other_params_settings = rules_data.get("other_params_settings", {})
                self.other_params_tab.set_settings(other_params_settings)

                # 静默更新显示，不弹出消息
                self.update_window_title()

                self.modified = False  # 加载后重置修改标记
                return

            # 如果没有上次打开的配置文件，尝试加载默认配置文件
            default_config = "switch_rules.mtd"
            if os.path.exists(default_config):
                self.current_config_file = default_config
                self.config_file_name = os.path.basename(default_config)

                rules_data = load_method(default_config)

                # 加载方法切换规则（优先新格式，回退旧格式）
                switch_rules = rules_data.get("switch_rules")
                if switch_rules is None:
                    default_rules = rules_data.get("default_rules", [])
                    filename_rules = rules_data.get("filename_rules", [])
                    switch_rules = self.convert_old_rules(default_rules, filename_rules)

                self.method_tab.set_rules(switch_rules)

                # 加载查询规则
                query_rules = rules_data.get("query_rules", [])
                self.query_tab.set_rules(query_rules)

                # 加载称样量规则和参数
                processing_rules = rules_data.get("processing_rules", [])
                weighing_params = rules_data.get("weighing_params", {})
                self.weighing_tab.set_rules_and_params(processing_rules, weighing_params)

                # 加载谱图上传设置
                spectrum_settings = rules_data.get("spectrum_upload_settings", {})
                self.spectrum_tab.set_settings(spectrum_settings)

                # 加载其他参数设置
                other_params_settings = rules_data.get("other_params_settings", {})
                self.other_params_tab.set_settings(other_params_settings)

                # 静默更新显示，不弹出消息
                self.update_window_title()

                self.modified = False  # 加载后重置修改标记
        except Exception as e:
            print(f"静默加载规则失败: {str(e)}")  # 仅在控制台输出错误

    def load_config_file(self, file_path):
        """加载指定路径的方法文件并填充各标签页（供 SequenceMaster 等外部入口调用）"""
        if not file_path or not os.path.exists(file_path):
            return False
        try:
            rules_data = load_method(file_path)
        except Exception as e:
            messagebox.showerror("错误", f"加载方法文件失败:\n{e}")
            return False
        self.current_config_file = file_path
        self.config_file_name = os.path.basename(file_path)
        # ponytail: 填充逻辑与 silent_load_rules 一致，未抽取共用以免改动正常工作的代码
        switch_rules = rules_data.get("switch_rules")
        if switch_rules is None:
            default_rules = rules_data.get("default_rules", [])
            filename_rules = rules_data.get("filename_rules", [])
            switch_rules = self.convert_old_rules(default_rules, filename_rules)
        self.method_tab.set_rules(switch_rules)
        self.query_tab.set_rules(rules_data.get("query_rules", []))
        self.weighing_tab.set_rules_and_params(
            rules_data.get("processing_rules", []), rules_data.get("weighing_params", {}))
        self.spectrum_tab.set_settings(rules_data.get("spectrum_upload_settings", {}))
        self.other_params_tab.set_settings(rules_data.get("other_params_settings", {}))
        self.update_window_title()
        self.modified = False
        self.update_window_title()  # 重置标记后刷新标题(填充控件时 trace 触发 mark_modified 会先显示*)
        return True

    def create_menubar(self):
        """创建菜单栏（加载/保存/另存为/重置/查询），置于标签页上方，配色贴合主题"""
        c = ttkb.Style().colors  # sandstone-light 主题色
        menu_opts = dict(
            bg=c.bg, fg=c.fg,
            activebackground=c.light, activeforeground=c.fg,
            borderwidth=0, relief="flat",
        )
        menubar = tk.Menu(self.root, **menu_opts)

        # 文件菜单：新建、加载、保存、另存为
        file_menu = tk.Menu(menubar, tearoff=False, **menu_opts)
        file_menu.add_command(label="新建", command=self.new_file)
        file_menu.add_command(label="加载", command=self.load_rules)
        file_menu.add_command(label="保存", command=self.save_all_rules)
        file_menu.add_command(label="另存为", command=self.save_as_rules)
        menubar.add_cascade(label="文件", menu=file_menu)

        # 操作菜单：重置、查询
        action_menu = tk.Menu(menubar, tearoff=False, **menu_opts)
        action_menu.add_command(label="重置", command=self.reset_form)
        action_menu.add_command(label="查询", command=self.execute_query)
        menubar.add_cascade(label="操作", menu=action_menu)

        self.root.config(menu=menubar)

    def new_file(self):
        """新建空白配置文件：销毁现有标签页并重建，清空所有规则与设置，恢复到未加载状态"""
        if self.modified:
            if not messagebox.askyesno("确认新建", "当前有未保存的修改，确定要新建吗？"):
                return
        # 销毁现有标签页
        for tab_id in self.notebook.tabs():
            self.notebook.forget(tab_id)
        # 重建各标签页（构造时以默认空值初始化）
        self.query_tab = QueryTab(self.notebook, self)
        self.method_tab = MethodTab(self.notebook, self)
        self.other_params_tab = OtherParamsTab(self.notebook, self)
        self.other_params_tab.create_standard_tab()          # 标液设备
        self.weighing_tab = WeighingTab(self.notebook, self)  # 称样量
        self.spectrum_tab = SpectrumUploadTab(self.notebook, self)  # 谱图获取
        self.other_params_tab.create_submit_tab()            # 其他参数(固定参数+提交签名)
        # 重置文件状态
        self.current_config_file = None
        self.config_file_name = "未加载配置文件"
        self.modified = False
        self.update_window_title()
        self.notebook.select(0)

    def execute_query(self):
        # 获取查询条件
        project = self.query_tab.query_project_entry.get()
        method = self.query_tab.query_method_entry.get()
        cancel_test = self.query_tab.query_cancel_test_var.get()

        # 验证输入 - 至少需要项目或检测方法中的一个
        if not project and not method:
            messagebox.showwarning("输入错误", "请至少填写项目或检测方法")
            return

        # 在实际应用中，这里会执行查询逻辑
        query_result = (
            f"查询条件:\n"
            f"项目: {project}\n"
            f"检测方法: {method}\n"
            f"Retest: {'是' if cancel_test else '否'}"
        )
        messagebox.showinfo("查询结果", query_result)

    def reset_form(self):
        # 重置所有输入
        self.method_tab.from_id_entry.delete(0, tk.END)
        self.method_tab.to_id_entry.delete(0, tk.END)
        self.method_tab.project_name_entry.delete(0, tk.END)
        self.method_tab.filename_entry.delete(0, tk.END)
        self.method_tab.desc_entry.delete(0, tk.END)
        self.query_tab.query_project_entry.delete(0, tk.END)
        self.query_tab.query_method_entry.delete(0, tk.END)
        self.query_tab.query_cancel_test_var.set(False)
        self.query_tab.query_max_select_entry.delete(0, tk.END)
        self.query_tab.query_mode_var.set("方法")
        self.method_tab.update_rules_display()

    def save_all_rules(self):
        """保存所有规则到当前配置文件（.mtd 二进制格式）"""
        # 验证谱图检查参数
        spectrum_errors = self.spectrum_tab.validate_spectrum_check_params()
        if spectrum_errors:
            error_message = "谱图检查参数错误:\n" + "\n".join(spectrum_errors)
            messagebox.showwarning("保存失败", error_message)
            return

        if not self.current_config_file:
            # 如果没有当前配置文件，使用默认名称
            self.current_config_file = "switch_rules.mtd"
            self.config_file_name = "switch_rules.mtd"

        # 获取各标签页的规则
        method_rules = self.method_tab.get_rules()
        query_rules = self.query_tab.get_rules()
        weighing_data = self.weighing_tab.get_rules_and_params()
        spectrum_settings = self.spectrum_tab.get_settings()
        other_params_settings = self.other_params_tab.get_settings()  # 新增其他参数设置

        rules_data = {
            "switch_rules": method_rules["switch_rules"],
            "query_rules": query_rules,
            "processing_rules": weighing_data["processing_rules"],
            "weighing_params": weighing_data["weighing_params"],
            "spectrum_upload_settings": spectrum_settings,
            "other_params_settings": other_params_settings  # 新增其他参数设置
        }

        try:
            save_method(self.current_config_file, rules_data)

            # 保存后重置修改标记
            self.modified = False
            self.update_window_title()
            self.save_last_config_path()
        except Exception as e:
            messagebox.showerror("保存失败", f"保存规则时出错: {str(e)}")

    def save_as_rules(self):
        """另存所有规则到新的配置文件（.mtd 二进制格式）"""
        file_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="另存配置文件",
            defaultextension=".mtd",
            filetypes=[("方法文件", "*.mtd"), ("All files", "*.*")]
        )

        if not file_path:
            return  # 用户取消了保存

        self.current_config_file = file_path
        self.config_file_name = os.path.basename(file_path)
        self.save_all_rules()

    def load_rules(self):
        """从文件加载规则（.mtd 二进制格式）"""
        file_path = filedialog.askopenfilename(
            parent=self.root,
            title="选择配置文件",
            defaultextension=".mtd",
            filetypes=[("方法文件", "*.mtd"), ("All files", "*.*")]
        )

        if not file_path:
            return  # 用户取消了选择

        try:
            rules_data = load_method(file_path)

            # 加载方法切换规则（优先新格式，回退旧格式）
            switch_rules = rules_data.get("switch_rules")
            if switch_rules is None:
                default_rules = rules_data.get("default_rules", [])
                filename_rules = rules_data.get("filename_rules", [])
                switch_rules = self.convert_old_rules(default_rules, filename_rules)

            self.method_tab.set_rules(switch_rules)

            # 加载查询规则
            query_rules = rules_data.get("query_rules", [])
            self.query_tab.set_rules(query_rules)

            # 加载称样量规则和参数
            processing_rules = rules_data.get("processing_rules", [])
            weighing_params = rules_data.get("weighing_params", {})
            self.weighing_tab.set_rules_and_params(processing_rules, weighing_params)

            # 加载谱图上传设置
            spectrum_settings = rules_data.get("spectrum_upload_settings", {})
            self.spectrum_tab.set_settings(spectrum_settings)

            # 加载其他参数设置
            other_params_settings = rules_data.get("other_params_settings", {})
            self.other_params_tab.set_settings(other_params_settings)

            # 更新当前配置文件信息
            self.current_config_file = file_path
            self.config_file_name = os.path.basename(file_path)

            # 更新显示
            self.update_window_title()

            # 重置修改标记
            self.modified = False
            self.update_window_title()  # 重置后刷新标题(填充控件时 trace 触发 mark_modified 会先显示*)

            # 保存最后打开的配置文件路径
            self.save_last_config_path()

        except Exception as e:
            messagebox.showerror("加载失败", f"加载规则时出错: {str(e)}")


class OtherParamsTab:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.standard_type = "fresh"  # 默认标液类型：现配现用
        self.preparation_number = ""  # 配制序号
        self.instrument_setting = "default"  # 默认仪器设置：使用立方默认设备
        self.device_number = ""  # 设备编号
        self.equipment_rules = []  # 条件设备匹配规则: [{filename, desc, project_name, device_number}, ...]
        self.standard_rules = []  # 条件标液匹配规则: [{filename, desc, project_name, preparation_number}, ...]
        self.fixed_params = []  # 固定参数规则：[{trigger, param_name, param_value}]
        self.lab_proc_rules = []  # 条件实验过程规则: [{filename, desc, project_name, lab_proc}]，命中用 lab_proc 覆盖 experimentProcess
        # 不在 __init__ 建标签页：由 app 调 create_standard_tab / create_submit_tab 分两步建，
        # 以便「标液设备」与「其他参数」之间插入称样量/谱图获取(标签页顺序见 app 建页处)。

    def create_standard_tab(self):
        """建「标液设备」标签页(标液类型+仪器设备)。"""
        std_frame = ttk.Frame(self.parent)
        self.parent.add(std_frame, text="标液设备")
        self.create_standard_type_section(std_frame)
        self.create_instrument_setting_section(std_frame)

    def create_submit_tab(self):
        """建「其他参数」标签页(固定参数+提交签名)。"""
        submit_frame = ttk.Frame(self.parent)
        self.parent.add(submit_frame, text="其他参数")

        # 启用固定参数设置（默认隐藏，通过复选框控制）
        self.fp_enable_var = tk.BooleanVar(value=False)
        fp_toggle_frame = ttk.Frame(submit_frame)
        fp_toggle_frame.pack(fill='x', padx=5, pady=(2, 0))
        ttk.Checkbutton(fp_toggle_frame, text=" 启用固定列参数设置",
                        variable=self.fp_enable_var,
                        style=self.app.large_cb_style,
                        command=self.toggle_fixed_params).pack(side='left')
        self.create_fixed_params_section(submit_frame)
        # 加载时初始隐藏（set_settings 里根据数据有无再打开）
        self.fp_frame.pack_forget()

        # 启用条件实验过程 labProc（默认隐藏，通过复选框控制）
        self.lab_proc_enable_var = tk.BooleanVar(value=False)
        self.lab_proc_toggle_frame = ttk.Frame(submit_frame)
        self.lab_proc_toggle_frame.pack(fill='x', padx=5, pady=(2, 0))
        ttk.Checkbutton(self.lab_proc_toggle_frame, text=" 启用条件实验过程 labProc",
                        variable=self.lab_proc_enable_var,
                        style=self.app.large_cb_style,
                        command=self.toggle_lab_proc).pack(side='left')
        self.create_lab_proc_section(submit_frame)
        self.lab_proc_frame.pack_forget()

        # 启用稀释备注（默认隐藏，通过复选框控制）
        self.dilution_remark_enable_var = tk.BooleanVar(value=False)
        self.dil_remark_toggle_frame = ttk.Frame(submit_frame)
        self.dil_remark_toggle_frame.pack(fill='x', padx=5, pady=(2, 0))
        ttk.Checkbutton(self.dil_remark_toggle_frame, text=" 启用稀释备注（按稀释倍数+定容体积自动生成）",
                        variable=self.dilution_remark_enable_var,
                        style=self.app.large_cb_style,
                        command=self.toggle_dilution_remark).pack(side='left')
        self.create_dilution_remark_section(submit_frame)
        self.dil_remark_frame.pack_forget()

        # 提交签名：勾选后序列录入完数据改调 submitOcExperiment（提交签名/推进工作流），否则 saveOcExperiment（仅保存）。
        self.require_signature_var = tk.BooleanVar(value=False)
        self.sign_frame = ttk.Frame(submit_frame)
        self.sign_frame.pack(fill='x', padx=5, pady=(2, 5))
        ttk.Checkbutton(self.sign_frame, text=" 提交签名（序列录入后调用 submitOcExperiment）",
                        variable=self.require_signature_var,
                        style=self.app.large_cb_style,
                        command=self.app.mark_modified).pack(side='left')

    def create_standard_type_section(self, parent):
        """创建标液类型设置区域"""
        standard_frame = ttk.LabelFrame(parent, text="标液类型设置", padding=10)
        standard_frame.pack(fill='x', padx=5, pady=5)

        # 标液类型选择
        type_frame = ttk.Frame(standard_frame)
        type_frame.pack(fill='x', pady=5)

        self.standard_type_var = tk.StringVar(value="fresh")

        # 现配现用单选按钮
        ttkb.Radiobutton(
            type_frame,
            text="现配现用",
            variable=self.standard_type_var,
            value="fresh",
            command=self.on_standard_type_change, bootstyle="primary").pack(side='left', padx=(0, 20))

        # 固定标液单选按钮
        ttkb.Radiobutton(
            type_frame,
            text="固定标液",
            variable=self.standard_type_var,
            value="fixed",
            command=self.on_standard_type_change, bootstyle="primary").pack(side='left', padx=(0, 20))

        # 新增：无需标液单选按钮
        ttkb.Radiobutton(
            type_frame,
            text="无需标液",
            variable=self.standard_type_var,
            value="none",
            command=self.on_standard_type_change, bootstyle="primary").pack(side='left', padx=(0, 20))

        # 条件标液单选按钮（类似条件设备，按文件名/项目名/试样描述匹配不同配制序号）
        ttkb.Radiobutton(
            type_frame,
            text="条件标液",
            variable=self.standard_type_var,
            value="conditional",
            command=self.on_standard_type_change, bootstyle="primary").pack(side='left')

        # 配制序号输入区域（始终显示，但状态根据选择变化）
        self.preparation_frame = ttk.Frame(standard_frame)
        self.preparation_frame.pack(fill='x', pady=5)  # 始终显示

        ttk.Label(self.preparation_frame, text="配制序号:").pack(side='left', padx=(20, 5))
        self.preparation_number_var = tk.StringVar()
        self.preparation_entry = ttk.Entry(
            self.preparation_frame,
            textvariable=self.preparation_number_var,
            width=80  # 增加宽度
        )
        self.preparation_entry.pack(side='left')

        # 配制序号说明另起一行
        prep_hint = ttk.Frame(standard_frame)
        prep_hint.pack(fill='x', padx=(20, 5), pady=(0, 2))
        ttk.Label(prep_hint, text="（多个标液用逗号分隔，如 D-9210,D-9211）",
                  foreground="gray").pack(side='left')

        # 绑定配制序号变化事件
        self.preparation_number_var.trace('w', self.on_preparation_number_change)

        # 条件标液规则表格（仅 standard_type=conditional 时显示）
        self.standard_rules_frame = ttk.Frame(standard_frame)
        # 输入行
        sr_input = ttk.Frame(self.standard_rules_frame)
        sr_input.pack(fill='x', pady=2)
        ttk.Label(sr_input, text="文件名:").pack(side='left', padx=(0, 3))
        self.sr_filename_var = tk.StringVar()
        ttk.Entry(sr_input, textvariable=self.sr_filename_var, width=8).pack(side='left', padx=(0, 6))
        ttk.Label(sr_input, text="项目名:").pack(side='left', padx=(0, 3))
        self.sr_project_var = tk.StringVar()
        ttk.Entry(sr_input, textvariable=self.sr_project_var, width=12).pack(side='left', padx=(0, 6))
        ttk.Label(sr_input, text="试样描述:").pack(side='left', padx=(0, 3))
        self.sr_desc_var = tk.StringVar()
        ttk.Entry(sr_input, textvariable=self.sr_desc_var, width=12).pack(side='left', padx=(0, 6))
        ttk.Label(sr_input, text="配制序号:").pack(side='left', padx=(0, 3))
        self.sr_prep_var = tk.StringVar()
        ttk.Entry(sr_input, textvariable=self.sr_prep_var, width=16).pack(side='left', padx=(0, 6))
        ttkb.Button(sr_input, text="添加", command=self.add_standard_rule, bootstyle="secondary").pack(side='left')
        # 表格
        sr_list = ttk.Frame(self.standard_rules_frame)
        sr_list.pack(fill='both', expand=True, pady=3)
        scols = ("序号", "文件名关键词", "项目名关键词", "试样描述关键词", "配制序号")
        self.sr_tree = ttk.Treeview(sr_list, columns=scols, show="headings", height=3)
        for c in scols:
            self.sr_tree.heading(c, text=c, anchor='w')
        self.sr_tree.column("序号", width=8, anchor='w')
        self.sr_tree.column("文件名关键词", width=70, anchor='w')
        self.sr_tree.column("项目名关键词", width=80, anchor='w')
        self.sr_tree.column("试样描述关键词", width=80, anchor='w')
        self.sr_tree.column("配制序号", width=520, anchor='w')
        self.sr_tree.pack(side='left', fill='both', expand=True)
        sr_scroll = ttk.Scrollbar(sr_list, orient="vertical", command=self.sr_tree.yview)
        sr_scroll.pack(side='right', fill='y')
        self.sr_tree.configure(yscrollcommand=sr_scroll.set)
        _setup_cell_tooltip(self.sr_tree, ("配制序号",))
        self.sr_tree.bind("<Double-1>", self.on_standard_rule_double_click)
        # 操作按钮
        sr_btn = ttk.Frame(self.standard_rules_frame)
        sr_btn.pack(fill='x', pady=2)
        ttkb.Button(sr_btn, text="删除选中", command=self.delete_standard_rule, bootstyle="danger").pack(side='left', padx=2)
        ttkb.Button(sr_btn, text="上移", command=lambda: self.move_standard_rule(-1), bootstyle="secondary").pack(side='left', padx=2)
        ttkb.Button(sr_btn, text="下移", command=lambda: self.move_standard_rule(1), bootstyle="secondary").pack(side='left', padx=2)

        # 初始状态设置
        self.on_standard_type_change()

    def create_instrument_setting_section(self, parent):
        """创建仪器设置区域"""
        instrument_frame = ttk.LabelFrame(parent, text="仪器设置", padding=10)
        instrument_frame.pack(fill='x', padx=5, pady=5)

        # 仪器设置选择
        setting_frame = ttk.Frame(instrument_frame)
        setting_frame.pack(fill='x', pady=5)

        self.instrument_setting_var = tk.StringVar(value="default")

        # 使用立方默认设备单选按钮
        ttkb.Radiobutton(
            setting_frame,
            text="默认设备",
            variable=self.instrument_setting_var,
            value="default",
            command=self.on_instrument_setting_change, bootstyle="primary").pack(side='left', padx=(0, 20))

        # 使用指定设备单选按钮
        ttkb.Radiobutton(
            setting_frame,
            text="指定设备",
            variable=self.instrument_setting_var,
            value="specified",
            command=self.on_instrument_setting_change, bootstyle="primary").pack(side='left', padx=(0, 20))

        # 条件设备单选按钮
        ttkb.Radiobutton(
            setting_frame,
            text="条件设备",
            variable=self.instrument_setting_var,
            value="conditional",
            command=self.on_instrument_setting_change, bootstyle="primary").pack(side='left')

        # 设备编号输入区域（始终显示，但状态根据选择变化）
        self.device_frame = ttk.Frame(instrument_frame)
        self.device_frame.pack(fill='x', pady=5)  # 始终显示

        ttk.Label(self.device_frame, text="设备编号(中/英文; 分隔):").pack(side='left', padx=(20, 5))
        self.device_number_var = tk.StringVar()
        self.device_entry = ttk.Entry(
            self.device_frame,
            textvariable=self.device_number_var,
            width=80  # 增加宽度
        )
        self.device_entry.pack(side='left')

        # 绑定设备编号变化事件
        self.device_number_var.trace('w', self.on_device_number_change)

        # 条件设备规则表格（仅 instrument_setting=conditional 时显示）
        self.equipment_rules_frame = ttk.Frame(instrument_frame)
        self.equipment_rules_frame.pack(fill='both', expand=True, pady=5)
        # 输入行
        er_input = ttk.Frame(self.equipment_rules_frame)
        er_input.pack(fill='x', pady=2)
        ttk.Label(er_input, text="文件名:").pack(side='left', padx=(0, 3))
        self.er_filename_var = tk.StringVar()
        ttk.Entry(er_input, textvariable=self.er_filename_var, width=8).pack(side='left', padx=(0, 6))
        ttk.Label(er_input, text="项目名:").pack(side='left', padx=(0, 3))
        self.er_project_var = tk.StringVar()
        ttk.Entry(er_input, textvariable=self.er_project_var, width=12).pack(side='left', padx=(0, 6))
        ttk.Label(er_input, text="试样描述:").pack(side='left', padx=(0, 3))
        self.er_desc_var = tk.StringVar()
        ttk.Entry(er_input, textvariable=self.er_desc_var, width=12).pack(side='left', padx=(0, 6))
        ttk.Label(er_input, text="设备编号:").pack(side='left', padx=(0, 3))
        self.er_device_var = tk.StringVar()
        ttk.Entry(er_input, textvariable=self.er_device_var, width=16).pack(side='left', padx=(0, 6))
        ttkb.Button(er_input, text="添加", command=self.add_equipment_rule, bootstyle="secondary").pack(side='left')
        # 表格
        er_list = ttk.Frame(self.equipment_rules_frame)
        er_list.pack(fill='both', expand=True, pady=3)
        cols = ("序号", "文件名关键词", "项目名关键词", "试样描述关键词", "设备编号")
        self.er_tree = ttk.Treeview(er_list, columns=cols, show="headings", height=3)
        for c in cols:
            self.er_tree.heading(c, text=c, anchor='w')
        self.er_tree.column("序号", width=8, anchor='w')
        self.er_tree.column("文件名关键词", width=70, anchor='w')
        self.er_tree.column("项目名关键词", width=80, anchor='w')
        self.er_tree.column("试样描述关键词", width=80, anchor='w')
        self.er_tree.column("设备编号", width=520, anchor='w')
        self.er_tree.pack(side='left', fill='both', expand=True)
        er_scroll = ttk.Scrollbar(er_list, orient="vertical", command=self.er_tree.yview)
        er_scroll.pack(side='right', fill='y')
        self.er_tree.configure(yscrollcommand=er_scroll.set)
        _setup_cell_tooltip(self.er_tree, ("设备编号",))
        self.er_tree.bind("<Double-1>", self.on_equipment_rule_double_click)
        # 操作按钮
        er_btn = ttk.Frame(self.equipment_rules_frame)
        er_btn.pack(fill='x', pady=2)
        ttkb.Button(er_btn, text="删除选中", command=self.delete_equipment_rule, bootstyle="danger").pack(side='left', padx=2)
        ttkb.Button(er_btn, text="上移", command=lambda: self.move_equipment_rule(-1), bootstyle="secondary").pack(side='left', padx=2)
        ttkb.Button(er_btn, text="下移", command=lambda: self.move_equipment_rule(1), bootstyle="secondary").pack(side='left', padx=2)

        # 初始状态设置
        self.on_instrument_setting_change()

    def create_fixed_params_section(self, parent):
        """创建固定参数设置区域：一个触发条件显示为一行，可含多个「参数=值」"""
        self.fp_frame = ttk.LabelFrame(parent, text="固定列参数设置", padding=5)
        self.fp_frame.pack(fill='x', padx=5, pady=2)
        frame = self.fp_frame

        # 输入区：触发条件(字段=值) + 参数名 + 值
        input_frame = ttk.Frame(frame)
        input_frame.pack(fill='x', pady=2)

        ttk.Label(input_frame, text="触发条件:").pack(side='left', padx=(0, 5))
        self.fp_trigger_field_var = tk.StringVar(value="检测项目")
        ttk.Combobox(input_frame, textvariable=self.fp_trigger_field_var,
                     values=["检测项目", "检测方法", "标准值", "默认"], state="readonly", width=10).pack(side='left', padx=(0, 2))
        ttk.Label(input_frame, text="=").pack(side='left', padx=(0, 2))
        self.fp_trigger_value_entry = ttk.Entry(input_frame, width=14)
        self.fp_trigger_value_entry.pack(side='left', padx=(0, 12))
        _bind_entry_tooltip(self.fp_trigger_value_entry)  # 多值触发较长，悬停显示完整内容

        ttk.Label(input_frame, text="参数名:").pack(side='left', padx=(0, 5))
        self.fp_param_name_entry = ttk.Entry(input_frame, width=18)
        self.fp_param_name_entry.pack(side='left', padx=(0, 8))

        ttk.Label(input_frame, text="值:").pack(side='left', padx=(0, 5))
        self.fp_param_value_entry = ttk.Entry(input_frame, width=12)
        self.fp_param_value_entry.pack(side='left', padx=(0, 10))

        ttkb.Button(input_frame, text="添加", command=self.add_fixed_param, bootstyle="secondary").pack(side='left')

        # 列表区：一个触发条件一行，多个参数合并显示
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill='both', expand=True, pady=2)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        columns = ("触发条件", "参数设置")
        self.fp_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=4)
        self.fp_tree.heading("触发条件", text="触发条件", anchor="w")
        self.fp_tree.column("触发条件", width=160, anchor="w")
        self.fp_tree.heading("参数设置", text="参数设置", anchor="w")
        self.fp_tree.column("参数设置", width=520, anchor="w")

        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.fp_tree.yview)
        hsb = ttk.Scrollbar(list_frame, orient="horizontal", command=self.fp_tree.xview)
        self.fp_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.fp_tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        self.fp_tree.bind("<Double-1>", self.on_fixed_param_double_click)
        # 触发条件/参数设置 列悬停显示全部内容(多值触发如「检测项目=可溶性铅（Pb），...」较长，窄列被截断)
        _setup_cell_tooltip(self.fp_tree, ("触发条件", "参数设置"))

        # 说明
        ttk.Label(frame, foreground="gray", wraplength=700, justify='left',
                  text="触发：检测项目=/检测方法=/标准值= 单条件；多条件用 ';' 连接(均需命中)，"
                       "如「检测项目=可溶性六价铬（CrVI）;标准值=≤0.005」。值支持逗号多值(任一命中)。").pack(fill='x', pady=(2, 0))

        # 操作按钮
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', pady=3)
        ttkb.Button(btn_frame, text="删除选中", command=self.delete_fixed_param, bootstyle="danger").pack(side='left', padx=2)

    def add_fixed_param(self):
        """添加一个参数到触发条件对应的规则（同一触发条件只占一行）"""
        trigger_field = self.fp_trigger_field_var.get().strip()
        trigger_value = self.fp_trigger_value_entry.get().strip()
        param_name = self.fp_param_name_entry.get().strip()
        param_value = self.fp_param_value_entry.get().strip()

        if not param_name:
            messagebox.showwarning("输入错误", "请填写参数名")
            return

        if trigger_field == "默认":
            trigger = "默认"
        else:
            if not trigger_value:
                messagebox.showwarning("输入错误", "请填写触发条件值")
                return
            trigger = f"{trigger_field}={trigger_value}"
        # 同一触发条件并入同一行
        rule = next((r for r in self.fixed_params if r["trigger"] == trigger), None)
        if rule is None:
            rule = {"trigger": trigger, "params": []}
            self.fixed_params.append(rule)
        rule["params"].append({"name": param_name, "value": param_value})

        self.refresh_fixed_params_tree()

        # 清空参数输入（保留触发条件，便于连续为同一条件添加多个参数）
        self.fp_param_name_entry.delete(0, tk.END)
        self.fp_param_value_entry.delete(0, tk.END)

        self.app.mark_modified()

    def delete_fixed_param(self):
        """删除选中的触发条件规则（整行）"""
        selected_items = self.fp_tree.selection()
        if not selected_items:
            messagebox.showwarning("选择错误", "请先选择要删除的行")
            return

        if not messagebox.askyesno("确认删除", "确定要删除选中的触发条件及其全部参数吗？"):
            return

        for item in reversed(selected_items):
            index = self.fp_tree.index(item)
            if 0 <= index < len(self.fixed_params):
                del self.fixed_params[index]

        self.refresh_fixed_params_tree()
        self.app.mark_modified()

    def refresh_fixed_params_tree(self):
        """刷新固定参数列表（一个触发条件一行）"""
        for item in self.fp_tree.get_children():
            self.fp_tree.delete(item)

        for rule in self.fixed_params:
            params_str = "; ".join(f'{p["name"]}={p["value"]}' for p in rule.get("params", []))
            self.fp_tree.insert("", "end", values=(rule.get("trigger", ""), params_str))

    def on_fixed_param_double_click(self, event):
        """双击编辑：触发条件列直接改；参数设置列按「名=值; 名=值」解析"""
        item = self.fp_tree.selection()
        if not item:
            return

        item = item[0]
        column = self.fp_tree.identify_column(event.x)
        column_index = int(column.replace('#', '')) - 1
        column_name = self.fp_tree.heading(column_index)['text']

        current_values = self.fp_tree.item(item, 'values')
        current_value = current_values[column_index]

        x, y, width, height = self.fp_tree.bbox(item, column)
        entry = ttk.Entry(self.fp_tree)
        entry.place(x=x, y=y - 4, width=width, height=height + 8)
        entry.insert(0, current_value)
        entry.focus_set()

        def save_edit(event=None):
            if not entry.winfo_exists():
                return
            new_value = entry.get()

            new_values = list(current_values)
            new_values[column_index] = new_value
            self.fp_tree.item(item, values=new_values)

            index = self.fp_tree.index(item)
            if 0 <= index < len(self.fixed_params):
                rule = self.fixed_params[index]
                if column_name == "触发条件":
                    rule["trigger"] = new_value
                elif column_name == "参数设置":
                    # 解析「名=值; 名=值」回 params 列表（分号中英文皆可）
                    params = []
                    for part in new_value.replace("；", ";").split(";"):
                        part = part.strip()
                        if not part:
                            continue
                        name, _, value = part.partition("=")
                        params.append({"name": name.strip(), "value": value.strip()})
                    rule["params"] = params

            self.app.mark_modified()
            entry.destroy()

        def cancel_edit(event=None):
            if entry.winfo_exists():
                entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", cancel_edit)

    @staticmethod
    def _normalize_fixed_params(raw):
        """将加载的固定参数统一为 [{trigger, params:[{name,value}]}]；兼容旧扁平格式"""
        result = []
        for item in raw:
            if not isinstance(item, dict):
                continue
            if "params" in item:
                result.append({"trigger": item.get("trigger", ""),
                               "params": list(item.get("params", []))})
            elif "param_name" in item:
                # 旧扁平格式：按 trigger 合并成一行
                trig = item.get("trigger", "")
                rule = next((r for r in result if r["trigger"] == trig), None)
                if rule is None:
                    rule = {"trigger": trig, "params": []}
                    result.append(rule)
                rule["params"].append({"name": item.get("param_name", ""),
                                       "value": item.get("param_value", "")})
        return result

    def on_standard_type_change(self):
        """标液类型改变时的处理"""
        standard_type = self.standard_type_var.get()

        if standard_type == "fixed":
            # 固定标液：启用配制序号输入框
            self.preparation_entry.config(state="normal")
        else:
            # 现配现用/无需标液/条件标液：禁用配制序号输入框
            self.preparation_entry.config(state="disabled")

        # 条件标液：显示规则表格；其余隐藏
        self.standard_rules_frame.pack_forget()
        if standard_type == "conditional":
            self.standard_rules_frame.pack(fill='both', expand=True, pady=5)

        # 标记已修改
        self.app.mark_modified()

    def on_instrument_setting_change(self):
        """仪器设置改变时的处理"""
        mode = self.instrument_setting_var.get()
        self.equipment_rules_frame.pack_forget()
        if mode == "specified":
            self.device_frame.pack(fill='x', pady=5)
            self.device_entry.config(state="normal")
        elif mode == "conditional":
            self.device_frame.pack_forget()
            self.equipment_rules_frame.pack(fill='both', expand=True, pady=5)
        else:  # default
            self.device_frame.pack(fill='x', pady=5)
            self.device_entry.config(state="disabled")

        # 标记已修改
        self.app.mark_modified()

    def on_preparation_number_change(self, *args):
        """配制序号改变时的处理"""
        # 标记已修改
        self.app.mark_modified()

    def toggle_fixed_params(self):
        """复选框控制固定参数设置的显示/隐藏"""
        if self.fp_enable_var.get():
            self.fp_frame.pack(fill='x', padx=5, pady=2, before=self.lab_proc_toggle_frame)
        else:
            self.fp_frame.pack_forget()

    def toggle_lab_proc(self):
        """复选框控制条件实验过程设置的显示/隐藏"""
        if self.lab_proc_enable_var.get():
            self.lab_proc_frame.pack(fill='x', padx=5, pady=2, before=self.sign_frame)
        else:
            self.lab_proc_frame.pack_forget()

    def toggle_dilution_remark(self):
        """复选框控制稀释备注设置的显示/隐藏"""
        if self.dilution_remark_enable_var.get():
            self.dil_remark_frame.pack(fill='x', padx=5, pady=2, before=self.sign_frame)
        else:
            self.dil_remark_frame.pack_forget()
        self.app.mark_modified()

    def on_device_number_change(self, *args):
        """设备编号改变时的处理"""
        # 标记已修改
        self.app.mark_modified()

    def get_settings(self):
        """获取其他参数设置"""
        return {
            "standard_type": self.standard_type_var.get(),
            "preparation_number": self.preparation_number_var.get().strip(),
            "standard_rules": self.standard_rules,
            "instrument_setting": self.instrument_setting_var.get(),
            "device_number": self.device_number_var.get().strip(),
            "equipment_rules": self.equipment_rules,
            "fixed_params": self.fixed_params,
            "lab_proc_enabled": self.lab_proc_enable_var.get(),
            "lab_proc_rules": self.lab_proc_rules,
            "require_signature": self.require_signature_var.get(),
            "dilution_remark_enabled": self.dilution_remark_enable_var.get(),
            "dilution_volume_column": self.dilution_volume_col_var.get().strip()
        }

    def set_settings(self, settings):
        """设置其他参数配置"""
        if not settings:
            return

        # 设置标液类型
        if "standard_type" in settings:
            self.standard_type_var.set(settings["standard_type"])

        # 设置配制序号
        if "preparation_number" in settings:
            self.preparation_number_var.set(settings["preparation_number"])

        # 设置提交签名
        if "require_signature" in settings:
            self.require_signature_var.set(bool(settings["require_signature"]))

        # 设置仪器设置
        if "instrument_setting" in settings:
            self.instrument_setting_var.set(settings["instrument_setting"])

        # 设置设备编号
        if "device_number" in settings:
            self.device_number_var.set(settings["device_number"])

        # 设置条件设备规则
        if "equipment_rules" in settings:
            self.equipment_rules = list(settings["equipment_rules"])
            self.refresh_equipment_rules_tree()

        # 设置条件标液规则
        if "standard_rules" in settings:
            self.standard_rules = list(settings["standard_rules"])
            self.refresh_standard_rules_tree()

        # 设置固定参数（兼容旧扁平格式：按 trigger 合并为一行）
        fp_data = settings.get("fixed_params", [])
        self.fixed_params = self._normalize_fixed_params(fp_data)
        self.refresh_fixed_params_tree()
        if fp_data:
            self.fp_enable_var.set(True)
            self.fp_frame.pack(fill='x', padx=5, pady=2, before=self.lab_proc_toggle_frame)

        # 设置条件实验过程规则
        if "lab_proc_enabled" in settings:
            self.lab_proc_enable_var.set(bool(settings["lab_proc_enabled"]))
        if "lab_proc_rules" in settings:
            self.lab_proc_rules = list(settings["lab_proc_rules"])
            self.refresh_lab_proc_rules_tree()
        if self.lab_proc_enable_var.get() and self.lab_proc_rules:
            self.lab_proc_frame.pack(fill='x', padx=5, pady=2, before=self.sign_frame)

        # 设置稀释备注
        if "dilution_remark_enabled" in settings:
            self.dilution_remark_enable_var.set(bool(settings["dilution_remark_enabled"]))
        if "dilution_volume_column" in settings:
            self.dilution_volume_col_var.set(settings["dilution_volume_column"])
        if self.dilution_remark_enable_var.get():
            self.dil_remark_frame.pack(fill='x', padx=5, pady=2, before=self.sign_frame)

        # 更新显示
        self.on_standard_type_change()
        self.on_instrument_setting_change()

    # ---------- 条件设备规则 CRUD ----------

    def add_equipment_rule(self):
        """添加一条条件设备规则"""
        fn = self.er_filename_var.get().strip()
        proj = self.er_project_var.get().strip()
        desc = self.er_desc_var.get().strip()
        dev = self.er_device_var.get().strip()
        if not dev:
            messagebox.showwarning("输入错误", "请填写设备编号")
            return
        self.equipment_rules.append({
            "filename": fn,
            "project_name": proj,
            "desc": desc,
            "device_number": dev
        })
        self.refresh_equipment_rules_tree()
        # 清空输入
        self.er_filename_var.set("")
        self.er_project_var.set("")
        self.er_desc_var.set("")
        self.er_device_var.set("")
        self.app.mark_modified()

    def delete_equipment_rule(self):
        """删除选中的条件设备规则"""
        sel = self.er_tree.selection()
        if not sel:
            messagebox.showwarning("选择错误", "请先选择要删除的规则")
            return
        idx = int(sel[0]) - 1
        if 0 <= idx < len(self.equipment_rules):
            del self.equipment_rules[idx]
            self.refresh_equipment_rules_tree()
            self.app.mark_modified()

    def move_equipment_rule(self, delta):
        """上移(delta=-1)或下移(delta=1)选中的规则"""
        sel = self.er_tree.selection()
        if not sel:
            return
        i = int(sel[0]) - 1
        if not (0 <= i < len(self.equipment_rules)):
            return
        j = i + delta
        if not (0 <= j < len(self.equipment_rules)):
            return
        self.equipment_rules[i], self.equipment_rules[j] = self.equipment_rules[j], self.equipment_rules[i]
        self.refresh_equipment_rules_tree()
        # 重新选中移动后的行
        for item in self.er_tree.get_children():
            if int(item) == j + 1:
                self.er_tree.selection_set(item)
                break
        self.app.mark_modified()

    def refresh_equipment_rules_tree(self):
        """重建条件设备规则表格"""
        for item in self.er_tree.get_children():
            self.er_tree.delete(item)
        for i, r in enumerate(self.equipment_rules):
            self.er_tree.insert("", "end", iid=str(i + 1),
                               values=(str(i + 1), r.get("filename", ""),
                                       r.get("project_name", ""), r.get("desc", ""),
                                       r.get("device_number", "")))

    def on_equipment_rule_double_click(self, event):
        """双击条件设备规则单元格就地编辑（序号列不可编辑；对齐 on_fixed_param_double_click 模式）"""
        sel = self.er_tree.selection()
        if not sel:
            return
        item = sel[0]
        column_index = int(self.er_tree.identify_column(event.x).replace('#', '')) - 1
        col_key = {1: "filename", 2: "project_name", 3: "desc", 4: "device_number"}.get(column_index)
        if col_key is None:  # 序号列或越界
            return

        current_values = self.er_tree.item(item, 'values')
        current_value = current_values[column_index]

        x, y, width, height = self.er_tree.bbox(item, self.er_tree.identify_column(event.x))
        entry = ttk.Entry(self.er_tree)
        entry.place(x=x, y=y - 4, width=width, height=height + 8)
        entry.insert(0, current_value)
        entry.focus_set()

        def save_edit(event=None):
            if not entry.winfo_exists():
                return
            new_value = entry.get()
            new_values = list(current_values)
            new_values[column_index] = new_value
            self.er_tree.item(item, values=new_values)
            idx = int(item) - 1
            if 0 <= idx < len(self.equipment_rules):
                self.equipment_rules[idx][col_key] = new_value
            self.app.mark_modified()
            entry.destroy()

        def cancel_edit(event=None):
            if entry.winfo_exists():
                entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", cancel_edit)

    # ---------- 条件标液规则 CRUD ----------

    def add_standard_rule(self):
        """添加一条条件标液规则"""
        fn = self.sr_filename_var.get().strip()
        proj = self.sr_project_var.get().strip()
        desc = self.sr_desc_var.get().strip()
        prep = self.sr_prep_var.get().strip()
        if not prep:
            messagebox.showwarning("输入错误", "请填写配制序号")
            return
        self.standard_rules.append({
            "filename": fn,
            "project_name": proj,
            "desc": desc,
            "preparation_number": prep
        })
        self.refresh_standard_rules_tree()
        # 清空输入
        self.sr_filename_var.set("")
        self.sr_project_var.set("")
        self.sr_desc_var.set("")
        self.sr_prep_var.set("")
        self.app.mark_modified()

    def delete_standard_rule(self):
        """删除选中的条件标液规则"""
        sel = self.sr_tree.selection()
        if not sel:
            messagebox.showwarning("选择错误", "请先选择要删除的规则")
            return
        idx = int(sel[0]) - 1
        if 0 <= idx < len(self.standard_rules):
            del self.standard_rules[idx]
            self.refresh_standard_rules_tree()
            self.app.mark_modified()

    def move_standard_rule(self, delta):
        """上移(delta=-1)或下移(delta=1)选中的规则"""
        sel = self.sr_tree.selection()
        if not sel:
            return
        i = int(sel[0]) - 1
        if not (0 <= i < len(self.standard_rules)):
            return
        j = i + delta
        if not (0 <= j < len(self.standard_rules)):
            return
        self.standard_rules[i], self.standard_rules[j] = self.standard_rules[j], self.standard_rules[i]
        self.refresh_standard_rules_tree()
        # 重新选中移动后的行
        for item in self.sr_tree.get_children():
            if int(item) == j + 1:
                self.sr_tree.selection_set(item)
                break
        self.app.mark_modified()

    def refresh_standard_rules_tree(self):
        """重建条件标液规则表格"""
        for item in self.sr_tree.get_children():
            self.sr_tree.delete(item)
        for i, r in enumerate(self.standard_rules):
            self.sr_tree.insert("", "end", iid=str(i + 1),
                               values=(str(i + 1), r.get("filename", ""),
                                       r.get("project_name", ""), r.get("desc", ""),
                                       r.get("preparation_number", "")))

    def on_standard_rule_double_click(self, event):
        """双击条件标液规则单元格就地编辑"""
        sel = self.sr_tree.selection()
        if not sel:
            return
        item = sel[0]
        column_index = int(self.sr_tree.identify_column(event.x).replace('#', '')) - 1
        col_key = {1: "filename", 2: "project_name", 3: "desc", 4: "preparation_number"}.get(column_index)
        if col_key is None:  # 序号列或越界
            return

        current_values = self.sr_tree.item(item, 'values')
        current_value = current_values[column_index]

        x, y, width, height = self.sr_tree.bbox(item, self.sr_tree.identify_column(event.x))
        entry = ttk.Entry(self.sr_tree)
        entry.place(x=x, y=y - 4, width=width, height=height + 8)
        entry.insert(0, current_value)
        entry.focus_set()

        def save_edit(event=None):
            if not entry.winfo_exists():
                return
            new_value = entry.get()
            new_values = list(current_values)
            new_values[column_index] = new_value
            self.sr_tree.item(item, values=new_values)
            idx = int(item) - 1
            if 0 <= idx < len(self.standard_rules):
                self.standard_rules[idx][col_key] = new_value
            self.app.mark_modified()
            entry.destroy()

        def cancel_edit(event=None):
            if entry.winfo_exists():
                entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", cancel_edit)

    # ---------- 条件实验过程(labProc)规则 ----------
    # 镜像条件设备规则：文件名/项目名/试样描述 关键词 AND 匹配 → 命中取首条，用其 lab_proc 覆盖该样品 experimentProcess。

    def create_dilution_remark_section(self, parent):
        """创建稀释备注设置区域(初始会被 create_submit_tab pack_forget)。"""
        self.dil_remark_frame = ttk.LabelFrame(parent, text="稀释备注设置", padding=5)
        self.dil_remark_frame.pack(fill='x', padx=5, pady=2)
        dr = ttk.Frame(self.dil_remark_frame)
        dr.pack(fill='x', pady=2)
        ttk.Label(dr, text="定容体积列名:").pack(side='left', padx=(0, 3))
        self.dilution_volume_col_var = tk.StringVar()
        ttk.Entry(dr, textvariable=self.dilution_volume_col_var, width=24).pack(side='left', padx=(0, 6))
        self.dilution_volume_col_var.trace('w', lambda *a: self.app.mark_modified())
        ttk.Label(dr, text="（如 体积V(mL)；基准稀释倍数走「固定参数 稀释因子F」）",
                  foreground="gray").pack(side='left')

    def create_lab_proc_section(self, parent):
        """创建条件实验过程规则区域(初始会被 create_submit_tab pack_forget)。"""
        self.lab_proc_frame = ttk.LabelFrame(parent, text="条件实验过程设置", padding=5)
        self.lab_proc_frame.pack(fill='x', padx=5, pady=2)
        lp_input = ttk.Frame(self.lab_proc_frame)
        lp_input.pack(fill='x', pady=2)
        ttk.Label(lp_input, text="文件名:").pack(side='left', padx=(0, 3))
        self.lp_filename_var = tk.StringVar()
        ttk.Entry(lp_input, textvariable=self.lp_filename_var, width=8).pack(side='left', padx=(0, 6))
        ttk.Label(lp_input, text="项目名:").pack(side='left', padx=(0, 3))
        self.lp_project_var = tk.StringVar()
        ttk.Entry(lp_input, textvariable=self.lp_project_var, width=12).pack(side='left', padx=(0, 6))
        ttk.Label(lp_input, text="试样描述:").pack(side='left', padx=(0, 3))
        self.lp_desc_var = tk.StringVar()
        ttk.Entry(lp_input, textvariable=self.lp_desc_var, width=12).pack(side='left', padx=(0, 6))
        ttk.Label(lp_input, text="试验过程:").pack(side='left', padx=(0, 3))
        self.lp_proc_var = tk.StringVar()
        ttk.Entry(lp_input, textvariable=self.lp_proc_var, width=16).pack(side='left', padx=(0, 6))
        ttkb.Button(lp_input, text="添加", command=self.add_lab_proc_rule, bootstyle="secondary").pack(side='left')
        lp_list = ttk.Frame(self.lab_proc_frame)
        lp_list.pack(fill='both', expand=True, pady=3)
        cols = ("序号", "文件名关键词", "项目名关键词", "试样描述关键词", "试验过程")
        self.lp_tree = ttk.Treeview(lp_list, columns=cols, show="headings", height=3)
        for c in cols:
            self.lp_tree.heading(c, text=c, anchor='w')
        self.lp_tree.column("序号", width=8, anchor='w')
        self.lp_tree.column("文件名关键词", width=70, anchor='w')
        self.lp_tree.column("项目名关键词", width=80, anchor='w')
        self.lp_tree.column("试样描述关键词", width=80, anchor='w')
        self.lp_tree.column("试验过程", width=300, anchor='w')
        self.lp_tree.pack(side='left', fill='both', expand=True)
        lp_scroll = ttk.Scrollbar(lp_list, orient="vertical", command=self.lp_tree.yview)
        lp_scroll.pack(side='right', fill='y')
        self.lp_tree.configure(yscrollcommand=lp_scroll.set)
        _setup_cell_tooltip(self.lp_tree, ("试验过程",), wraplength=700)
        self.lp_tree.bind("<Double-1>", self.on_lab_proc_rule_double_click)
        lp_btn = ttk.Frame(self.lab_proc_frame)
        lp_btn.pack(fill='x', pady=2)
        ttkb.Button(lp_btn, text="删除选中", command=self.delete_lab_proc_rule, bootstyle="danger").pack(side='left', padx=2)

    def add_lab_proc_rule(self):
        """添加一条条件实验过程规则"""
        proc = self.lp_proc_var.get().strip()
        if not proc:
            messagebox.showwarning("输入错误", "请填写试验过程")
            return
        self.lab_proc_rules.append({
            "filename": self.lp_filename_var.get().strip(),
            "project_name": self.lp_project_var.get().strip(),
            "desc": self.lp_desc_var.get().strip(),
            "lab_proc": proc
        })
        self.refresh_lab_proc_rules_tree()
        self.lp_filename_var.set(""); self.lp_project_var.set(""); self.lp_desc_var.set(""); self.lp_proc_var.set("")
        self.app.mark_modified()

    def delete_lab_proc_rule(self):
        """删除选中的条件实验过程规则"""
        sel = self.lp_tree.selection()
        if not sel:
            messagebox.showwarning("选择错误", "请先选择要删除的规则")
            return
        idx = int(sel[0]) - 1
        if 0 <= idx < len(self.lab_proc_rules):
            del self.lab_proc_rules[idx]
            self.refresh_lab_proc_rules_tree()
            self.app.mark_modified()

    def refresh_lab_proc_rules_tree(self):
        """重建条件实验过程规则表格"""
        for item in self.lp_tree.get_children():
            self.lp_tree.delete(item)
        for i, r in enumerate(self.lab_proc_rules):
            self.lp_tree.insert("", "end", iid=str(i + 1),
                               values=(str(i + 1), r.get("filename", ""), r.get("project_name", ""),
                                       r.get("desc", ""), r.get("lab_proc", "")))

    def on_lab_proc_rule_double_click(self, event):
        """双击条件实验过程规则单元格就地编辑（序号列不可编辑）"""
        sel = self.lp_tree.selection()
        if not sel:
            return
        item = sel[0]
        column_index = int(self.lp_tree.identify_column(event.x).replace('#', '')) - 1
        col_key = {1: "filename", 2: "project_name", 3: "desc", 4: "lab_proc"}.get(column_index)
        if col_key is None:
            return
        current_values = self.lp_tree.item(item, 'values')
        current_value = current_values[column_index]
        x, y, width, height = self.lp_tree.bbox(item, self.lp_tree.identify_column(event.x))
        entry = ttk.Entry(self.lp_tree)
        entry.place(x=x, y=y - 4, width=width, height=height + 8)
        entry.insert(0, current_value)
        entry.focus_set()

        def save_edit(event=None):
            if not entry.winfo_exists():
                return
            new_value = entry.get()
            new_values = list(current_values)
            new_values[column_index] = new_value
            self.lp_tree.item(item, values=new_values)
            idx = int(item) - 1
            if 0 <= idx < len(self.lab_proc_rules):
                self.lab_proc_rules[idx][col_key] = new_value
            self.app.mark_modified()
            entry.destroy()

        def cancel_edit(event=None):
            if entry.winfo_exists():
                entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<FocusOut>", save_edit)
        entry.bind("<Escape>", cancel_edit)


if __name__ == "__main__":
    root = ttkb.Window(themename="sandstone-light")
    app = QueryAppFixed(root)
    root.mainloop()