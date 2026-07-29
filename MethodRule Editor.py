import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk
from tkinter import messagebox, filedialog
import ttkbootstrap as ttkb  # 档1: 现代主题(sandstone-light)，ttk 控件自动套用
import yaml
import os
import paths
import random


class QueryTab:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.query_rules = []
        self.exclusion_rules = []  # 新增：排除项目规则
        self.sum_rules = []  # 新增：总和规则
        self.create_tab()

    def _setup_cell_tooltip(self, tree, columns):
        """为 tree 指定列悬停显示单元格全部内容(项目/检测方法等长文本列)。"""
        tip = tk.Toplevel(tree)
        tip.withdraw()
        tip.overrideredirect(True)
        label = ttk.Label(tip, background="#ffffe0", relief="solid", borderwidth=1,
                          padding=(6, 3), wraplength=420)
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

    def create_tab(self):
        # 创建查询条件标签页
        query_frame = ttk.Frame(self.parent)
        self.parent.add(query_frame, text="查询条件")

        # 查询条件规则管理区域
        self.create_query_rules_management(query_frame)

    def create_query_rules_management(self, parent):
        """创建查询条件规则管理区域"""
        # 创建规则管理框架
        rules_frame = ttk.LabelFrame(parent, text="查询条件规则管理", padding=5)
        rules_frame.pack(fill='both', expand=True, padx=5, pady=5)

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

        # 第二行：Retest / 最大选中数量 / Mode / 添加
        row1 = ttk.Frame(input_frame)
        row1.pack(fill='x', pady=1)
        row1.columnconfigure(6, weight=1)  # 末尾空白撑开，把添加按钮推到右侧

        ttk.Label(row1, text="Retest:").grid(row=0, column=0, padx=(0, 5), pady=1, sticky='w')
        self.query_cancel_test_var = tk.BooleanVar()
        self.query_cancel_test_check = ttkb.Checkbutton(
            row1, variable=self.query_cancel_test_var, bootstyle="primary")
        self.query_cancel_test_check.grid(row=0, column=1, padx=(0, 10), pady=1, sticky='w')

        ttk.Label(row1, text="最大选中数量:").grid(row=0, column=2, padx=(0, 5), pady=1, sticky='w')
        self.query_max_select_entry = ttk.Entry(row1, width=8)
        self.query_max_select_entry.grid(row=0, column=3, padx=(0, 10), pady=1, sticky='w')

        ttk.Label(row1, text="Mode:").grid(row=0, column=4, padx=(0, 5), pady=1, sticky='w')
        self.query_mode_var = tk.StringVar(value="方法")
        self.query_mode_combo = ttk.Combobox(row1, textvariable=self.query_mode_var,
                                             values=["方法", "样品"], state="readonly", width=8)
        self.query_mode_combo.grid(row=0, column=5, padx=(0, 5), pady=1, sticky='w')

        self.add_query_rule_btn = ttkb.Button(row1, text="添加", command=self.add_query_rule, bootstyle="secondary")
        self.add_query_rule_btn.grid(row=0, column=7, padx=(5, 0), pady=1, sticky='e')

        # 创建规则显示区域
        display_frame = ttk.Frame(rules_frame)
        display_frame.pack(fill='both', expand=True, pady=5)

        # 创建规则表格 - 列顺序与输入框一致：项目/检测方法/Retest/最大选中数量；Mode 为非输入项置末
        columns = ("No", "项目", "检测方法", "Retest", "Max", "Mode")
        self.query_rules_tree = ttk.Treeview(display_frame, columns=columns, show="headings", height=8)

        # 设置列标题和宽度
        # 短列(No/Retest/Max/Mode)固定窄列宽不随窗口拉伸；文本列(项目/检测方法)随窗口伸缩
        column_configs = {
            "No": {"width": 50, "anchor": "center", "stretch": False},  # 序号列居中对齐
            "项目": {"width": 200, "anchor": "center"},  # 项目列居中对齐
            "检测方法": {"width": 200, "anchor": "center"},  # 检测方法列居中对齐
            "Retest": {"width": 90, "anchor": "center", "stretch": False},  # 注销复测列居中对齐
            "Max": {"width": 80, "anchor": "center", "stretch": False},  # 最大选中数量列居中对齐
            "Mode": {"width": 90, "anchor": "center", "stretch": False}  # 录入方式列居中对齐
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
        self._setup_cell_tooltip(self.query_rules_tree, ("项目", "检测方法"))

        # 创建规则操作按钮 - 新增上移、下移按钮，以及排除和总和按钮
        button_frame = ttk.Frame(rules_frame)
        button_frame.pack(fill='x', pady=3)

        # 左侧按钮
        ttkb.Button(button_frame, text="上移", command=self.move_query_rule_up, bootstyle="secondary").pack(side='left', padx=2)
        ttkb.Button(button_frame, text="下移", command=self.move_query_rule_down, bootstyle="secondary").pack(side='left', padx=2)
        ttkb.Button(button_frame, text="删除", command=self.delete_query_rule, bootstyle="danger").pack(side='left', padx=2)

        # 右侧按钮 - 新增排除和总和按钮
        ttkb.Button(button_frame, text="排除", command=self.show_exclusion_dialog, bootstyle="secondary").pack(side='right', padx=2)
        ttkb.Button(button_frame, text="总和", command=self.show_sum_dialog, bootstyle="secondary").pack(side='right', padx=2)

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

    def show_sum_dialog(self):
        """显示总和设置对话框"""
        dialog = ttkb.Toplevel(self.parent)
        dialog.title("总和设置")
        dialog.geometry("1020x760")  # 增加宽度以容纳更多内容
        dialog.transient(self.parent)
        dialog.grab_set()

        # 允许对话框调整大小
        dialog.resizable(True, True)

        # 计算居中位置
        self.center_dialog(dialog, 1020, 760)

        # 主容器
        main_frame = ttk.Frame(dialog, padding=10)
        main_frame.pack(fill='both', expand=True)
        main_frame.columnconfigure(0, weight=1)

        # 创建输入区域 - 所有控件在一行
        input_frame = ttk.LabelFrame(main_frame, text="添加总和条款", padding=5)
        input_frame.pack(fill='x', pady=(0, 5))

        # 配置列权重，使所有输入框都能扩展
        input_frame.columnconfigure(0, weight=0)  # 项目标签
        input_frame.columnconfigure(1, weight=3)  # 项目输入框 - 增加权重
        input_frame.columnconfigure(2, weight=0)  # 方法标签
        input_frame.columnconfigure(3, weight=3)  # 方法输入框 - 增加权重
        input_frame.columnconfigure(4, weight=0)  # ID1标签
        input_frame.columnconfigure(5, weight=1)  # ID1输入框 - 减少权重
        input_frame.columnconfigure(6, weight=0)  # ID2标签
        input_frame.columnconfigure(7, weight=1)  # ID2输入框 - 减少权重
        input_frame.columnconfigure(8, weight=0)  # 添加按钮

        # 项目输入 - 增加初始宽度
        ttk.Label(input_frame, text="项目:").grid(row=0, column=0, padx=(0, 0), pady=3, sticky='w')
        sum_project_var = tk.StringVar()
        sum_project_entry = ttk.Entry(input_frame, textvariable=sum_project_var, width=15)
        sum_project_entry.grid(row=0, column=1, padx=(0, 10), pady=3, sticky='ew')

        # 检测方法输入 - 增加初始宽度
        ttk.Label(input_frame, text="检测方法:").grid(row=0, column=2, padx=(0, 0), pady=3, sticky='w')
        sum_method_var = tk.StringVar()
        sum_method_entry = ttk.Entry(input_frame, textvariable=sum_method_var, width=15)
        sum_method_entry.grid(row=0, column=3, padx=(0, 10), pady=3, sticky='ew')

        # ID1输入 - 减少初始宽度
        ttk.Label(input_frame, text="ID1:").grid(row=0, column=4, padx=(0, 0), pady=3, sticky='w')
        sum_id1_var = tk.StringVar()
        sum_id1_entry = ttk.Entry(input_frame, textvariable=sum_id1_var, width=8)
        sum_id1_entry.grid(row=0, column=5, padx=(0, 10), pady=3, sticky='ew')

        # ID2输入 - 减少初始宽度
        ttk.Label(input_frame, text="ID2:").grid(row=0, column=6, padx=(0, 0), pady=3, sticky='w')
        sum_id2_var = tk.StringVar()
        sum_id2_entry = ttk.Entry(input_frame, textvariable=sum_id2_var, width=8)
        sum_id2_entry.grid(row=0, column=7, padx=(0, 10), pady=3, sticky='ew')

        # 添加按钮
        def add_sum_rule():
            project = sum_project_var.get().strip()
            method = sum_method_var.get().strip()
            id1 = sum_id1_var.get().strip()
            id2 = sum_id2_var.get().strip()

            # 修改验证逻辑：只要有一个字段不为空即可，不要求所有字段都必须填写
            if not project and not method and not id1 and not id2:
                messagebox.showwarning("输入错误", "请至少填写一个字段")
                return

            # 检查是否已存在相同的规则
            new_rule = {
                "project": project,
                "method": method,
                "id1": id1,
                "id2": id2
            }

            # 检查重复
            for existing_rule in self.sum_rules:
                if (existing_rule.get("project", "") == project and
                        existing_rule.get("method", "") == method and
                        existing_rule.get("id1", "") == id1 and
                        existing_rule.get("id2", "") == id2):
                    messagebox.showwarning("重复条目", "已存在相同的总和条款，无法重复添加")
                    return

            self.sum_rules.append(new_rule)
            refresh_sum_list()

            # 清空输入框
            sum_project_var.set("")
            sum_method_var.set("")
            sum_id1_var.set("")
            sum_id2_var.set("")

            # 标记已修改
            self.app.mark_modified()

        ttkb.Button(input_frame, text="添加", command=add_sum_rule, bootstyle="secondary").grid(row=0, column=8, padx=5, pady=3, sticky='e')

        ttkb.Button(input_frame, text="添加", command=add_sum_rule, bootstyle="secondary").grid(row=0, column=8, padx=5, pady=3, sticky='e')

        # 总和条款列表
        list_frame = ttk.LabelFrame(main_frame, text="已添加的总和条款", padding=5)
        list_frame.pack(fill='both', expand=True, pady=(0, 5))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        # 创建列表
        columns = ("项目", "检测方法", "ID1", "ID2")
        sum_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)

        # 设置列宽 - 增加项目和方法列宽，减少ID1和ID2列宽
        sum_tree.heading("项目", text="项目", anchor="w")
        sum_tree.column("项目", width=190, anchor="w")
        sum_tree.heading("检测方法", text="检测方法", anchor="w")
        sum_tree.column("检测方法", width=190, anchor="w")
        sum_tree.heading("ID1", text="ID1", anchor="w")
        sum_tree.column("ID1", width=40, anchor="w")  # 减少ID1列宽
        sum_tree.heading("ID2", text="ID2", anchor="w")
        sum_tree.column("ID2", width=40, anchor="w")  # 减少ID2列宽

        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=sum_tree.yview)
        sum_tree.configure(yscrollcommand=scrollbar.set)

        sum_tree.grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')

        # 绑定双击事件，用于编辑总和规则
        sum_tree.bind("<Double-1>", lambda event: self.on_sum_rule_double_click(event, sum_tree))

        def refresh_sum_list():
            # 清空现有列表
            for item in sum_tree.get_children():
                sum_tree.delete(item)

            # 添加总和规则到列表
            for rule in self.sum_rules:
                sum_tree.insert("", "end", values=(
                    rule.get("project", ""),
                    rule.get("method", ""),
                    rule.get("id1", ""),
                    rule.get("id2", "")
                ))

        # 初始刷新列表
        refresh_sum_list()

        # 底部按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=5)

        # 删除选中按钮
        def delete_selected_sum():
            selected_items = sum_tree.selection()
            if not selected_items:
                messagebox.showwarning("选择错误", "请先选择要删除的总和条款")
                return

            if messagebox.askyesno("确认删除", "确定要删除选中的总和条款吗？"):
                # 从后往前删除
                for item in reversed(selected_items):
                    index = sum_tree.index(item)
                    if 0 <= index < len(self.sum_rules):
                        del self.sum_rules[index]

                refresh_sum_list()
                self.app.mark_modified()

        ttkb.Button(button_frame, text="删除选中", command=delete_selected_sum, bootstyle="danger").pack(side='left', padx=5)

        # 关闭按钮
        ttkb.Button(button_frame, text="关闭", command=dialog.destroy, bootstyle="secondary").pack(side='right', padx=5)

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
    def on_sum_rule_double_click(self, event, tree):
        """双击总和规则进行编辑"""
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
            if 0 <= index < len(self.sum_rules):
                rule = self.sum_rules[index]
                if column_index == 0:  # 项目列
                    rule["project"] = new_value
                elif column_index == 1:  # 检测方法列
                    rule["method"] = new_value
                elif column_index == 2:  # ID1列
                    rule["id1"] = new_value
                elif column_index == 3:  # ID2列
                    rule["id2"] = new_value

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
            "max_select": max_select
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
                rule.get("input_method", "方法")  # 默认值为"方法"
            ))

    def get_rules(self):
        """获取查询规则"""
        return {
            "query_rules": self.query_rules,
            "exclusion_rules": self.exclusion_rules,
            "sum_rules": self.sum_rules
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
        self.query_rules = query_rules

        # 设置排除规则和总和规则
        self.exclusion_rules = rules_data.get("exclusion_rules", [])
        self.sum_rules = rules_data.get("sum_rules", [])

        self.refresh_query_rules_tree()


class MethodTab:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.default_rules = []
        self.filename_rules = []
        self.create_tab()
        self.show_all_var.set(True)
        self.on_show_all_change()  # 调用更新显示方法

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

        # 创建切换模式选择行
        mode_frame = ttk.Frame(rules_frame)
        mode_frame.pack(fill='x', pady=2)

        # 切换模式变量
        self.switch_mode = tk.StringVar(value="default")

        # 默认切换类型
        ttkb.Radiobutton(mode_frame, text="默认切换",
                        variable=self.switch_mode, value="default", bootstyle="primary").pack(side='left', padx=(0, 15))

        # 名称切换：文件名关键字 + 试样描述关键字(均非空时需同时命中)切换方法
        ttkb.Radiobutton(mode_frame, text="名称切换",
                        variable=self.switch_mode, value="filename", bootstyle="primary").pack(side='left', padx=(0, 15))

        # 全部复选框
        self.show_all_var = tk.BooleanVar()
        self.show_all_check = ttkb.Checkbutton(
            mode_frame, text="全部", variable=self.show_all_var, command=self.on_show_all_change, bootstyle="primary")
        self.show_all_check.pack(side='left')

        # 绑定切换模式变化事件
        self.switch_mode.trace('w', self.on_mode_change)

        # 创建规则输入区域
        input_frame = ttk.Frame(rules_frame)
        input_frame.pack(fill='x', pady=2)

        # 配置列权重
        input_frame.columnconfigure(0, weight=0)  # 描述标签
        input_frame.columnconfigure(1, weight=0)  # 描述输入框
        input_frame.columnconfigure(2, weight=0)  # 文件名标签
        input_frame.columnconfigure(3, weight=0)  # 文件名输入框
        input_frame.columnconfigure(4, weight=0)  # 项目名标签
        input_frame.columnconfigure(5, weight=0)  # 项目名输入框
        input_frame.columnconfigure(6, weight=0)  # 原ID标签
        input_frame.columnconfigure(7, weight=0)  # 原ID输入框
        input_frame.columnconfigure(8, weight=0)  # 目标ID标签
        input_frame.columnconfigure(9, weight=1)  # 目标ID输入框增大宽度
        input_frame.columnconfigure(10, weight=1)  # 空白区域扩展
        input_frame.columnconfigure(11, weight=0)  # 添加按钮

        # 文件名输入
        self.filename_label = ttk.Label(input_frame, text="文件名:")
        self.filename_entry = ttk.Entry(input_frame, width=12)

        # 项目名输入(名称切换模式下复用为"试样描述"关键字输入)
        self.project_name_label = ttk.Label(input_frame, text="项目名:")
        self.project_name_entry = ttk.Entry(input_frame, width=12)

        # 原方法ID
        self.from_id_label = ttk.Label(input_frame, text="原ID:")
        self.from_id_entry = ttk.Entry(input_frame, width=10)

        # 目标方法ID
        self.to_id_label = ttk.Label(input_frame, text="目标ID:")
        self.to_id_entry = ttk.Entry(input_frame, width=15)

        # 添加规则按钮
        self.add_rule_btn = ttkb.Button(input_frame, text="添加", command=self.add_rule, bootstyle="secondary")
        self.add_rule_btn.grid(row=0, column=11, padx=(5, 0), pady=1, sticky='e')

        # 创建规则显示区域
        display_frame = ttk.Frame(rules_frame)
        display_frame.pack(fill='both', expand=True, pady=5)

        # 创建规则表格
        columns = ("模式", "文件名", "试样描述", "项目名", "原ID", "目标ID")
        self.rules_tree = ttk.Treeview(display_frame, columns=columns, show="headings", height=6)

        # 设置列标题和宽度（stretch=True 列宽随窗口缩放；表头与内容 anchor 一致避免错位）
        column_configs = {
            "模式":   {"width": 60,  "anchor": "center"},
            "文件名": {"width": 120, "anchor": "w"},
            "试样描述": {"width": 140, "anchor": "w"},
            "项目名": {"width": 120, "anchor": "w"},
            "原ID":   {"width": 70,  "anchor": "center"},
            "目标ID": {"width": 90,  "anchor": "center"},
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

        # 模式列不允许编辑
        if column_index == 0:
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

        def save_edit(event=None):
            # 获取新值
            new_value = entry.get()

            # 更新显示
            new_values = list(current_values)
            new_values[column_index] = new_value
            self.rules_tree.item(item, values=new_values)

            # 更新数据
            index = self.rules_tree.index(item)
            column_name = self.rules_tree.heading(column_index)['text']

            # 确定规则属于哪种模式
            mode = current_values[0]  # 模式列
            if mode == "默认":
                rules_list = self.default_rules
            else:  # 名称
                rules_list = self.filename_rules

            # 在规则列表中查找匹配的规则
            for i, rule in enumerate(rules_list):
                if mode == "默认":
                    if (rule.get("project_name", "") == current_values[3] and
                            rule["from_id"] == current_values[4] and
                            rule["to_id"] == current_values[5]):
                        if column_name == "项目名":
                            rule["project_name"] = new_value
                        elif column_name == "原ID":
                            rule["from_id"] = new_value
                        elif column_name == "目标ID":
                            rule["to_id"] = new_value
                        break
                else:  # 名称
                    # 文件名/试样描述/项目名/原ID/目标ID
                    if ([rule.get("filename", ""), rule.get("desc", ""), rule.get("project_name", ""),
                            rule.get("from_id", ""), rule.get("to_id", "")] ==
                            [current_values[1], current_values[2], current_values[3],
                             current_values[4], current_values[5]]):
                        if column_name == "文件名":
                            rule["filename"] = new_value
                        elif column_name == "试样描述":
                            rule["desc"] = new_value
                        elif column_name == "项目名":
                            rule["project_name"] = new_value
                        elif column_name == "原ID":
                            rule["from_id"] = new_value
                        elif column_name == "目标ID":
                            rule["to_id"] = new_value
                        break

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

        # 检查是否已经是第一个
        if index == 0:
            messagebox.showinfo("提示", "已经是第一个规则，无法上移")
            return

        # 获取当前模式
        mode = self.switch_mode.get()
        if mode == "default":
            rules_list = self.default_rules
        else:
            rules_list = self.filename_rules

        # 交换规则位置
        rules_list[index], rules_list[index - 1] = rules_list[index - 1], rules_list[index]

        # 标记已修改
        self.app.mark_modified()

        # 更新显示
        self.refresh_rules_tree()

        # 重新选中移动后的规则
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

        # 获取当前模式
        mode = self.switch_mode.get()
        if mode == "default":
            rules_list = self.default_rules
        else:
            rules_list = self.filename_rules

        # 检查是否已经是最后一个
        if index == len(rules_list) - 1:
            messagebox.showinfo("提示", "已经是最后一个规则，无法下移")
            return

        # 交换规则位置
        rules_list[index], rules_list[index + 1] = rules_list[index + 1], rules_list[index]

        # 标记已修改
        self.app.mark_modified()

        # 更新显示
        self.refresh_rules_tree()

        # 重新选中移动后的规则
        self.rules_tree.selection_set(self.rules_tree.get_children()[index + 1])

    def on_mode_change(self, *args):
        # 如果选择了"全部"复选框，取消选中
        if self.show_all_var.get():
            self.show_all_var.set(False)
        self.update_rules_display()

    def on_show_all_change(self):
        # 如果选择了"全部"复选框，禁用单选按钮
        if self.show_all_var.get():
            # 禁用单选按钮
            for widget in self.show_all_check.master.winfo_children():
                if isinstance(widget, ttk.Radiobutton):
                    widget.configure(state='disabled')
            # 禁用添加规则按钮
            self.add_rule_btn.configure(state='disabled')
            # 禁用上移下移按钮
            self.move_up_btn.configure(state='disabled')
            self.move_down_btn.configure(state='disabled')
        else:
            # 启用单选按钮
            for widget in self.show_all_check.master.winfo_children():
                if isinstance(widget, ttk.Radiobutton):
                    widget.configure(state='normal')
            # 启用添加规则按钮
            self.add_rule_btn.configure(state='normal')
            # 启用上移下移按钮
            self.move_up_btn.configure(state='normal')
            self.move_down_btn.configure(state='normal')

        self.update_rules_display()

    def update_rules_display(self):
        mode = self.switch_mode.get()
        show_all = self.show_all_var.get()

        if show_all:
            # 在全部模式下，隐藏所有输入控件
            self.filename_label.grid_remove()
            self.filename_entry.grid_remove()
            self.project_name_label.grid_remove()
            self.project_name_entry.grid_remove()
            self.from_id_label.grid_remove()
            self.from_id_entry.grid_remove()
            self.to_id_label.grid_remove()
            self.to_id_entry.grid_remove()
        else:
            # 显示ID输入控件
            self.from_id_label.grid(row=0, column=4, padx=(0, 0), pady=1, sticky='w')
            self.from_id_entry.grid(row=0, column=5, padx=(0, 5), pady=1, sticky='w')
            self.to_id_label.grid(row=0, column=6, padx=(0, 0), pady=1, sticky='w')
            self.to_id_entry.grid(row=0, column=7, padx=(0, 5), pady=1, sticky='w')

            # 根据模式调整输入框显示
            if mode == "filename":
                # 名称切换：文件名关键字 + 试样描述关键字(均非空时需同时命中)；项目名输入框位复用为试样描述
                self.filename_label.configure(text="文件名:")
                self.filename_label.grid(row=0, column=0, padx=(0, 0), pady=1, sticky='w')
                self.filename_entry.grid(row=0, column=1, padx=(0, 5), pady=1, sticky='w')
                self.project_name_label.configure(text="试样描述:")
                self.project_name_label.grid(row=0, column=2, padx=(0, 0), pady=1, sticky='w')
                self.project_name_entry.grid(row=0, column=3, padx=(0, 5), pady=1, sticky='w')
            else:
                # 默认切换：按项目名匹配(default_rules.project_name)
                self.filename_label.grid_remove()
                self.filename_entry.grid_remove()
                self.project_name_label.configure(text="项目名:")
                self.project_name_label.grid(row=0, column=0, padx=(0, 0), pady=1, sticky='w')
                self.project_name_entry.grid(row=0, column=1, padx=(0, 5), pady=1, sticky='w')

        # 更新规则显示
        self.refresh_rules_tree()

    def refresh_rules_tree(self):
        # 确保规则表格已创建
        if not hasattr(self, 'rules_tree'):
            return

        # 清空现有规则显示
        for item in self.rules_tree.get_children():
            self.rules_tree.delete(item)

        # 根据当前模式显示相应规则
        show_all = self.show_all_var.get()

        if show_all:
            # 显示所有规则
            all_rules = []
            # 添加默认切换规则
            for rule in self.default_rules:
                all_rules.append({
                    "mode": "默认",
                    "filename": "",
                    "desc": "",
                    "project_name": rule.get("project_name", ""),
                    "from_id": rule["from_id"],
                    "to_id": rule["to_id"]
                })
            # 添加名称切换规则(文件名 + 试样描述关键字)
            for rule in self.filename_rules:
                all_rules.append({
                    "mode": "名称",
                    "filename": rule.get("filename", ""),
                    "desc": rule.get("desc", ""),
                    "project_name": rule.get("project_name", ""),
                    "from_id": rule["from_id"],
                    "to_id": rule["to_id"]
                })

            # 添加规则到表格
            for rule in all_rules:
                self.rules_tree.insert("", "end", values=(
                    rule["mode"],
                    rule["filename"],
                    rule["desc"],
                    rule["project_name"],
                    rule["from_id"],
                    rule["to_id"]
                ))
        else:
            # 显示当前模式的规则
            mode = self.switch_mode.get()
            rules = self.default_rules if mode == "default" else self.filename_rules

            # 添加规则到表格
            for rule in rules:
                if mode == "default":
                    self.rules_tree.insert("", "end", values=(
                        "默认",
                        "",
                        "",
                        rule.get("project_name", ""),
                        rule["from_id"],
                        rule["to_id"]
                    ))
                else:
                    self.rules_tree.insert("", "end", values=(
                        "名称",
                        rule.get("filename", ""),
                        rule.get("desc", ""),
                        rule.get("project_name", ""),
                        rule["from_id"],
                        rule["to_id"]
                    ))

    def add_rule(self):
        # 获取输入值
        from_id = self.from_id_entry.get().strip()
        to_id = self.to_id_entry.get().strip()

        if not from_id or not to_id:
            messagebox.showwarning("输入错误", "请填写原方法ID和目标方法ID")
            return

        # 根据当前模式添加规则
        mode = self.switch_mode.get()
        if mode == "default":
            project_name = self.project_name_entry.get().strip()
            rule = {"from_id": from_id, "to_id": to_id, "project_name": project_name}
            self.default_rules.append(rule)
        else:
            filename = self.filename_entry.get().strip()
            # 名称切换：项目名输入框位复用为"试样描述"关键字(可多个逗号分隔)
            desc = self.project_name_entry.get().strip()
            # 验证输入 - 至少需要文件名或试样描述关键字中的一个
            if not filename and not desc:
                messagebox.showwarning("输入错误", "请至少填写文件名或试样描述关键字")
                return

            rule = {
                "from_id": from_id,
                "to_id": to_id,
                "filename": filename,
                "desc": desc
            }

            self.filename_rules.append(rule)

        # 标记已修改
        self.app.mark_modified()

        # 清空输入框
        self.from_id_entry.delete(0, tk.END)
        self.to_id_entry.delete(0, tk.END)
        if mode == "filename":
            self.filename_entry.delete(0, tk.END)
            self.project_name_entry.delete(0, tk.END)

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

        show_all = self.show_all_var.get()

        if show_all:
            # 在全部模式下，需要确定每个规则属于哪种模式
            for item in selected_items:
                values = self.rules_tree.item(item, "values")
                mode = values[0]  # 模式列

                # 重新计算实际索引，因为显示的是合并后的列表
                if mode == "默认":
                    # 在默认规则列表中查找匹配的规则
                    target_project_name = values[3]
                    target_from_id = values[4]
                    target_to_id = values[5]
                    for i, rule in enumerate(self.default_rules):
                        if (rule.get("project_name", "") == target_project_name and
                                rule["from_id"] == target_from_id and
                                rule["to_id"] == target_to_id):
                            del self.default_rules[i]
                            break
                else:  # 名称
                    # 在名称规则列表中查找匹配的规则(文件名/试样描述/项目名/原ID/目标ID)
                    target_filename = values[1]
                    target_desc = values[2]
                    target_project_name = values[3]
                    target_from_id = values[4]
                    target_to_id = values[5]
                    for i, rule in enumerate(self.filename_rules):
                        if (rule.get("filename", "") == target_filename and
                                rule.get("desc", "") == target_desc and
                                rule.get("project_name", "") == target_project_name and
                                rule["from_id"] == target_from_id and
                                rule["to_id"] == target_to_id):
                            del self.filename_rules[i]
                            break
        else:
            # 获取当前模式
            mode = self.switch_mode.get()

            # 删除选中的规则
            for item in selected_items:
                # 由于没有序号列，我们需要根据规则内容来删除
                values = self.rules_tree.item(item, "values")
                target_from_id = values[4]
                target_to_id = values[5]

                if mode == "default":
                    target_project_name = values[3]
                    for i, rule in enumerate(self.default_rules):
                        if (rule.get("project_name", "") == target_project_name and
                                rule["from_id"] == target_from_id and
                                rule["to_id"] == target_to_id):
                            del self.default_rules[i]
                            break
                else:
                    target_filename = values[1]
                    target_desc = values[2]
                    target_project_name = values[3]
                    for i, rule in enumerate(self.filename_rules):
                        if (rule.get("filename", "") == target_filename and
                                rule.get("desc", "") == target_desc and
                                rule.get("project_name", "") == target_project_name and
                                rule["from_id"] == target_from_id and
                                rule["to_id"] == target_to_id):
                            del self.filename_rules[i]
                            break

        # 标记已修改
        self.app.mark_modified()

        # 更新显示
        self.refresh_rules_tree()

    def get_rules(self):
        """获取方法切换规则"""
        return {
            "default_rules": self.default_rules,
            "filename_rules": self.filename_rules
        }

    def set_rules(self, default_rules, filename_rules):
        """设置方法切换规则"""
        self.default_rules = default_rules
        self.filename_rules = filename_rules
        self.refresh_rules_tree()


class WeighingTab:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.processing_rules = []
        self.weighing_params = {
            "decimal_places": "",
            "min_value": "",
            "max_value": "",
            "conversion_factor": "",
            "result_decimal_places": "",
            "weighing_mode": "random"
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

        # 创建内容区域 - 所有模式的内容都显示
        self.weighing_content_frame = ttk.Frame(weighing_frame)
        self.weighing_content_frame.pack(fill='both', expand=True, padx=5, pady=5)

        # 创建所有模式的内容区域
        self.create_all_weighing_modes()

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
        self.processing_rules_tree = ttk.Treeview(rules_list_frame, columns=columns, show="headings", height=5)

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

    def get_rules_and_params(self):
        """获取处理规则和称样量参数"""
        # 更新称样量参数
        self.weighing_params = {
            "decimal_places": self.decimal_places.get(),
            "min_value": self.min_value.get(),
            "max_value": self.max_value.get(),
            "conversion_factor": self.conversion_factor.get(),
            "result_decimal_places": self.result_decimal_places.get(),
            "weighing_mode": self.weighing_mode.get()
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
        self.quantitative_report_path = r"\\192.168.12.232\ElimsEquipIOTSMO\CIRS-Equip"
        self.undetected_threshold = ""
        self.marker = ""  # 新增：标记物参数
        self.clear_spectrum = False  # 新增：录入前是否清空谱图
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

        # 数据采集单选按钮（与"本地上传"互斥）
        self.data_acquisition_radio = ttkb.Radiobutton(
            mode_selection_frame,
            text="数据采集",
            variable=self.mode_var,
            value="data_acquisition",
            command=self.on_mode_change, bootstyle="primary")
        self.data_acquisition_radio.pack(side='left', padx=(0, 20))

        # 新增：本地+采集单选按钮
        self.local_acquisition_radio = ttkb.Radiobutton(
            mode_selection_frame,
            text="本地+采集",
            variable=self.mode_var,
            value="local_acquisition",
            command=self.on_mode_change, bootstyle="primary")
        self.local_acquisition_radio.pack(side='left')

        # 无需谱图（与上述模式互斥）
        self.no_spectrum_radio = ttkb.Radiobutton(
            mode_selection_frame,
            text="无需谱图",
            variable=self.mode_var,
            value="no_spectrum",
            command=self.on_mode_change, bootstyle="primary")
        self.no_spectrum_radio.pack(side='left', padx=(20, 0))

        # 是否清空谱图：选中后序列运行录入数据前调 deleteSpectrumByProjectIds 删除谱图再录入
        self.clear_spectrum_var = tk.BooleanVar(value=self.clear_spectrum)
        self.clear_spectrum_check = ttkb.Checkbutton(
            mode_selection_frame,
            text="录入前清空谱图",
            variable=self.clear_spectrum_var,
            command=self.on_clear_spectrum_change, bootstyle="primary")
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

        # 新增参数设置区域 - 放在底部
        self.create_additional_params_section(main_content_frame)

    def create_spectrum_check_section(self, parent):
        """创建谱图检查设置区域"""
        # 创建谱图检查框架
        check_frame = ttk.LabelFrame(parent, text="谱图检查", padding=0)
        check_frame.pack(fill='x', pady=(0, 5))  # 顶部不留边距，底部留5像素

        # 配置列的权重，确保与关键字区域对齐
        check_frame.columnconfigure(0, weight=0)  # 复选框列
        check_frame.columnconfigure(1, weight=0)  # 数量输入框列
        check_frame.columnconfigure(2, weight=1)  # 关键字输入框列，占据剩余空间

        # 创建紧凑的表格布局
        # 表头 - 使用网格布局确保与下方对齐
        header_frame = ttk.Frame(check_frame)
        header_frame.grid(row=0, column=0, columnspan=3, sticky='ew', pady=(0, 5))

        # 配置表头的列权重
        header_frame.columnconfigure(0, weight=0)
        header_frame.columnconfigure(1, weight=0)
        header_frame.columnconfigure(2, weight=1)

        # 计算表头各列的起始位置，确保与下方输入框对齐
        type_x = 10  # 与复选框对齐
        count_x = 70  # 与空白谱图输入框对齐
        keyword_x = count_x + 80  # 数量输入框宽度约80，关键字列紧随其后

        ttk.Label(header_frame, text="类型").grid(row=0, column=0, padx=(type_x, 0), sticky='w')
        ttk.Label(header_frame, text="数量").grid(row=0, column=1, padx=(count_x - type_x - 40, 0), sticky='w')  # 调整位置
        ttk.Label(header_frame, text="关键字").grid(row=0, column=2, padx=(keyword_x - count_x - 40, 10),
                                                    sticky='w')  # 调整位置

        # 空白谱图检查
        self.create_compact_check_row(check_frame, "blank", "空白", 1)

        # 标液谱图检查
        self.create_compact_check_row(check_frame, "standard", "标液", 2)

        # 线性谱图检查
        self.create_compact_check_row(check_frame, "linearity", "线性", 3)

        # 样品谱图检查 - 设置更大的下边距
        self.create_compact_check_row(check_frame, "sample", "样品", 4, bottom_margin=10)

    def create_additional_params_section(self, parent):
        """创建新增参数设置区域"""
        # 创建参数框架
        params_frame = ttk.LabelFrame(parent, text="采集参数", padding=0)
        params_frame.pack(fill='x', pady=0)  # 顶部不留边距

        # 使用网格布局实现输入框两端对齐
        params_frame.columnconfigure(0, weight=0)  # 标签列，不扩展
        params_frame.columnconfigure(1, weight=1)  # 输入框列，占据剩余空间

        # 定量报告上传地址 - 第一行
        ttk.Label(params_frame, text="报告上传IP:").grid(row=0, column=0, padx=(5, 5), pady=2, sticky='w')
        self.report_path_var = tk.StringVar(value=self.quantitative_report_path)
        self.report_path_entry = ttk.Entry(params_frame, textvariable=self.report_path_var)
        self.report_path_entry.grid(row=0, column=1, padx=(0, 10), pady=2, sticky='ew')
        self.report_path_entry.bind('<KeyRelease>', self.on_report_path_change)

        # 未检出判断值 - 第二行
        ttk.Label(params_frame, text="N.D判断值:").grid(row=1, column=0, padx=(5, 5), pady=2, sticky='w')
        self.threshold_var = tk.StringVar(value=self.undetected_threshold)
        self.threshold_entry = ttk.Entry(params_frame, textvariable=self.threshold_var)
        self.threshold_entry.grid(row=1, column=1, padx=(0, 10), pady=2, sticky='ew')
        self.threshold_entry.bind('<KeyRelease>', self.on_threshold_change)

        # 标记物 - 第三行
        ttk.Label(params_frame, text="标记物:").grid(row=2, column=0, padx=(5, 5), pady=2, sticky='w')
        self.marker_var = tk.StringVar(value=self.marker)
        self.marker_entry = ttk.Entry(params_frame, textvariable=self.marker_var)
        self.marker_entry.grid(row=2, column=1, padx=(0, 10), pady=2, sticky='ew')
        self.marker_entry.bind('<KeyRelease>', self.on_marker_change)

    def create_compact_check_row(self, parent, param_type, display_name, row_num, bottom_margin=None):
        """创建紧凑的谱图检查行"""
        # 复选框
        check_var = tk.BooleanVar(value=self.spectrum_check_params[param_type]["enabled"])
        check_button = ttkb.Checkbutton(
            parent,
            text=display_name,
            variable=check_var,
            command=lambda: self.on_spectrum_check_change(param_type, check_var), bootstyle="primary")

        # 如果是最后一行且有指定的底部边距，使用不同的pady值
        if bottom_margin is not None and row_num == 4:
            pady_value = (2, bottom_margin)  # 上边距5，下边距使用指定值
        else:
            pady_value = 2  # 普通行的边距

        check_button.grid(row=row_num, column=0, padx=(10, 5), pady=pady_value, sticky='w')

        # 数量输入框
        count_var = tk.StringVar(value=self.spectrum_check_params[param_type]["count"])
        count_entry = ttk.Entry(parent, textvariable=count_var, width=8)
        count_entry.grid(row=row_num, column=1, padx=(0, 5), pady=pady_value, sticky='w')
        count_entry.bind('<KeyRelease>', lambda e: self.on_spectrum_param_change(param_type, 'count', count_var.get()))

        # 关键字输入框
        keyword_var = tk.StringVar(value=self.spectrum_check_params[param_type]["keyword"])
        keyword_entry = ttk.Entry(parent, textvariable=keyword_var)
        keyword_entry.grid(row=row_num, column=2, padx=(0, 10), pady=pady_value, sticky='ew')
        keyword_entry.bind('<KeyRelease>', lambda e: self.on_spectrum_param_change(param_type, 'keyword', keyword_var.get()))

        # 保存引用
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

    def on_report_path_change(self, event=None):
        """定量报告上传地址改变"""
        self.quantitative_report_path = self.report_path_var.get()
        self.app.mark_modified()

    def on_threshold_change(self, event=None):
        """未检出判断值改变"""
        self.undetected_threshold = self.threshold_var.get()
        self.app.mark_modified()

    def on_marker_change(self, event=None):
        """标记物改变"""
        self.marker = self.marker_var.get()
        self.app.mark_modified()

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

        # 根据当前选择的模式启用或禁用本地上传模式区域
        if mode == "no_spectrum":
            # 无需谱图：禁用全部参数
            self.set_local_upload_frame_state("disabled")
            self.report_path_entry.config(state="disabled")
            self.threshold_entry.config(state="disabled")
            self.marker_entry.config(state="disabled")
            self.local_upload_frame.configure(text="本地参数设置(禁用)")
        elif mode == "local_upload" or mode == "local_acquisition":
            # 本地上传和本地+采集模式都启用本地上传模式区域
            self.set_local_upload_frame_state("normal")

            # 只有在本地+采集模式下才启用采集参数
            if mode == "local_acquisition":
                self.report_path_entry.config(state="normal")
                self.threshold_entry.config(state="normal")
                self.marker_entry.config(state="normal")
            else:
                self.report_path_entry.config(state="disabled")
                self.threshold_entry.config(state="disabled")
                self.marker_entry.config(state="disabled")

            # 更新标签框标题
            self.local_upload_frame.configure(text="本地参数设置")
        else:  # data_acquisition
            # 数据采集模式：启用报告上传IP，禁用其他采集参数
            self.set_local_upload_frame_state("disabled")
            # 但特别启用报告上传IP输入框
            self.report_path_entry.config(state="normal")

            # 更新标签框标题，添加"(禁用)"提示
            self.local_upload_frame.configure(text="本地参数设置(禁用)")

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
            "quantitative_report_path": self.quantitative_report_path,
            "undetected_threshold": self.undetected_threshold,
            "marker": self.marker,  # 新增标记物参数
            "clear_spectrum": self.clear_spectrum  # 新增：录入前是否清空谱图
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

        # 设置新增参数
        if "quantitative_report_path" in settings:
            self.quantitative_report_path = settings["quantitative_report_path"]
            self.report_path_var.set(self.quantitative_report_path)

        if "undetected_threshold" in settings:
            self.undetected_threshold = settings["undetected_threshold"]
            self.threshold_var.set(self.undetected_threshold)

        # 新增：设置标记物参数
        if "marker" in settings:
            self.marker = settings["marker"]
            self.marker_var.set(self.marker)

        # 新增：设置录入前清空谱图
        if "clear_spectrum" in settings:
            self.clear_spectrum = bool(settings["clear_spectrum"])
            self.clear_spectrum_var.set(self.clear_spectrum)

        # 更新显示
        self.on_mode_change()


class QueryAppFixed:
    def __init__(self, root):
        self.root = root
        self.root.title("录入方法编辑器_未加载配置文件")
        self.root.geometry("1400x980")  # 增大窗口以完整显示各标签页内容

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
        self.weighing_tab = WeighingTab(self.notebook, self)
        self.spectrum_tab = SpectrumUploadTab(self.notebook, self)
        self.other_params_tab = OtherParamsTab(self.notebook, self)  # 新增其他参数标签页

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

    def convert_old_rules(self, filename_rules):
        """将旧的规则格式转换为新的格式"""
        converted_rules = []
        for rule in filename_rules:
            # 创建新规则对象
            new_rule = {
                "from_id": rule["from_id"],
                "to_id": rule["to_id"],
                "description": rule.get("description", ""),
                "desc": rule.get("desc", "")
            }

            # 处理名称字段 - 优先使用新格式，如果没有则使用旧格式
            if "name" in rule:
                new_rule["filename"] = rule["name"]
                # 如果旧格式有name_type，将其转换为project_name
                if rule.get("name_type") == "项目名":
                    new_rule["project_name"] = rule["name"]
                    new_rule["filename"] = ""
            elif "filename" in rule:
                new_rule["filename"] = rule["filename"]
            else:
                # 如果既没有name也没有filename，设为空
                new_rule["filename"] = ""

            # 处理项目名字段
            if "project_name" in rule:
                new_rule["project_name"] = rule["project_name"]
            else:
                new_rule["project_name"] = ""

            converted_rules.append(new_rule)
        return converted_rules

    def silent_load_rules(self):
        """静默加载规则，不显示任何弹窗"""
        try:
            # 首先尝试加载上次打开的配置文件
            last_config = self.load_last_config_path()
            if last_config and os.path.exists(last_config):
                self.current_config_file = last_config
                self.config_file_name = os.path.basename(last_config)

                with open(last_config, 'r', encoding='utf-8') as f:
                    rules_data = yaml.safe_load(f)

                # 加载方法切换规则
                default_rules = rules_data.get("default_rules", [])
                filename_rules = rules_data.get("filename_rules", [])

                # 转换旧格式规则
                filename_rules = self.convert_old_rules(filename_rules)

                self.method_tab.set_rules(default_rules, filename_rules)

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
            default_config = "switch_rules.yaml"
            if os.path.exists(default_config):
                self.current_config_file = default_config
                self.config_file_name = os.path.basename(default_config)

                with open(default_config, 'r', encoding='utf-8') as f:
                    rules_data = yaml.safe_load(f)

                # 加载方法切换规则
                default_rules = rules_data.get("default_rules", [])
                filename_rules = rules_data.get("filename_rules", [])

                # 转换旧格式规则
                filename_rules = self.convert_old_rules(filename_rules)

                self.method_tab.set_rules(default_rules, filename_rules)

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
            with open(file_path, 'r', encoding='utf-8') as f:
                rules_data = yaml.safe_load(f) or {}
        except Exception as e:
            messagebox.showerror("错误", f"加载方法文件失败:\n{e}")
            return False
        self.current_config_file = file_path
        self.config_file_name = os.path.basename(file_path)
        # ponytail: 填充逻辑与 silent_load_rules 一致，未抽取共用以免改动正常工作的代码
        default_rules = rules_data.get("default_rules", [])
        filename_rules = self.convert_old_rules(rules_data.get("filename_rules", []))
        self.method_tab.set_rules(default_rules, filename_rules)
        self.query_tab.set_rules(rules_data.get("query_rules", []))
        self.weighing_tab.set_rules_and_params(
            rules_data.get("processing_rules", []), rules_data.get("weighing_params", {}))
        self.spectrum_tab.set_settings(rules_data.get("spectrum_upload_settings", {}))
        self.other_params_tab.set_settings(rules_data.get("other_params_settings", {}))
        self.update_window_title()
        self.modified = False
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

        # 文件菜单：加载、保存、另存为
        file_menu = tk.Menu(menubar, tearoff=False, **menu_opts)
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
        self.method_tab.switch_mode.set("default")
        self.method_tab.show_all_var.set(True)
        self.method_tab.from_id_entry.delete(0, tk.END)
        self.method_tab.to_id_entry.delete(0, tk.END)
        self.method_tab.filename_entry.delete(0, tk.END)
        self.method_tab.project_name_entry.delete(0, tk.END)
        self.query_tab.query_project_entry.delete(0, tk.END)
        self.query_tab.query_method_entry.delete(0, tk.END)
        self.query_tab.query_cancel_test_var.set(False)
        self.query_tab.query_max_select_entry.delete(0, tk.END)
        self.query_tab.query_mode_var.set("方法")
        self.method_tab.update_rules_display()

    def save_all_rules(self):
        """保存所有规则到当前配置文件（使用YAML格式）"""
        # 验证谱图检查参数
        spectrum_errors = self.spectrum_tab.validate_spectrum_check_params()
        if spectrum_errors:
            error_message = "谱图检查参数错误:\n" + "\n".join(spectrum_errors)
            messagebox.showwarning("保存失败", error_message)
            return

        if not self.current_config_file:
            # 如果没有当前配置文件，使用默认名称
            self.current_config_file = "switch_rules.yaml"
            self.config_file_name = "switch_rules.yaml"

        # 获取各标签页的规则
        method_rules = self.method_tab.get_rules()
        query_rules = self.query_tab.get_rules()
        weighing_data = self.weighing_tab.get_rules_and_params()
        spectrum_settings = self.spectrum_tab.get_settings()
        other_params_settings = self.other_params_tab.get_settings()  # 新增其他参数设置

        rules_data = {
            "default_rules": method_rules["default_rules"],
            "filename_rules": method_rules["filename_rules"],
            "query_rules": query_rules,
            "processing_rules": weighing_data["processing_rules"],
            "weighing_params": weighing_data["weighing_params"],
            "spectrum_upload_settings": spectrum_settings,
            "other_params_settings": other_params_settings  # 新增其他参数设置
        }

        try:
            with open(self.current_config_file, 'w', encoding='utf-8') as f:
                yaml.dump(rules_data, f, allow_unicode=True, indent=2, sort_keys=False)

            # 保存后重置修改标记
            self.modified = False
            self.update_window_title()
            self.save_last_config_path()
        except Exception as e:
            messagebox.showerror("保存失败", f"保存规则时出错: {str(e)}")

    def save_as_rules(self):
        """另存所有规则到新的配置文件（使用YAML格式）"""
        file_path = filedialog.asksaveasfilename(
            title="另存配置文件",
            defaultextension=".yaml",
            filetypes=[("YAML files", "*.yaml"), ("All files", "*.*")]
        )

        if not file_path:
            return  # 用户取消了保存

        self.current_config_file = file_path
        self.config_file_name = os.path.basename(file_path)
        self.save_all_rules()

    def load_rules(self):
        """从文件加载规则（使用YAML格式）"""
        file_path = filedialog.askopenfilename(
            title="选择配置文件",
            defaultextension=".yaml",
            filetypes=[("YAML files", "*.yaml"), ("All files", "*.*")]
        )

        if not file_path:
            return  # 用户取消了选择

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                rules_data = yaml.safe_load(f)

            # 加载方法切换规则
            default_rules = rules_data.get("default_rules", [])
            filename_rules = rules_data.get("filename_rules", [])

            # 转换旧格式规则
            filename_rules = self.convert_old_rules(filename_rules)

            self.method_tab.set_rules(default_rules, filename_rules)

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
        self.fixed_params = []  # 固定参数规则：[{trigger, param_name, param_value}]
        self.create_tab()

    def create_tab(self):
        """创建其他参数标签页"""
        other_frame = ttk.Frame(self.parent)
        self.parent.add(other_frame, text="其他参数")

        # 创建标液类型设置区域
        self.create_standard_type_section(other_frame)

        # 创建仪器设置区域
        self.create_instrument_setting_section(other_frame)

        # 创建固定参数设置区域
        self.create_fixed_params_section(other_frame)

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

        # 绑定配制序号变化事件
        self.preparation_number_var.trace('w', self.on_preparation_number_change)

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
            command=self.on_instrument_setting_change, bootstyle="primary").pack(side='left')

        # 设备编号输入区域（始终显示，但状态根据选择变化）
        self.device_frame = ttk.Frame(instrument_frame)
        self.device_frame.pack(fill='x', pady=5)  # 始终显示

        ttk.Label(self.device_frame, text="设备编号:").pack(side='left', padx=(20, 5))
        self.device_number_var = tk.StringVar()
        self.device_entry = ttk.Entry(
            self.device_frame,
            textvariable=self.device_number_var,
            width=80  # 增加宽度
        )
        self.device_entry.pack(side='left')

        # 绑定设备编号变化事件
        self.device_number_var.trace('w', self.on_device_number_change)

        # 初始状态设置
        self.on_instrument_setting_change()

    def create_fixed_params_section(self, parent):
        """创建固定参数设置区域：一个触发条件显示为一行，可含多个「参数=值」"""
        frame = ttk.LabelFrame(parent, text="固定参数设置", padding=10)
        frame.pack(fill='both', expand=True, padx=5, pady=5)

        # 输入区：触发条件(字段=值) + 参数名 + 值
        input_frame = ttk.Frame(frame)
        input_frame.pack(fill='x', pady=5)

        ttk.Label(input_frame, text="触发条件:").pack(side='left', padx=(0, 5))
        self.fp_trigger_field_var = tk.StringVar(value="检测项目")
        ttk.Combobox(input_frame, textvariable=self.fp_trigger_field_var,
                     values=["检测项目", "检测方法"], state="readonly", width=10).pack(side='left', padx=(0, 2))
        ttk.Label(input_frame, text="=").pack(side='left', padx=(0, 2))
        self.fp_trigger_value_entry = ttk.Entry(input_frame, width=14)
        self.fp_trigger_value_entry.pack(side='left', padx=(0, 12))

        ttk.Label(input_frame, text="参数名:").pack(side='left', padx=(0, 5))
        self.fp_param_name_entry = ttk.Entry(input_frame, width=18)
        self.fp_param_name_entry.pack(side='left', padx=(0, 8))

        ttk.Label(input_frame, text="值:").pack(side='left', padx=(0, 5))
        self.fp_param_value_entry = ttk.Entry(input_frame, width=12)
        self.fp_param_value_entry.pack(side='left', padx=(0, 10))

        ttkb.Button(input_frame, text="添加", command=self.add_fixed_param, bootstyle="secondary").pack(side='left')

        # 列表区：一个触发条件一行，多个参数合并显示
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill='both', expand=True, pady=5)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        columns = ("触发条件", "参数设置")
        self.fp_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=8)
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

        if not trigger_value:
            messagebox.showwarning("输入错误", "请填写触发条件值")
            return
        if not param_name:
            messagebox.showwarning("输入错误", "请填写参数名")
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
                    # 解析「名=值; 名=值」回 params 列表
                    params = []
                    for part in new_value.split(";"):
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
            # 现配现用或无需标液：禁用配制序号输入框
            self.preparation_entry.config(state="disabled")

        # 标记已修改
        self.app.mark_modified()

    def on_instrument_setting_change(self):
        """仪器设置改变时的处理"""
        if self.instrument_setting_var.get() == "specified":
            # 启用设备编号输入框
            self.device_entry.config(state="normal")
        else:
            # 禁用设备编号输入框
            self.device_entry.config(state="disabled")

        # 标记已修改
        self.app.mark_modified()

    def on_preparation_number_change(self, *args):
        """配制序号改变时的处理"""
        # 标记已修改
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
            "instrument_setting": self.instrument_setting_var.get(),
            "device_number": self.device_number_var.get().strip(),
            "fixed_params": self.fixed_params
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

        # 设置仪器设置
        if "instrument_setting" in settings:
            self.instrument_setting_var.set(settings["instrument_setting"])

        # 设置设备编号
        if "device_number" in settings:
            self.device_number_var.set(settings["device_number"])

        # 设置固定参数（兼容旧扁平格式：按 trigger 合并为一行）
        self.fixed_params = self._normalize_fixed_params(settings.get("fixed_params", []))
        self.refresh_fixed_params_tree()

        # 更新显示
        self.on_standard_type_change()
        self.on_instrument_setting_change()


if __name__ == "__main__":
    root = ttkb.Window(themename="sandstone-light")
    app = QueryAppFixed(root)
    root.mainloop()