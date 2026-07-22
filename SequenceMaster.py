import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import json
import glob
import threading
import queue
import types
import yaml
from datetime import datetime
from PIL import ImageTk

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from login import MultiUserLoginSystem
from detection_entry_api import DetectionAPI, build_grouped_experiment_data, _Box


class DraggableHeader:
    """可拖动的表头分隔线"""

    def __init__(self, parent, index, on_drag):
        self.parent = parent
        self.index = index
        self.on_drag = on_drag
        self.dragging = False

        # 创建分隔线 - 使用绝对定位
        self.separator = tk.Frame(parent, width=1, cursor="sb_h_double_arrow", bg="#e0e0e0")
        self.separator.place(in_=parent, x=0, y=0, relheight=1)

        # 绑定事件
        self.separator.bind("<ButtonPress-1>", self.on_press)
        self.separator.bind("<B1-Motion>", self.on_drag_motion)
        self.separator.bind("<ButtonRelease-1>", self.on_release)

        # 鼠标悬停效果
        self.separator.bind("<Enter>", self.on_enter)
        self.separator.bind("<Leave>", self.on_leave)

    def update_position(self, x_position):
        """更新分隔线位置 - 使用绝对坐标"""
        self.separator.place_configure(x=x_position)

    def on_enter(self, event):
        """鼠标进入分隔线区域"""
        self.separator.configure(bg="#3498db", width=2)

    def on_leave(self, event):
        """鼠标离开分隔线区域"""
        if not self.dragging:
            self.separator.configure(bg="#e0e0e0", width=1)

    def on_press(self, event):
        """鼠标按下分隔线"""
        self.dragging = True
        self.start_x = event.x_root
        self.separator.configure(bg="#2980b9", width=2)

    def on_drag_motion(self, event):
        """拖动分隔线"""
        if self.dragging:
            delta = event.x_root - self.start_x
            self.start_x = event.x_root
            self.on_drag(self.index, delta)

    def on_release(self, event):
        """释放鼠标"""
        self.dragging = False
        self.separator.configure(bg="#e0e0e0", width=4)


class CellSelectionManager:
    """单元格选择管理器"""

    def __init__(self):
        self.selected_cells = set()  # 存储选中的单元格 (row, col)
        self.selected_rows = set()  # 存储选中的整行
        self.drag_start = None  # 拖动起始位置
        self.dragging = False  # 是否正在拖动
        self.drag_type = None  # 拖动类型: 'row' 或 'cell'
        self.drag_column = None  # 拖动起始列（用于限制列范围）

    def clear_all_selection(self):
        """清除所有选择"""
        self.selected_cells.clear()
        self.selected_rows.clear()

    def select_cell(self, row, col):
        """选择单个单元格"""
        self.selected_cells.add((row, col))

    def deselect_cell(self, row, col):
        """取消选择单元格"""
        self.selected_cells.discard((row, col))

    def select_row(self, row):
        """选择整行"""
        self.selected_rows.add(row)

    def deselect_row(self, row):
        """取消选择行"""
        self.selected_rows.discard(row)

    def is_cell_selected(self, row, col):
        """检查单元格是否被选中"""
        return (row, col) in self.selected_cells or row in self.selected_rows

    def is_row_selected(self, row):
        """检查整行是否被选中"""
        return row in self.selected_rows

    def start_drag(self, row, col, drag_type='cell'):
        """开始拖动选择"""
        self.drag_start = (row, col)
        self.dragging = True
        self.drag_type = drag_type
        self.drag_column = col  # 记录起始列

    def update_drag(self, row, col, ctrl_pressed=False, shift_pressed=False):
        """更新拖动选择"""
        if not self.dragging or self.drag_start is None:
            return

        start_row, start_col = self.drag_start

        if self.drag_type == 'row':
            # 行拖动选择
            min_row = min(start_row, row)
            max_row = max(start_row, row)

            # 如果不是Ctrl或Shift操作，清除之前的选择
            if not ctrl_pressed and not shift_pressed:
                self.selected_rows.clear()

            for r in range(min_row, max_row + 1):
                self.selected_rows.add(r)
        else:
            # 单元格拖动选择 - 限制在当前列内
            min_row = min(start_row, row)
            max_row = max(start_row, row)

            # 使用起始列，而不是当前列，确保只选中当前列
            min_col = self.drag_column
            max_col = self.drag_column

            # 如果不是Ctrl或Shift操作，清除之前的选择
            if not ctrl_pressed and not shift_pressed:
                # 只清除当前列的选择
                cells_to_remove = [(r, c) for r, c in self.selected_cells.copy() if c == self.drag_column]
                for cell in cells_to_remove:
                    self.selected_cells.discard(cell)

            # 添加新选择 - 只选择当前列
            for r in range(min_row, max_row + 1):
                self.selected_cells.add((r, self.drag_column))

    def end_drag(self):
        """结束拖动选择"""
        self.dragging = False
        self.drag_start = None
        self.drag_type = None
        self.drag_column = None


class UniversalCell:
    """通用单元格组件，支持单击选中和双击编辑"""

    def __init__(self, parent, width, row_index, col_index, initial_value,
                 selection_manager, on_data_update, on_drag_start,
                 on_drag_update, on_drag_end, on_cell_select, has_button=False, on_button_click=None):
        self.parent = parent
        self.width = width
        self.row_index = row_index
        self.col_index = col_index
        self.selection_manager = selection_manager
        self.on_data_update = on_data_update
        self.on_drag_start = on_drag_start
        self.on_drag_update = on_drag_update
        self.on_drag_end = on_drag_end
        self.on_cell_select = on_cell_select
        self.has_button = has_button
        self.on_button_click = on_button_click

        # 编辑状态
        self.editing = False
        self.backup_value = ""
        self.ignore_focus_out = False
        self.click_timer = None  # 用于处理双击事件
        self.destroyed = False  # 标记部件是否已被销毁
        self.is_double_click = False  # 标记是否为双击事件

        self.create_widgets(initial_value)

    def create_widgets(self, initial_value):
        # 创建单元格框架
        self.cell_frame = tk.Frame(self.parent, width=self.width, height=30, bg="white")
        self.cell_frame.pack_propagate(False)

        # 添加右边框
        border_right = tk.Frame(self.cell_frame, width=1, bg="#c0c0c0")
        border_right.pack(side='right', fill='y')

        # 添加下边框
        border_bottom = tk.Frame(self.cell_frame, height=1, bg="#c0c0c0")
        border_bottom.pack(side='bottom', fill='x')

        # 创建按钮（如果有）
        if self.has_button and self.on_button_click:
            self.button = tk.Button(self.cell_frame, text="...", width=3, command=self.on_button_click)
            self.button.pack(side='right', padx=2, pady=1)

        # 创建标签（显示模式）
        self.value_var = tk.StringVar(value=initial_value)
        self.label = tk.Label(self.cell_frame, textvariable=self.value_var, anchor="w", bg="white", fg="black")

        # 创建输入框（编辑模式）- 初始隐藏，移除所有边框
        self.entry = tk.Entry(self.cell_frame, textvariable=self.value_var, bg="white",
                              relief="flat", bd=0, highlightthickness=0)  # 移除所有边框

        # 绑定输入框事件
        self.entry.bind("<Return>", self.on_entry_return)
        self.entry.bind("<Escape>", self.on_entry_escape)
        self.entry.bind("<FocusOut>", self.on_entry_focus_out)
        self.entry.bind("<Button-1>", self.on_entry_click)

        # 默认显示标签
        self.label.pack(fill='both', expand=True, padx=2, pady=1)

        # 绑定鼠标事件到整个单元格框架
        self.cell_frame.bind("<Button-1>", self.on_click)
        self.cell_frame.bind("<B1-Motion>", self.on_drag)
        self.cell_frame.bind("<ButtonRelease-1>", self.on_release)
        self.cell_frame.bind("<Double-Button-1>", self.on_double_click)

        # 绑定鼠标事件到标签
        self.label.bind("<Button-1>", self.on_click)
        self.label.bind("<B1-Motion>", self.on_drag)
        self.label.bind("<ButtonRelease-1>", self.on_release)
        self.label.bind("<Double-Button-1>", self.on_double_click)

        # 绑定值变化事件
        self.value_var.trace("w", self.on_value_change)

    def on_entry_click(self, event):
        """输入框点击事件 - 阻止事件传播到单元格"""
        return "break"

    def on_value_change(self, *args):
        """值变化时更新数据"""
        if not self.destroyed:
            self.on_data_update()

    def get_value(self):
        """获取单元格值"""
        return self.value_var.get()

    def set_value(self, value):
        """设置单元格值"""
        self.value_var.set(value)

    def on_click(self, event):
        """单击事件 - 选择单元格"""
        if self.destroyed:
            return "break"

        # 取消之前的计时器（如果有）
        if self.click_timer:
            self.cell_frame.after_cancel(self.click_timer)
            self.click_timer = None

        # 设置计时器，延迟执行选择操作
        self.click_timer = self.cell_frame.after(200, self.perform_single_click)
        return "break"

    def perform_single_click(self):
        """执行单击操作"""
        if not self.destroyed:
            self.is_double_click = False
            self.on_cell_select(self.row_index, self.col_index, None)
        self.click_timer = None

    def on_double_click(self, event):
        """双击事件 - 进入编辑模式"""
        if self.destroyed:
            return "break"

        # 取消单击计时器
        if self.click_timer:
            self.cell_frame.after_cancel(self.click_timer)
            self.click_timer = None

        # 标记为双击事件
        self.is_double_click = True

        # 选择单元格
        self.on_cell_select(self.row_index, self.col_index, event)

        # 立即进入编辑模式
        self.start_editing()
        return "break"

    def on_drag(self, event):
        """拖动事件"""
        if self.destroyed:
            return "break"

        # 取消单击计时器
        if self.click_timer:
            self.cell_frame.after_cancel(self.click_timer)
            self.click_timer = None

        self.on_drag_update(self.row_index, self.col_index)
        return "break"

    def on_release(self, event):
        """释放事件"""
        if not self.destroyed:
            self.on_drag_end()
        return "break"

    def start_editing(self):
        """开始编辑"""
        if self.editing or self.destroyed:
            return

        self.editing = True
        self.backup_value = self.value_var.get()
        self.ignore_focus_out = True

        # 隐藏标签，显示输入框
        self.label.pack_forget()

        # 检查部件是否存在
        if hasattr(self, 'entry') and self.entry.winfo_exists():
            self.entry.pack(fill='both', expand=True, padx=2, pady=1)  # 增加内边距
        else:
            self.editing = False
            return

        # 设置焦点并选中文本
        if hasattr(self, 'entry') and self.entry.winfo_exists():
            self.entry.focus_set()
            self.entry.icursor(tk.END)
            self.entry.select_range(0, tk.END)

            # 强制获取焦点并确保光标可见
            self.entry.focus_force()

        # 强制更新显示
        if hasattr(self, 'entry') and self.entry.winfo_exists():
            self.entry.update_idletasks()

        # 短暂延迟后重置标志
        if hasattr(self, 'entry') and self.entry.winfo_exists():
            self.entry.after(150, self.reset_ignore_focus_out)

    def reset_ignore_focus_out(self):
        """重置忽略焦点失去标志"""
        self.ignore_focus_out = False

    def stop_editing(self, confirm=True):
        """停止编辑"""
        if not self.editing or self.destroyed:
            return

        self.editing = False

        if not confirm:
            # 取消编辑，恢复原值
            self.value_var.set(self.backup_value)

        # 隐藏输入框，显示标签
        if hasattr(self, 'entry') and self.entry.winfo_exists():
            self.entry.pack_forget()

        if hasattr(self, 'label') and self.label.winfo_exists():
            self.label.pack(fill='both', expand=True, padx=2, pady=1)

        # 通知数据更新
        self.on_data_update()

    def on_entry_return(self, event):
        """输入框回车事件"""
        if not self.destroyed:
            self.stop_editing(confirm=True)
        return "break"

    def on_entry_escape(self, event):
        """输入框ESC事件"""
        if not self.destroyed:
            self.stop_editing(confirm=False)
        return "break"

    def on_entry_focus_out(self, event):
        """输入框失去焦点事件"""
        if self.ignore_focus_out or self.destroyed:
            return

        if self.editing:
            self.stop_editing(confirm=True)

    def update_display(self):
        """更新显示状态"""
        if self.destroyed:
            return

        is_selected = self.selection_manager.is_cell_selected(self.row_index, self.col_index)
        bg_color = "#e6f3ff" if is_selected else "white"

        if hasattr(self, 'cell_frame') and self.cell_frame.winfo_exists():
            self.cell_frame.configure(bg=bg_color)

        if hasattr(self, 'label') and self.label.winfo_exists():
            self.label.configure(bg=bg_color)

        if self.editing and hasattr(self, 'entry') and self.entry.winfo_exists():
            self.entry.configure(bg=bg_color)

        if hasattr(self, 'button') and self.button.winfo_exists():
            self.button.configure(bg=bg_color)

    def destroy(self):
        """销毁部件"""
        self.destroyed = True
        if self.click_timer:
            try:
                self.cell_frame.after_cancel(self.click_timer)
            except:
                pass
            self.click_timer = None
        if hasattr(self, 'cell_frame') and self.cell_frame.winfo_exists():
            try:
                self.cell_frame.destroy()
            except:
                pass


class SelectableRow:
    """可选择的行"""

    def __init__(self, parent, index, data, column_widths, selection_manager,
                 on_path_select, on_method_select, on_spectrum_select, on_cell_select,
                 on_drag_start, on_drag_update, on_drag_end, on_data_update):
        self.parent = parent
        self.index = index
        self.data = data
        self.column_widths = column_widths
        self.selection_manager = selection_manager
        self.on_path_select = on_path_select
        self.on_method_select = on_method_select
        self.on_spectrum_select = on_spectrum_select
        self.on_cell_select = on_cell_select
        self.on_drag_start = on_drag_start
        self.on_drag_update = on_drag_update
        self.on_drag_end = on_drag_end
        self.on_data_update = on_data_update

        # 所有单元格
        self.cells = []
        self.destroyed = False  # 标记部件是否已被销毁

        self.create_widgets()

    def create_widgets(self):
        # 创建行框架
        self.row_frame = tk.Frame(self.parent, bg="#c0c0c0")
        self.row_frame.pack(fill='x', pady=0)

        # 创建单元格
        self.create_cells()

        # 更新显示
        self.update_display()

    def create_cells(self):
        # 序号列 - 单击选择整行，双击编辑序号
        id_cell = UniversalCell(
            self.row_frame, self.column_widths[0], self.index, 0,
            str(self.data["id"]), self.selection_manager,
            self.on_id_data_update, self.on_drag_start,
            self.on_drag_update, self.on_drag_end,
            lambda row, col, event: self.on_row_select(row, event)
        )
        id_cell.cell_frame.pack(side='left', fill='y')
        self.cells.append(id_cell)

        # 称样记录路径 - 单击选择单元格，双击编辑内容
        weighing_cell = UniversalCell(
            self.row_frame, self.column_widths[1], self.index, 1,
            self.data["weighing_path"], self.selection_manager,
            self.on_weighing_data_update, self.on_drag_start,
            self.on_drag_update, self.on_drag_end,
            self.on_cell_select,
            has_button=True, on_button_click=lambda: self.on_weighing_button_click()
        )
        weighing_cell.cell_frame.pack(side='left', fill='y')
        self.cells.append(weighing_cell)

        # 录入方法文件 - 单击选择单元格，双击编辑内容
        method_cell = UniversalCell(
            self.row_frame, self.column_widths[2], self.index, 2,
            self.data["method_file"], self.selection_manager,
            self.on_method_data_update, self.on_drag_start,
            self.on_drag_update, self.on_drag_end,
            self.on_cell_select,
            has_button=True, on_button_click=lambda: self.on_method_button_click()
        )
        method_cell.cell_frame.pack(side='left', fill='y')
        self.cells.append(method_cell)

        # 标准物质 - 单击选择单元格，双击编辑内容
        reference_cell = UniversalCell(
            self.row_frame, self.column_widths[3], self.index, 3,
            self.data["reference_material"], self.selection_manager,
            self.on_reference_data_update, self.on_drag_start,
            self.on_drag_update, self.on_drag_end,
            self.on_cell_select
        )
        reference_cell.cell_frame.pack(side='left', fill='y')
        self.cells.append(reference_cell)

        # 设备 - 单击选择单元格，双击编辑内容
        equipment_cell = UniversalCell(
            self.row_frame, self.column_widths[4], self.index, 4,
            self.data["equipment"], self.selection_manager,
            self.on_equipment_data_update, self.on_drag_start,
            self.on_drag_update, self.on_drag_end,
            self.on_cell_select
        )
        equipment_cell.cell_frame.pack(side='left', fill='y')
        self.cells.append(equipment_cell)

        # 谱图文件路径 - 单击选择单元格，双击编辑内容
        spectrum_cell = UniversalCell(
            self.row_frame, self.column_widths[5], self.index, 5,
            self.data["spectrum_path"], self.selection_manager,
            self.on_spectrum_data_update, self.on_drag_start,
            self.on_drag_update, self.on_drag_end,
            self.on_cell_select,
            has_button=True, on_button_click=lambda: self.on_spectrum_button_click()
        )
        spectrum_cell.cell_frame.pack(side='left', fill='y')
        self.cells.append(spectrum_cell)

        # 状态列（阶段1）- 非 UniversalCell，填充剩余宽度，不动 cells 列索引
        self.status_label = tk.Label(self.row_frame, text=self.data.get("status", "待运行"),
                                     anchor='center', relief='flat', padx=4,
                                     font=("Segoe UI", 9))
        self.status_label.pack(side='left', fill='x', expand=True)

    # 状态显示样式（阶段1）
    STATUS_STYLE = {
        "待运行": ("#f0f0f0", "#888888"),
        "运行中": ("#dbeafe", "#2563eb"),
        "成功":   ("#dcfce7", "#16a34a"),
        "失败":   ("#fee2e2", "#dc2626"),
        "跳过":   ("#f5f5f5", "#9ca3af"),
    }

    def set_status(self, status, error_msg=""):
        """更新本行状态显示（由主线程队列轮询调用）"""
        if self.destroyed:
            return
        self.data["status"] = status
        self.data["error_msg"] = error_msg
        bg, fg = self.STATUS_STYLE.get(status, ("#f0f0f0", "#888888"))
        text = status
        if status == "失败" and error_msg:
            short = error_msg if len(error_msg) <= 24 else error_msg[:24] + "…"
            text = f"失败:{short}"
        if self.status_label.winfo_exists():
            self.status_label.configure(text=text, bg=bg, fg=fg)

    def on_weighing_button_click(self):
        """称样记录路径按钮点击事件"""
        if self.selection_manager.is_cell_selected(self.index, 1):
            # 如果单元格被选中，进入编辑模式
            self.cells[1].start_editing()
        else:
            # 否则执行原来的路径选择功能
            self.on_path_select(self.index)

    def on_method_button_click(self):
        """方法文件按钮点击事件"""
        if self.selection_manager.is_cell_selected(self.index, 2):
            # 如果单元格被选中，进入编辑模式
            self.cells[2].start_editing()
        else:
            # 否则执行原来的路径选择功能
            self.on_method_select(self.index)

    def on_spectrum_button_click(self):
        """谱图文件路径按钮点击事件"""
        if self.selection_manager.is_cell_selected(self.index, 5):
            # 如果单元格被选中，进入编辑模式
            self.cells[5].start_editing()
        else:
            # 否则执行原来的路径选择功能
            self.on_spectrum_select(self.index)

    def on_row_select(self, row, event):
        """处理行选择（序号列）"""
        if not self.destroyed:
            self.on_drag_start(row, 0, 'row')

    def on_id_data_update(self):
        """序号数据更新"""
        if not self.destroyed:
            try:
                self.data["id"] = int(self.cells[0].get_value())
            except ValueError:
                self.cells[0].set_value(str(self.data["id"]))
            self.on_data_update()

    def on_weighing_data_update(self):
        """称样记录路径数据更新"""
        if not self.destroyed:
            self.data["weighing_path"] = self.cells[1].get_value()
            self.on_data_update()

    def on_method_data_update(self):
        """方法文件路径数据更新"""
        if not self.destroyed:
            self.data["method_file"] = self.cells[2].get_value()
            self.on_data_update()

    def on_reference_data_update(self):
        """标准物质数据更新"""
        if not self.destroyed:
            self.data["reference_material"] = self.cells[3].get_value()
            self.on_data_update()

    def on_equipment_data_update(self):
        """设备数据更新"""
        if not self.destroyed:
            self.data["equipment"] = self.cells[4].get_value()
            self.on_data_update()

    def on_spectrum_data_update(self):
        """谱图文件路径数据更新"""
        if not self.destroyed:
            self.data["spectrum_path"] = self.cells[5].get_value()
            self.on_data_update()

    def update_display(self):
        """更新显示状态"""
        if not self.destroyed:
            for cell in self.cells:
                cell.update_display()

    def destroy(self):
        """销毁行"""
        self.destroyed = True
        for cell in self.cells:
            if hasattr(cell, 'destroy'):
                cell.destroy()
        if hasattr(self, 'row_frame') and self.row_frame.winfo_exists():
            try:
                self.row_frame.destroy()
            except:
                pass


class SequenceMaster:
    def __init__(self, root):
        self.root = root
        self.root.title("SequenceMaster - 序列编辑器")
        self.root.geometry("1000x450")

        # 存储序列数据
        self.sequence_data = []

        # 选择管理器
        self.selection_manager = CellSelectionManager()

        # 键盘状态
        self.ctrl_pressed = False
        self.shift_pressed = False

        # 列宽配置
        self.column_widths = [30, 180, 180, 100, 100, 180]

        # 存储分隔线引用
        self.draggable_headers = []

        # 存储行组件引用
        self.row_widgets = []

        # LIMS 登录与 API（阶段1）
        self.login_system = MultiUserLoginSystem()
        self.api = DetectionAPI(self.login_system)
        self.logged_in = False

        # 序列运行器状态（阶段1）：1 worker 线程 + queue + 4 Event，UI 更新全部 marshal 回主线程
        self._ui_q = queue.Queue()
        self._worker = None
        self._running = False
        self._stop = threading.Event()
        self._pause = threading.Event(); self._pause.set()  # set=运行中
        self._confirm_done = threading.Event()
        self._confirm_result = None
        self._confirm_win = None

        self.create_widgets()

        # 启动 UI 队列轮询（主线程）
        self.root.after(100, self._drain_ui_queue)

        # 绑定键盘事件
        self.root.bind('<Control_L>', self.on_ctrl_press)
        self.root.bind('<Control_R>', self.on_ctrl_press)
        self.root.bind('<KeyRelease-Control_L>', self.on_ctrl_release)
        self.root.bind('<KeyRelease-Control_R>', self.on_ctrl_release)
        self.root.bind('<Shift_L>', self.on_shift_press)
        self.root.bind('<Shift_R>', self.on_shift_press)
        self.root.bind('<KeyRelease-Shift_L>', self.on_shift_release)
        self.root.bind('<KeyRelease-Shift_R>', self.on_shift_release)
        self.root.bind('<Escape>', self.clear_selection)

        # 绑定全局点击事件，用于退出编辑模式
        self.root.bind('<Button-1>', self.on_global_click)

        # 初始添加一行
        self.add_row()

    def create_widgets(self):
        # 主框架
        main_frame = tk.Frame(self.root, bg="#f0f0f0")
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # 按钮框架
        top_button_frame = tk.Frame(main_frame, bg="#f0f0f0")
        top_button_frame.pack(fill='x', pady=(0, 0))

        # 操作按钮
        button_style = {
            "relief": "flat",
            "bd": 1,
            "bg": "#f0f0f0",
            "activebackground": "#e0e0e0",
            "font": ("Segoe UI", 9)
        }

        add_btn = tk.Button(top_button_frame, text="Add", command=self.add_row, width=4, **button_style)
        add_btn.pack(side='left', padx=(0, 0))

        delete_btn = tk.Button(top_button_frame, text="✕", command=self.delete_selected_rows,
                               width=3, **button_style)
        delete_btn.pack(side='left', padx=(0, 0))

        fill_btn = tk.Button(top_button_frame, text="↓", command=self.fill_down,
                             width=3, **button_style)
        fill_btn.pack(side='left', padx=(0, 0))

        clear_btn = tk.Button(top_button_frame, text="Clear", command=self.clear_all, width=5, **button_style)
        clear_btn.pack(side='left')

        edit_method_btn = tk.Button(top_button_frame, text="📝 方法", command=self.edit_method, width=7, **button_style)
        edit_method_btn.pack(side='left', padx=(8, 0))

        # 立方登录区（右上）：登录按钮 + 登录状态
        self.login_btn = tk.Button(top_button_frame, text="立方登录", command=self._show_login_dialog, width=8, **button_style)
        self.login_btn.pack(side='right')
        self.login_status_var = tk.StringVar(value="🔴 未登录")
        self.login_lbl = tk.Label(top_button_frame, textvariable=self.login_status_var,
                                  bg="#f0f0f0", fg="#dc2626", font=("Segoe UI", 9, "bold"))
        self.login_lbl.pack(side='right', padx=(8, 4))

        # 创建表格容器
        self.create_table_container(main_frame)

        # 运行控制按钮（阶段1）
        run_ctrl = tk.Frame(main_frame, bg="#f0f0f0")
        run_ctrl.pack(fill='x', pady=(5, 0))
        tk.Label(run_ctrl, text="运行控制:", bg="#f0f0f0").pack(side='left', padx=(0, 5))
        self.pause_btn = tk.Button(run_ctrl, text="⏸ 暂停/继续", command=self.toggle_pause, width=10, state='disabled')
        self.pause_btn.pack(side='left', padx=2)
        self.skip_btn = tk.Button(run_ctrl, text="⏭ 跳过当前", command=self.skip_current, width=10, state='disabled')
        self.skip_btn.pack(side='left', padx=2)
        self.abort_btn = tk.Button(run_ctrl, text="⏹ 中止", command=self.abort_run, width=10, state='disabled')
        self.abort_btn.pack(side='left', padx=2)

        # 运行日志区（阶段1）
        log_frame = tk.LabelFrame(main_frame, text="运行日志", bg="#f0f0f0")
        log_frame.pack(fill='both', expand=False, pady=(5, 0))
        self.log_text = tk.Text(log_frame, height=7, wrap='word', state='disabled',
                                bg="#1e1e1e", fg="#d4d4d4", font=("Consolas", 9))
        log_scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side='left', fill='both', expand=True)
        log_scroll.pack(side='right', fill='y')

        # 底部按钮框架
        bottom_frame = tk.Frame(main_frame, bg="#f0f0f0")
        bottom_frame.pack(fill='x', pady=(5, 0))

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = tk.Label(bottom_frame, textvariable=self.status_var, anchor='w', bg="#f0f0f0")
        status_bar.pack(side='left', fill='x', expand=True, padx=5, pady=3)

        # 右侧按钮
        run_btn = tk.Button(bottom_frame, text="运行", command=self.run_sequence, width=5)
        run_btn.pack(side='right', padx=(5, 5), pady=3)

        load_btn = tk.Button(bottom_frame, text="加载", command=self.load_sequence, width=5)
        load_btn.pack(side='right', padx=(5, 0), pady=3)

        save_btn = tk.Button(bottom_frame, text="保存", command=self.save_sequence, width=5)
        save_btn.pack(side='right', padx=(5, 0), pady=3)

    def create_table_container(self, parent):
        """创建表格容器，包含表头和表格"""
        table_frame = tk.Frame(parent, bg="white", relief="solid", bd=1,
                               highlightbackground="#c0c0c0", highlightcolor="#c0c0c0",
                               highlightthickness=1)
        table_frame.pack(fill='both', expand=True)

        # 创建表头
        self.create_table_header(table_frame)

        # 创建表格主体
        self.create_table_body(table_frame)

    def create_table_header(self, parent):
        """创建表格表头"""
        self.header_frame = tk.Frame(parent, height=30, bg="white")
        self.header_frame.pack(fill='x')
        self.header_frame.pack_propagate(False)

        headers = [
            {"text": "No", "anchor": "center"},
            {"text": "称样记录路径", "anchor": "w"},
            {"text": "录入方法", "anchor": "w"},
            {"text": "标准物质", "anchor": "w"},
            {"text": "设备", "anchor": "w"},
            {"text": "谱图文件路径", "anchor": "w"}
        ]

        self.header_cells = []
        self.draggable_headers = []

        # 创建表头标签
        for i, header in enumerate(headers):
            cell_frame = tk.Frame(self.header_frame, height=30, bg="white")

            if i == 0:
                cell_frame.place(x=0, y=0, width=self.column_widths[i], height=30)
            else:
                prev_width = sum(self.column_widths[:i])
                cell_frame.place(x=prev_width, y=0, width=self.column_widths[i], height=30)

            # 添加边框
            border_right = tk.Frame(cell_frame, width=1, bg="#c0c0c0")
            border_right.pack(side='right', fill='y')

            border_bottom = tk.Frame(cell_frame, height=1, bg="#c0c0c0")
            border_bottom.pack(side='bottom', fill='x')

            label = tk.Label(cell_frame, text=header["text"], anchor=header["anchor"],
                             bg="white", fg="#2c3e50",
                             font=("Segoe UI", 9))
            label.pack(fill='both', expand=True, padx=4, pady=3)

            self.header_cells.append(cell_frame)

        # 创建可拖动的分隔线
        for i in range(len(headers)):
            separator = DraggableHeader(self.header_frame, i, self.on_column_drag)
            self.draggable_headers.append(separator)

            separator_x = sum(self.column_widths[:i + 1])
            separator.update_position(separator_x)
            separator.separator.lift()

    def on_column_drag(self, column_index, delta):
        """处理列宽调整"""
        if column_index < len(self.column_widths) - 1:
            self.column_widths[column_index] += delta
            self.column_widths[column_index + 1] -= delta
        else:
            self.column_widths[column_index] += delta

        # 确保最小宽度
        min_width = 50
        for i in range(len(self.column_widths)):
            if self.column_widths[i] < min_width:
                self.column_widths[i] = min_width

        # 更新表头和分隔线位置
        self.update_header_layout()

        # 更新内容行
        self.refresh_table()

    def update_header_layout(self):
        """更新表头布局"""
        for i, cell in enumerate(self.header_cells):
            if i == 0:
                cell.place_configure(x=0, width=self.column_widths[i])
            else:
                prev_width = sum(self.column_widths[:i])
                cell.place_configure(x=prev_width, width=self.column_widths[i])

        for i, separator in enumerate(self.draggable_headers):
            separator_x = sum(self.column_widths[:i + 1])
            separator.update_position(separator_x)

    def create_table_body(self, parent):
        """创建表格主体"""
        table_container = tk.Frame(parent, bg="white")
        table_container.pack(fill='both', expand=True)

        # 创建Canvas和滚动条
        self.canvas = tk.Canvas(table_container, bg='white', highlightthickness=0, height=250)
        self.scrollbar = tk.Scrollbar(table_container, orient="vertical", command=self.canvas.yview)

        # 滚动区域框架
        self.scrollable_frame = tk.Frame(self.canvas, bg="white")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # 绑定鼠标滚轮事件
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        self.scrollable_frame.bind("<MouseWheel>", self.on_mousewheel)

    def on_mousewheel(self, event):
        """处理鼠标滚轮事件"""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def add_row(self):
        """添加新行"""
        row_id = len(self.sequence_data) + 1
        new_row = {
            "id": row_id,
            "weighing_path": "",
            "method_file": "",
            "reference_material": "",
            "equipment": "",
            "spectrum_path": "",
            "sample_code": "",
            "configure_order": "",
            "status": "待运行",
            "error_msg": "",
            "experiment_code": ""
        }
        self.sequence_data.append(new_row)
        self.refresh_table()
        self.status_var.set(f"已添加第 {row_id} 行")

    def refresh_table(self):
        """刷新表格显示"""
        # 清空现有行
        for widget in self.row_widgets:
            if hasattr(widget, 'destroy'):
                widget.destroy()
        self.row_widgets = []

        # 添加新行
        for i, row_data in enumerate(self.sequence_data):
            row_widget = SelectableRow(
                self.scrollable_frame,
                i,
                row_data,
                self.column_widths,
                self.selection_manager,
                self.select_weighing_path,
                self.select_method_file,
                self.select_spectrum_path,
                self.handle_cell_select,
                self.handle_drag_start,
                self.handle_drag_update,
                self.handle_drag_end,
                self.on_data_update
            )
            self.row_widgets.append(row_widget)

    def on_data_update(self):
        """数据更新时的回调函数"""
        pass

    def on_global_click(self, event):
        """全局点击事件处理，用于退出编辑模式"""
        widget = event.widget

        # 检查是否点击了任何可编辑单元格
        for row_widget in self.row_widgets:
            for cell in row_widget.cells:
                if cell and hasattr(cell, 'entry') and widget == cell.entry:
                    return
                if cell and hasattr(cell, 'cell_frame') and widget == cell.cell_frame:
                    return
                if cell and hasattr(cell, 'label') and widget == cell.label:
                    return

        # 如果不是任何可编辑单元格，结束所有编辑状态
        for row_widget in self.row_widgets:
            for cell in row_widget.cells:
                if cell and hasattr(cell, 'editing') and cell.editing:
                    cell.stop_editing(confirm=True)

    # 键盘事件处理
    def on_ctrl_press(self, event):
        self.ctrl_pressed = True
        return "break"

    def on_ctrl_release(self, event):
        self.ctrl_pressed = False
        return "break"

    def on_shift_press(self, event):
        self.shift_pressed = True
        return "break"

    def on_shift_release(self, event):
        self.shift_pressed = False
        return "break"

    def clear_selection(self, event=None):
        """清除所有选择"""
        self.selection_manager.clear_all_selection()
        self.refresh_table()
        self.status_var.set("已清除选择")
        return "break"

    def handle_cell_select(self, row, col, event):
        """处理单元格选择"""
        # 检查是否是双击事件
        is_double_click = False
        if event and hasattr(event, 'widget'):
            # 通过检查事件来源的单元格来确定是否是双击
            for row_widget in self.row_widgets:
                if row_widget.index == row:
                    for cell in row_widget.cells:
                        if cell.col_index == col and hasattr(cell, 'is_double_click') and cell.is_double_click:
                            is_double_click = True
                            cell.is_double_click = False  # 重置标志
                            break
                    break

        if is_double_click:
            # 对于双击事件，我们直接选择单元格但不刷新表格
            self.selection_manager.clear_all_selection()
            self.selection_manager.select_cell(row, col)
            self.last_selected_cell = (row, col)

            # 只更新当前行的显示，不刷新整个表格
            if 0 <= row < len(self.row_widgets):
                self.row_widgets[row].update_display()

            self.update_status()
        else:
            # 对于单击和其他事件，使用原来的逻辑
            if self.ctrl_pressed:
                # Ctrl+点击：切换单元格选择状态
                if self.selection_manager.is_cell_selected(row, col):
                    self.selection_manager.deselect_cell(row, col)
                else:
                    self.selection_manager.select_cell(row, col)
            elif self.shift_pressed and hasattr(self, 'last_selected_cell') and self.last_selected_cell is not None:
                # Shift+点击：选择连续的单元格区域
                last_row, last_col = self.last_selected_cell
                start_row, end_row = sorted([last_row, row])
                start_col, end_col = sorted([last_col, col])

                for r in range(start_row, end_row + 1):
                    for c in range(start_col, end_col + 1):
                        self.selection_manager.select_cell(r, c)
            else:
                # 普通点击：只选择当前单元格
                self.selection_manager.clear_all_selection()
                self.selection_manager.select_cell(row, col)
                self.last_selected_cell = (row, col)

            self.refresh_table()
            self.update_status()

    def handle_drag_start(self, row, col, drag_type):
        """处理拖动开始"""
        if not self.ctrl_pressed and not self.shift_pressed:
            self.selection_manager.clear_all_selection()

        self.selection_manager.start_drag(row, col, drag_type)

        if drag_type == 'row':
            self.last_selected_row = row
            self.selection_manager.select_row(row)
        else:
            self.last_selected_cell = (row, col)
            self.selection_manager.select_cell(row, col)

        self.refresh_table()

    def handle_drag_update(self, row, col):
        """处理拖动更新"""
        self.selection_manager.update_drag(row, col, self.ctrl_pressed, self.shift_pressed)
        self.refresh_table()

    def handle_drag_end(self):
        """处理拖动结束"""
        self.selection_manager.end_drag()
        self.update_status()

    def update_status(self):
        """更新状态栏"""
        selected_rows = len(self.selection_manager.selected_rows)
        selected_cells = len(self.selection_manager.selected_cells)

        if selected_rows > 0:
            self.status_var.set(f"已选择 {selected_rows} 行")
        elif selected_cells > 0:
            self.status_var.set(f"已选择 {selected_cells} 个单元格")
        else:
            self.status_var.set("就绪")

    def select_weighing_path(self, row_index):
        """选择称样记录路径"""
        path = filedialog.askdirectory(title="选择称样记录路径")
        if path and 0 <= row_index < len(self.sequence_data):
            self.sequence_data[row_index]["weighing_path"] = path
            self.refresh_table()
            self.status_var.set(f"第 {row_index + 1} 行称样记录路径已设置")

    def select_method_file(self, row_index):
        """选择录入方法文件"""
        file_path = filedialog.askopenfilename(
            title="选择录入方法文件",
            filetypes=[("YAML files", "*.yaml"), ("配置文件", "*.yaml;*.yml;*.json"), ("All files", "*.*")]
        )
        if file_path and 0 <= row_index < len(self.sequence_data):
            self.sequence_data[row_index]["method_file"] = file_path
            self.refresh_table()
            self.status_var.set(f"第 {row_index + 1} 行录入方法文件已设置")

    def select_spectrum_path(self, row_index):
        """选择谱图文件路径"""
        path = filedialog.askdirectory(title="选择谱图文件夹路径")
        if path and 0 <= row_index < len(self.sequence_data):
            self.sequence_data[row_index]["spectrum_path"] = path
            self.refresh_table()
            self.status_var.set(f"第 {row_index + 1} 行谱图文件路径已设置")

    def delete_selected_rows(self):
        """删除选中的行"""
        selected_rows = set()
        for row in self.selection_manager.selected_rows:
            selected_rows.add(row)
        for row, col in self.selection_manager.selected_cells:
            selected_rows.add(row)

        if not selected_rows:
            messagebox.showwarning("警告", "没有选中的行")
            return

        if not messagebox.askyesno("确认删除", f"确定要删除选中的 {len(selected_rows)} 行吗？"):
            return

        for index in sorted(selected_rows, reverse=True):
            if 0 <= index < len(self.sequence_data):
                del self.sequence_data[index]

        self.selection_manager.clear_all_selection()
        self.renumber_rows()
        self.refresh_table()
        self.status_var.set(f"已删除 {len(selected_rows)} 行")

    def renumber_rows(self):
        """重新编号所有行"""
        for i, row_data in enumerate(self.sequence_data, 1):
            row_data["id"] = i

    def fill_down(self):
        """向下填充选中的内容"""
        if self.selection_manager.selected_rows:
            if len(self.selection_manager.selected_rows) > 1:
                messagebox.showwarning("警告", "只能选择一行作为填充源")
                return

            source_row = next(iter(self.selection_manager.selected_rows))

            if source_row >= len(self.sequence_data) - 1:
                messagebox.showinfo("提示", "已经是最后一行，无需向下填充")
                return

            source_data = self.sequence_data[source_row]

            for i in range(source_row + 1, len(self.sequence_data)):
                self.sequence_data[i]["weighing_path"] = source_data["weighing_path"]
                self.sequence_data[i]["method_file"] = source_data["method_file"]
                self.sequence_data[i]["reference_material"] = source_data["reference_material"]
                self.sequence_data[i]["equipment"] = source_data["equipment"]
                self.sequence_data[i]["spectrum_path"] = source_data["spectrum_path"]

            self.refresh_table()
            self.status_var.set(f"已从第 {source_row + 1} 行向下填充到第 {len(self.sequence_data)} 行")

        elif self.selection_manager.selected_cells:
            selected_cells = list(self.selection_manager.selected_cells)

            if not selected_cells:
                messagebox.showwarning("警告", "请先选择要填充的源单元格")
                return

            cells_by_column = {}
            for row, col in selected_cells:
                if col not in cells_by_column:
                    cells_by_column[col] = []
                cells_by_column[col].append(row)

            for col, rows in cells_by_column.items():
                if len(rows) != 1:
                    messagebox.showwarning("警告",
                                           f"第 {['序号', '称样记录路径', '录入方法', '标准物质', '设备', '谱图文件路径'][col]} 列只能选择一个源单元格")
                    continue

                source_row = rows[0]

                if source_row >= len(self.sequence_data) - 1:
                    messagebox.showinfo("提示", "已经是最后一行，无需向下填充")
                    continue

                if col == 0:
                    source_data = self.sequence_data[source_row]["id"]
                    for i in range(source_row + 1, len(self.sequence_data)):
                        self.sequence_data[i]["id"] = source_data
                elif col == 1:
                    source_data = self.sequence_data[source_row]["weighing_path"]
                    for i in range(source_row + 1, len(self.sequence_data)):
                        self.sequence_data[i]["weighing_path"] = source_data
                elif col == 2:
                    source_data = self.sequence_data[source_row]["method_file"]
                    for i in range(source_row + 1, len(self.sequence_data)):
                        self.sequence_data[i]["method_file"] = source_data
                elif col == 3:
                    source_data = self.sequence_data[source_row]["reference_material"]
                    for i in range(source_row + 1, len(self.sequence_data)):
                        self.sequence_data[i]["reference_material"] = source_data
                elif col == 4:
                    source_data = self.sequence_data[source_row]["equipment"]
                    for i in range(source_row + 1, len(self.sequence_data)):
                        self.sequence_data[i]["equipment"] = source_data
                elif col == 5:
                    source_data = self.sequence_data[source_row]["spectrum_path"]
                    for i in range(source_row + 1, len(self.sequence_data)):
                        self.sequence_data[i]["spectrum_path"] = source_data

            self.refresh_table()
            self.status_var.set("向下填充完成")
        else:
            messagebox.showwarning("警告", "请先选择要填充的行或单元格")

    def clear_all(self):
        """清空所有行内容"""
        if not self.sequence_data:
            messagebox.showwarning("警告", "没有数据可清空")
            return

        if not messagebox.askyesno("确认清空", "确定要清空所有行的内容吗？"):
            return

        for row_data in self.sequence_data:
            row_data["weighing_path"] = ""
            row_data["method_file"] = ""
            row_data["reference_material"] = ""
            row_data["equipment"] = ""
            row_data["spectrum_path"] = ""

        self.selection_manager.clear_all_selection()
        self.refresh_table()
        self.status_var.set("已清空所有行内容")

    def save_sequence(self):
        """保存序列到文件"""
        if not self.sequence_data:
            messagebox.showwarning("警告", "没有数据可保存")
            return

        file_path = filedialog.asksaveasfilename(
            title="保存序列文件",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("Text files", "*.txt"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            save_data = {
                "sequence_data": [],
                "column_widths": self.column_widths
            }

            for row in self.sequence_data:
                save_data["sequence_data"].append({
                    "id": row["id"],
                    "weighing_path": row["weighing_path"],
                    "method_file": row["method_file"],
                    "reference_material": row["reference_material"],
                    "equipment": row["equipment"],
                    "spectrum_path": row["spectrum_path"],
                    "sample_code": row.get("sample_code", ""),
                    "configure_order": row.get("configure_order", ""),
                    "status": row.get("status", "待运行"),
                    "error_msg": row.get("error_msg", ""),
                    "experiment_code": row.get("experiment_code", "")
                })

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(save_data, f, ensure_ascii=False, indent=2)

            self.status_var.set(f"序列已保存到: {file_path}")
            messagebox.showinfo("成功", f"序列已保存到: {file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {str(e)}")

    def load_sequence(self):
        """从文件加载序列"""
        file_path = filedialog.askopenfilename(
            title="加载序列文件",
            filetypes=[("JSON files", "*.json"), ("Text files", "*.txt"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                loaded_data = json.load(f)

            self.sequence_data = []
            self.selection_manager.clear_all_selection()

            if "column_widths" in loaded_data:
                self.column_widths = loaded_data["column_widths"]
                self.update_header_layout()

            sequence_data = loaded_data.get("sequence_data", loaded_data)
            for item in sequence_data:
                new_row = {
                    "id": item["id"],
                    "weighing_path": item["weighing_path"],
                    "method_file": item["method_file"],
                    "reference_material": item.get("reference_material", ""),
                    "equipment": item.get("equipment", ""),
                    "spectrum_path": item["spectrum_path"],
                    "sample_code": item.get("sample_code", ""),
                    "configure_order": item.get("configure_order", ""),
                    "status": item.get("status", "待运行"),
                    "error_msg": item.get("error_msg", ""),
                    "experiment_code": item.get("experiment_code", "")
                }
                self.sequence_data.append(new_row)

            self.refresh_table()
            self.status_var.set(f"已加载序列文件: {file_path}")
            messagebox.showinfo("成功", f"已加载序列文件: {file_path}")

        except Exception as e:
            messagebox.showerror("错误", f"加载失败: {str(e)}")

    def edit_method(self):
        """方法编辑入口：取选中行的方法文件并打开方法编辑器"""
        target = None
        sel = sorted(self.selection_manager.selected_rows) or sorted({r for r, _ in self.selection_manager.selected_cells})
        if sel:
            idx = sel[0]
            if 0 <= idx < len(self.sequence_data):
                mf = self.sequence_data[idx].get("method_file")
                if mf:
                    target = mf
        if not target:
            self._log("未选中含方法文件的行，将打开编辑器(可在编辑器内点'加载')")
        self._open_method_editor(target)

    def _open_method_editor(self, file_path=None):
        """用 Toplevel 打开 MethodRule Editor(带空格文件名，用 importlib 导入)，可选加载指定方法文件"""
        if hasattr(self, "_method_editor_top") and self._method_editor_top is not None and self._method_editor_top.winfo_exists():
            self._method_editor_top.lift()
            self._method_editor_top.focus_force()
            return
        import importlib.util
        editor_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "MethodRule Editor.py")
        if not os.path.exists(editor_path):
            messagebox.showerror("错误", "找不到 MethodRule Editor.py")
            return
        spec = importlib.util.spec_from_file_location("method_rule_editor", editor_path)
        mre = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mre)
        top = tk.Toplevel(self.root)
        top.title("录入方法编辑器")
        app = mre.QueryAppFixed(top)
        if file_path:
            app.load_config_file(file_path)
        self._method_editor_top = top
        top.protocol("WM_DELETE_WINDOW", lambda: (setattr(self, "_method_editor_top", None), top.destroy()))
        self._log(f"打开方法编辑器{': ' + os.path.basename(file_path) if file_path else ''}")

    def run_sequence(self):
        """运行序列（阶段1：真实逐行提交，半自动）"""
        if self._running:
            messagebox.showwarning("提示", "序列正在运行中")
            return
        if not self.sequence_data:
            messagebox.showwarning("警告", "没有可运行的序列数据")
            return
        invalid = [i + 1 for i, r in enumerate(self.sequence_data)
                   if not r["method_file"] or not r["spectrum_path"]]
        if invalid:
            messagebox.showwarning("数据不完整",
                                   f"以下行缺少方法文件或谱图路径:\n第 {', '.join(map(str, invalid))} 行")
            return
        if not messagebox.askyesno("确认运行", f"确定要运行 {len(self.sequence_data)} 行序列吗？"):
            return
        if not self.logged_in:
            if not self._show_login_dialog():
                return
        # 重置状态
        self._stop.clear()
        self._pause.set()
        for i, r in enumerate(self.sequence_data):
            r["status"] = "待运行"
            r["error_msg"] = ""
            r["experiment_code"] = ""
            if i < len(self.row_widgets):
                self.row_widgets[i].set_status("待运行")
        self._set_running(True)
        self._log(f"==== 开始运行序列，共 {len(self.sequence_data)} 行 ====")
        self._worker = threading.Thread(target=self._run_worker, daemon=True)
        self._worker.start()

    def _set_running(self, running):
        self._running = running
        self.pause_btn.configure(state='normal' if running else 'disabled', text="⏸ 暂停")
        self.skip_btn.configure(state='normal' if running else 'disabled')
        self.abort_btn.configure(state='normal' if running else 'disabled')
        self.status_var.set("运行中..." if running else "就绪")

    def _run_worker(self):
        """worker 线程：逐行执行，所有 UI 更新经 _ui_q"""
        n = len(self.sequence_data)
        for idx in range(n):
            if self._stop.is_set():
                break
            self._pause.wait()
            if self._stop.is_set():
                break
            self._ui_q.put(("status", (idx, "运行中", "")))
            self._log(f"--- 第 {idx + 1}/{n} 行 ---")
            try:
                outcome = self._run_one_row(idx)
            except Exception as e:
                self._ui_q.put(("status", (idx, "失败", str(e))))
                self._log(f"[行{idx + 1}] 异常: {e}")
                outcome = "fail"
            if outcome == "abort":
                self._log("用户中止序列")
                break
        self._ui_q.put(("done", None))

    def _run_one_row(self, idx):
        """单行流水线（worker 线程内）。返回 'ok'/'skip'/'abort'，失败自行 put status。"""
        row = self.sequence_data[idx]
        log = lambda m: self._log(f"[行{idx + 1}] {m}")

        # 1. 定位谱图 PDF + 解析 sampleCode
        pdf_path, sample_code = self._resolve_spectrum_pdf(row, log)
        if not sample_code:
            self._ui_q.put(("status", (idx, "失败", "无法确定样品编号")))
            log("失败: 谱图目录无匹配PDF且未填样品编号")
            return "fail"
        self._ui_q.put(("rowdata", (idx, {"sample_code": sample_code})))

        # 2. 查样品
        log(f"查询样品 {sample_code} ...")
        projects = self.api.query_samples_by_conditions(
            sample_code=sample_code, exact_match=True, log_func=log)
        if not projects:
            self._ui_q.put(("status", (idx, "失败", "未查到样品")))
            log("失败: LIMS 未查到该样品(可能未登记或超30天)")
            return "fail"

        projects = self._filter_projects_by_method(projects, row, log)
        if not projects:
            self._ui_q.put(("status", (idx, "失败", "无匹配方法项目")))
            log("失败: 该样品下没有匹配方法文件的项目")
            return "fail"

        first = projects[0]
        sample_id = first.get("sampleId")
        project_ids = [str(p["projectId"]) for p in projects if p.get("projectId")]
        project_names = [p.get("projectName", "") for p in projects]
        sp_ids_str = ",".join(project_ids)
        method_name = projects[0].get("standardNo") or ""

        # 3. 配置中段（复刻 submit_single_method_group 纯 API 部分）
        log("清暂存 / 取实验配置 ...")
        self.api.clear_experiment_cache(sp_ids_str, log)
        initial = self.api.get_experiment_config(sp_ids_str, method_name, "", sample_id, log)
        if not initial:
            self._ui_q.put(("status", (idx, "失败", "取实验配置失败")))
            log("失败: get_experiment_config 返回空")
            return "fail"
        oc = initial.get("ocMethodSettings", {}) or {}
        method_id = oc.get("methodId")
        actual_method_name = method_name
        actual_method_id = method_id
        if method_id and str(method_id) in self.api.sub_method_map:
            sub = self.api.sub_method_map[str(method_id)]
            actual_method_id = sub.get("sub_method_id")
            self.api.update_method(sp_ids_str, actual_method_id, project_names, log)
            std = self.api.get_method_standard_no_by_id(actual_method_id, log)
            if std:
                actual_method_name = std
        log("取全量配置 ...")
        all_cfg = self.api.get_all_configs(sp_ids_str, method_name, "", sample_id, log, method_id, project_names)
        if not all_cfg:
            self._ui_q.put(("status", (idx, "失败", "取全量配置失败")))
            return "fail"
        experiment_config = all_cfg.get("experiment")
        equipment_config = all_cfg.get("equipment")
        dynamic_columns = all_cfg.get("dynamic_columns") or []
        actual_method_name = all_cfg.get("actual_method_name", actual_method_name)
        actual_method_id = all_cfg.get("actual_method_id", actual_method_id)

        # 4. 实验编号
        log("生成实验编号 ...")
        experiment_code = self.api.generate_experiment_code(method_name=actual_method_name, log_func=log)
        if not experiment_code:
            self._ui_q.put(("status", (idx, "失败", "生成实验编号失败")))
            return "fail"

        configure_order = (row.get("configure_order") or "").strip()

        # 5. 标准溶液预检
        if configure_order:
            log(f"校验标准溶液 {configure_order} ...")
            cid, emsg = self.api.get_solution_configure_id(configure_order, log)
            if not cid:
                self._ui_q.put(("status", (idx, "失败", f"标准溶液:{emsg}")))
                log(f"失败: 标准溶液校验未过 - {emsg}")
                return "fail"

        # 6. 暂停等人工确认
        prompt = {
            "idx": idx, "sample_code": sample_code, "method": actual_method_name,
            "experiment_code": experiment_code, "configure_order": configure_order,
            "dynamic_columns": dynamic_columns,
        }
        result = self._prompt_confirm(prompt)
        if result == "skip":
            self._ui_q.put(("status", (idx, "跳过", "")))
            log("已跳过本行")
            return "skip"
        if result == "abort":
            return "abort"
        values = result

        # 7. 构造 HeadlessHost + build + submit
        log("构建并提交 ...")
        host = self._build_headless_host(
            values, experiment_config, equipment_config, dynamic_columns,
            actual_method_name, actual_method_id, configure_order)
        experiment_data = build_grouped_experiment_data(host, projects, experiment_code, actual_method_name)
        ok = self.api.submit_experiment_data(experiment_data, actual_method_name, log)
        if not ok:
            self._ui_q.put(("status", (idx, "失败", "实验数据提交失败")))
            log("失败: submit_experiment_data 返回 False")
            return "fail"
        if configure_order:
            exp_codes = {pid: experiment_code for pid in project_ids}
            sok, serr = self.api.submit_solution_with_experiment(project_ids, configure_order, exp_codes, log)
            if not sok:
                self._ui_q.put(("status", (idx, "失败", f"标准溶液关联:{serr}")))
                log(f"失败: 标准溶液关联失败 - {serr}")
                return "fail"

        # 8. 成功
        self._ui_q.put(("status", (idx, "成功", "")))
        self._ui_q.put(("rowdata", (idx, {"experiment_code": experiment_code})))
        log(f"成功，实验编号 {experiment_code}")
        return "ok"

    def _resolve_spectrum_pdf(self, row, log):
        """定位本行谱图PDF(目录内按文件名前缀=sampleCode匹配)。返回 (pdf_path, sample_code)"""
        sp = row.get("spectrum_path") or ""
        sc = (row.get("sample_code") or "").strip()
        if os.path.isfile(sp):
            return sp, (sc or os.path.basename(sp).split("-", 1)[0])
        pdfs = sorted(glob.glob(os.path.join(sp, "*.pdf"))) if os.path.isdir(sp) else []
        if not pdfs:
            return None, sc
        if sc:
            for p in pdfs:
                if os.path.basename(p).startswith(sc):
                    return p, sc
            log(f"警告: 目录内无文件名以 {sc} 开头的PDF")
            return None, sc
        if len(pdfs) == 1:
            return pdfs[0], os.path.basename(pdfs[0]).split("-", 1)[0]
        log(f"目录有 {len(pdfs)} 个PDF且未填样品编号，无法确定(请一行一PDF或先填sample_code)")
        return None, ""

    def _filter_projects_by_method(self, projects, row, log):
        """若方法yaml有query_rules.method(ID)，按其标准号过滤projects；否则原样返回"""
        method_file = row.get("method_file") or ""
        method_id = None
        if method_file and os.path.isfile(method_file):
            try:
                with open(method_file, "r", encoding="utf-8") as f:
                    y = yaml.safe_load(f) or {}
                qr = y.get("query_rules") or []
                if qr and qr[0].get("method"):
                    method_id = str(qr[0]["method"])
            except Exception as e:
                log(f"警告: 读取方法文件失败 {e}")
        if not method_id:
            return projects
        try:
            std = self.api.get_method_standard_no_by_id(method_id, log)
        except Exception:
            std = None
        if std:
            filtered = [p for p in projects if p.get("standardNo") == std]
            if filtered:
                log(f"按方法 {std} 过滤出 {len(filtered)} 个项目")
                return filtered
            log(f"警告: 无项目匹配标准号 {std}，使用全部 {len(projects)} 个项目")
        return projects

    def _build_headless_host(self, values, experiment_config, equipment_config,
                             dynamic_columns, actual_method_name, actual_method_id, configure_order):
        """构造无头 host 供 build_grouped_experiment_data 使用"""
        data_fields = {code: _Box(val) for code, val in values.items()}
        today = datetime.now().strftime("%Y-%m-%d")
        return types.SimpleNamespace(
            login_system=self.login_system,
            api=self.api,
            log=lambda m: self._log(f"[行] {m}"),
            experiment_config=experiment_config or {},
            equipment_config=equipment_config or {},
            dynamic_columns=dynamic_columns or [],
            data_fields=data_fields,
            remark_text=_Box(""),
            solution_type_var=_Box(configure_order or ""),
            actual_method_name=actual_method_name,
            actual_method_id=actual_method_id,
            temperature_var=_Box("22"),
            humidity_var=_Box("55"),
            start_date_var=_Box(today),
            end_date_var=_Box(today),
        )

    def _prompt_confirm(self, prompt):
        """worker 线程：请求主线程弹确认面板并阻塞。返回 dict/'skip'/'abort'"""
        self._confirm_done.clear()
        self._confirm_result = None
        self._ui_q.put(("prompt", prompt))
        self._confirm_done.wait()
        return self._confirm_result

    def _drain_ui_queue(self):
        """主线程：轮询 worker 消息并更新 UI"""
        try:
            while True:
                kind, payload = self._ui_q.get_nowait()
                if kind == "log":
                    self._append_log(payload)
                elif kind == "status":
                    idx, st, err = payload
                    if idx < len(self.row_widgets):
                        self.row_widgets[idx].set_status(st, err)
                    if idx < len(self.sequence_data):
                        self.sequence_data[idx]["status"] = st
                        self.sequence_data[idx]["error_msg"] = err
                elif kind == "rowdata":
                    idx, extras = payload
                    if idx < len(self.sequence_data):
                        self.sequence_data[idx].update(extras)
                elif kind == "prompt":
                    self._build_confirm_dialog(payload)
                elif kind == "done":
                    self._on_run_done()
        except queue.Empty:
            pass
        self.root.after(100, self._drain_ui_queue)

    def _on_run_done(self):
        self._set_running(False)
        self._confirm_done.set()
        self._append_log("==== 序列运行结束 ====")
        ok = sum(1 for r in self.sequence_data if r.get("status") == "成功")
        fail = sum(1 for r in self.sequence_data if r.get("status") == "失败")
        self.status_var.set(f"完成: 成功 {ok} / 失败 {fail}")
        messagebox.showinfo("完成", f"序列运行结束\n成功 {ok} / 失败 {fail}")

    def _append_log(self, msg):
        self.log_text.configure(state='normal')
        self.log_text.insert('end', msg + "\n")
        self.log_text.see('end')
        self.log_text.configure(state='disabled')

    def _log(self, msg):
        self._ui_q.put(("log", msg))

    def _build_confirm_dialog(self, prompt):
        """主线程：构建确认面板（非 grab，不挡表格/日志）"""
        if self._confirm_win is not None and self._confirm_win.winfo_exists():
            self._confirm_win.destroy()
        win = tk.Toplevel(self.root)
        win.title(f"确认提交 - 第 {prompt['idx'] + 1} 行")
        win.geometry("540x480")
        win.transient(self.root)

        info = tk.Frame(win, padx=10, pady=8)
        info.pack(fill='x')
        tk.Label(info, text=f"样品: {prompt['sample_code']}", anchor='w').pack(fill='x')
        tk.Label(info, text=f"方法: {prompt['method']}", anchor='w').pack(fill='x')
        tk.Label(info, text=f"实验编号: {prompt['experiment_code']}", anchor='w', fg="#2563eb").pack(fill='x')
        co = prompt['configure_order']
        tk.Label(info, text=f"标准溶液: {co or '(无)'}", anchor='w',
                 fg=("#16a34a" if co else "#999")).pack(fill='x')

        fields_frame = tk.LabelFrame(win, text="结果数值（请逐项核对填写）", padx=8, pady=6)
        fields_frame.pack(fill='both', expand=True, padx=8, pady=6)
        entries = {}
        cols = prompt['dynamic_columns'] or []
        if not cols:
            tk.Label(fields_frame, text="(无动态字段配置)").pack()
        for col in cols:
            code = col.get("columeCode", "")
            name = col.get("columeName", code)
            r = tk.Frame(fields_frame)
            r.pack(fill='x', pady=2)
            tk.Label(r, text=name, width=22, anchor='w').pack(side='left')
            e = tk.Entry(r, width=24)
            e.insert(0, str(col.get("defaultVal", "")))
            e.pack(side='left', padx=4)
            entries[code] = e

        btns = tk.Frame(win, pady=8)
        btns.pack(fill='x')

        def on_submit():
            self._confirm_result = {code: e.get().strip() for code, e in entries.items()}
            self._confirm_done.set()
            win.destroy()

        def on_skip():
            self._confirm_result = "skip"
            self._confirm_done.set()
            win.destroy()

        def on_abort():
            self._confirm_result = "abort"
            self._confirm_done.set()
            win.destroy()

        tk.Button(btns, text="确认提交", command=on_submit, width=12, bg="#dbeafe").pack(side='left', padx=6)
        tk.Button(btns, text="跳过本行", command=on_skip, width=10).pack(side='left', padx=6)
        tk.Button(btns, text="中止序列", command=on_abort, width=10, bg="#fee2e2").pack(side='left', padx=6)
        win.protocol("WM_DELETE_WINDOW", on_skip)
        self._confirm_win = win

    def toggle_pause(self):
        if not self._running:
            return
        if self._pause.is_set():
            self._pause.clear()
            self.pause_btn.configure(text="▶ 继续")
            self._log("已暂停(行边界生效)")
            self.status_var.set("已暂停")
        else:
            self._pause.set()
            self.pause_btn.configure(text="⏸ 暂停")
            self._log("已继续")
            self.status_var.set("运行中...")

    def skip_current(self):
        if not self._running:
            return
        if not self._confirm_done.is_set():
            self._confirm_result = "skip"
            self._confirm_done.set()
        self._log("请求跳过当前行(确认面板或下一行生效)")

    def abort_run(self):
        if not self._running:
            return
        if messagebox.askyesno("确认", "确定中止序列运行吗？"):
            self._stop.set()
            self._pause.set()
            if not self._confirm_done.is_set():
                self._confirm_result = "abort"
                self._confirm_done.set()
            self._log("请求中止...")

    def _show_login_dialog(self):
        """主线程：弹出登录对话框(选用户名+验证码)。返回是否登录成功"""
        users = list(self.login_system.users.keys())
        if not users:
            messagebox.showerror("错误", "没有可用用户(请先在主程序登录保存用户)")
            return False
        win = tk.Toplevel(self.root)
        win.title("登录 LIMS")
        win.geometry("320x320")
        win.transient(self.root)
        win.grab_set()

        tk.Label(win, text="用户:").pack(pady=(12, 0))
        user_var = tk.StringVar(value=users[0])
        ttk.Combobox(win, textvariable=user_var, values=users, state='readonly', width=22).pack()

        cap_label = tk.Label(win, text="正在获取验证码...", bg="#fff")
        cap_label.pack(pady=8)
        cap_photo = {'img': None}
        cap_entry_var = tk.StringVar()

        def fetch_captcha():
            img = self.login_system.get_captcha_image()
            if img is None:
                cap_label.configure(image="", text="获取验证码失败，点刷新重试")
                return
            photo = ImageTk.PhotoImage(img)
            cap_photo['img'] = photo
            cap_label.configure(image=photo, text="")

        def on_refresh():
            cap_label.configure(image="", text="正在获取验证码...")
            win.after(10, fetch_captcha)

        def on_login():
            ok, msg = self.login_system.login_with_user(user_var.get(), cap_entry_var.get().strip())
            if ok:
                self.logged_in = True
                self.login_status_var.set(f"🟢 已登录: {user_var.get()}")
                self.login_lbl.configure(fg="#16a34a")
                self.login_btn.configure(text="切换用户")
                self.status_var.set(f"已登录: {msg}")
                win.destroy()
            else:
                messagebox.showwarning("登录失败", msg, parent=win)
                cap_entry_var.set("")
                on_refresh()

        tk.Label(win, text="验证码:").pack()
        cap_row = tk.Frame(win)
        cap_row.pack(pady=4)
        tk.Entry(cap_row, textvariable=cap_entry_var, width=14).pack(side='left')
        tk.Button(cap_row, text="刷新", command=on_refresh, width=6).pack(side='left', padx=4)
        tk.Button(win, text="登录", command=on_login, width=12, bg="#dbeafe").pack(pady=10)

        win.after(50, fetch_captcha)
        win.wait_window(win)
        return self.logged_in


if __name__ == "__main__":
    root = tk.Tk()
    app = SequenceMaster(root)
    root.mainloop()