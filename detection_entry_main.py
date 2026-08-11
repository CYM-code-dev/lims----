# detection_entry_main.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as ttkb  # 档1: 现代主题(sandstone-light)，ttk 控件自动套用
import json
import time
from datetime import datetime, timedelta
import webbrowser
import argparse
import os
import re
import sys

# 导入登录系统和API模块
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from login import MultiUserLoginSystem
from detection_entry_api import DetectionAPI, build_grouped_experiment_data
import threading
import paths
from report_parser import parse_pdf_report, parse_pdf_report_multi, _dilution_factor, filter_samples_by_code
from alias_evaluator import evaluate_alias

# ponytail: 与 SequenceMaster._strip_parallel_suffix 同逻辑；两模块刻意解耦不互导，故复制。
_PARALLEL_SUFFIX_RE = re.compile(r'^([A-Za-z]+\d{8})\d{3}([A-Za-z]*)$')


def _strip_parallel_suffix(code):
    """去掉称样编号的平行小号(001)：TN26070620001→TN26070620，用作谱图前缀兜底匹配
    (ICP 文件名常只用 8 位流水 TN26070620.pdf，无小号)。不符结构原样返回。"""
    m = _PARALLEL_SUFFIX_RE.match((code or "").strip())
    return (m.group(1) + m.group(2)) if m else (code or "")


def _make_large_checkbutton_style(scale=1.5):
    """放大 ttk.Checkbutton 勾选框：复用 ttkbootstrap 自带位图按 scale 重建 indicator element，
    风格与 sandstone-light 主题保持一致。ponytail: indicator 是固定尺寸位图，只能重建放大；
    依赖 ttkbootstrap 内部 API，失败则回退默认勾选框。"""
    try:
        from ttkbootstrap.style.layout import image_element, layout, El
        from ttkbootstrap.style.assets import RecolorRenderer
        from PIL import ImageTk
        s = ttkb.Style()
        c = s.colors
        size = int(36 * scale)
        def _img(name, ink):
            return ImageTk.PhotoImage(RecolorRenderer.render(name, (size, size), '#ffffff', ink, None, None))
        chk, ind = _img('checkbox_checked', c.primary), _img('checkbox_indeterminate', c.primary)
        unk = _img('checkbox_unchecked', c.fg)
        dchk, dind = _img('checkbox_checked', c.border), _img('checkbox_indeterminate', c.border)
        dunk = _img('checkbox_unchecked', c.border)
        sn = 'Large.TCheckbutton'
        image_element(s, sn + '.indicator', default=chk,
                      states={'disabled selected': dchk, 'disabled alternate': dind, 'disabled': dunk,
                              'alternate': ind, '!selected': unk}, border=0, padding=0, sticky=tk.W)
        layout(s, sn, El('Checkbutton.padding', sticky=tk.NSEW, children=[
            El(sn + '.indicator', side=tk.LEFT, sticky=''),
            El('Checkbutton.focus', side=tk.LEFT, sticky='', children=[El('Checkbutton.label', sticky=tk.NSEW)])]))
        s.configure(sn, background=c.bg, foreground=c.fg, focuscolor=c.fg)
        s._large_cb_imgs = (chk, unk, ind, dchk, dunk, dind)  # 持有 PhotoImage 引用防 GC
        return sn
    except Exception:
        return None


def _bind_tooltip(widget, text):
    """鼠标悬停显示完整文本（用于被列宽截断的单元格）"""
    tip = {'win': None}
    def show(_):
        if not text or tip['win']:
            return
        tw = tk.Toplevel(widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{widget.winfo_rootx() + 18}+{widget.winfo_rooty() + widget.winfo_height() + 4}")
        ttk.Label(tw, text=text, background="#ffffe0", relief='solid', borderwidth=1,
                  font=("微软雅黑", 9)).pack(ipadx=4, ipady=2)
        tip['win'] = tw
    def hide(_):
        if tip['win']:
            tip['win'].destroy()
            tip['win'] = None
    widget.bind('<Enter>', show)
    widget.bind('<Leave>', hide)


class DetectionEntrySystem:
    def __init__(self, root, result_checkin_ids=None, sample_id=None, sample_project_ids=None, url_params=None):
        self.root = root
        self.root.title("检测数据录入系统")
        self.root.geometry("1500x1200")
        self.solution_type_var = None
        self.solution_status_var = None
        # 使用登录系统和API模块
        self.login_system = MultiUserLoginSystem()
        self.api = DetectionAPI(self.login_system)

        # 动态参数
        self.current_result_checkin_id = result_checkin_ids
        self.current_sample_id = sample_id
        self.current_sample_project_ids = sample_project_ids
        self.url_params = url_params or {}

        # 存储数据
        self.samples_info = {}  # 存储所有样品信息
        self.detected_projects = []
        self.project_vars = []
        self.project_id_to_var = {}
        self.dynamic_columns = []
        self.filtered_projects = []  # 存储过滤后的项目
        self.experiment_config = {}  # 存储实验配置
        self.experiment_codes = {}
        # 初始化界面相关的变量
        self.sample_code_var = None
        self.project_name_var = None
        self.method_name_var = None
        self.retest_var = None
        self.search_results_frame = None

        # 初始化手动输入相关的变量
        self.detection_no_var = None
        self.sample_small_no_var = None
        self.is_retest_var = None
        self.old_sample_project_id_var = None
        self.old_sample_project_id_entry = None

        # 初始化调试相关的变量
        self.debug_text = None

        # 新增：标准溶液相关变量
        self.solution_type_var = None
        self.solution_status_var = None

        # 谱图上传相关变量
        self.spectrum_rows = []  # 每行 {project, file_path_var, status_var, record}
        self.spectrum_uploaded = []  # 已上传文件记录 {fileId, fileName, url}
        self.spectrum_list_frame = None
        self.spectrum_status_var = None

        # 创建界面
        self.setup_ui()

        # 检查登录状态
        self.check_login_status()

    def check_login_status(self):
        """检查登录状态"""
        if self.login_system.load_session() and self.login_system.verify_session():
            disp = self.login_system.users.get(self.login_system.current_user, {}).get('display_name', self.login_system.current_user)
            self.update_status(f"已登录: {disp}", "green")
        else:
            self.update_status("未登录", "red")

    def setup_ui(self):
        """设置用户界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 状态栏
        self.setup_status_bar(main_frame)

        # 创建标签页
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        # 标签页标题字体（样品信息/环境与设备/检测数据/谱图上传）
        ttk.Style().configure("TNotebook.Tab", font=("微软雅黑", 10), padding=(28, 8))
        # 选中的标签页文字变蓝，便于区分当前页
        ttk.Style().map("TNotebook.Tab", foreground=[("selected", "#1565C0")])

        # 档3: 放大复选框勾选框（ttkbootstrap 默认 indicator 偏小）
        self.large_cb_style = _make_large_checkbutton_style(0.9)

        # 样品信息标签页
        sample_frame = ttk.Frame(notebook, padding="10")
        notebook.add(sample_frame, text="样品信息")

        # 环境条件标签页
        env_frame = ttk.Frame(notebook, padding="10")
        notebook.add(env_frame, text="环境与设备")

        # 检测数据标签页
        data_frame = ttk.Frame(notebook, padding="10")
        notebook.add(data_frame, text="检测数据")

        # 谱图上传标签页
        spectrum_frame = ttk.Frame(notebook, padding="10")
        notebook.add(spectrum_frame, text="谱图上传")

        # 设置各个标签页
        self.setup_sample_tab(sample_frame)
        self.setup_env_tab(env_frame)
        self.setup_data_tab(data_frame)
        self.setup_spectrum_tab(spectrum_frame)

    def setup_status_bar(self, parent):
        """设置状态栏"""
        status_frame = ttk.Frame(parent)
        status_frame.pack(fill=tk.X)

        # 状态信息
        self.status_var = tk.StringVar(value="状态: 未登录")
        status_label = ttk.Label(status_frame, textvariable=self.status_var)
        status_label.pack(side=tk.LEFT)

        # 显示当前样品信息
        self.sample_info_var = tk.StringVar(value="")
        sample_info_label = ttk.Label(status_frame, textvariable=self.sample_info_var)
        sample_info_label.pack(side=tk.LEFT, padx=(20, 0))

        # 显示检测项目和检测方法
        self.project_method_var = tk.StringVar(value="")
        project_method_label = ttk.Label(status_frame, textvariable=self.project_method_var)
        project_method_label.pack(side=tk.LEFT, padx=(20, 0))

    def setup_sample_tab(self, parent):
        """设置样品信息标签页 - 4个输入框+模糊查询"""
        # 查询条件区域
        query_frame = ttk.LabelFrame(parent, text="查询条件", padding="10")
        query_frame.pack(fill=tk.X, pady=(0, 10))

        # 第一行 - 样品编号和项目
        row1_frame = ttk.Frame(query_frame)
        row1_frame.pack(fill=tk.X, pady=5)

        ttk.Label(row1_frame, text="样品编号:").pack(side=tk.LEFT)
        self.sample_code_var = tk.StringVar(value="")  # 默认空；原测试期默认 TR26070030 会污染"只按检测方法"查询(被当成 keyword 与方法取交集)
        sample_code_entry = ttk.Entry(row1_frame, textvariable=self.sample_code_var, width=20)
        sample_code_entry.pack(side=tk.LEFT, padx=(5, 20))

        ttk.Label(row1_frame, text="项目:").pack(side=tk.LEFT)
        self.project_name_var = tk.StringVar()
        project_name_entry = ttk.Entry(row1_frame, textvariable=self.project_name_var, width=20)
        project_name_entry.pack(side=tk.LEFT, padx=(5, 20))

        # 检测方法、注销复测（与样品编号/项目同一行）
        ttk.Label(row1_frame, text="检测方法:").pack(side=tk.LEFT)
        self.method_name_var = tk.StringVar()
        method_name_entry = ttk.Entry(row1_frame, textvariable=self.method_name_var, width=20)
        method_name_entry.pack(side=tk.LEFT, padx=(5, 20))

        # 注销复测改为复选框
        self.retest_var = tk.BooleanVar(value=False)
        retest_checkbox = ttk.Checkbutton(row1_frame, text="注销复测", variable=self.retest_var, style=self.large_cb_style)
        retest_checkbox.pack(side=tk.LEFT, padx=(5, 0))

        # 查询按钮区域
        button_frame = ttk.Frame(query_frame)
        button_frame.pack(fill=tk.X, pady=10)

        ttkb.Button(button_frame, text="查询", command=self.query_samples, bootstyle="primary").pack(side=tk.LEFT, padx=(0, 10))
        ttkb.Button(button_frame, text="清空条件", command=self.clear_query_conditions, bootstyle="secondary").pack(side=tk.LEFT, padx=(0, 10))

        # 查询结果统计
        self.result_count_var = tk.StringVar(value="查询结果: 0 条")
        ttk.Label(button_frame, textvariable=self.result_count_var).pack(side=tk.LEFT, padx=(20, 0))

        # 查询结果区域
        results_frame = ttk.LabelFrame(parent, text="查询结果", padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 创建滚动框架用于结果显示
        canvas = tk.Canvas(results_frame, bg=ttkb.Style().colors.bg, highlightthickness=0)
        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=canvas.yview)
        self.search_results_frame = ttk.Frame(canvas)

        self.search_results_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.search_results_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 项目选择按钮
        selection_frame = ttk.Frame(parent)
        selection_frame.pack(fill=tk.X, pady=5)

        ttkb.Button(selection_frame, text="全选", command=self.select_all_projects, bootstyle="secondary").pack(side=tk.LEFT, padx=(0, 10))
        ttkb.Button(selection_frame, text="取消全选", command=self.deselect_all_projects, bootstyle="secondary").pack(side=tk.LEFT,
                                                                                              padx=(0, 10))

        # 新增：获取动态配置按钮
        ttkb.Button(selection_frame, text="获取动态配置", command=self.get_dynamic_config, bootstyle="primary").pack(side=tk.LEFT,
                                                                                               padx=(0, 10))

        # 已选择项目统计
        self.selected_count_var = tk.StringVar(value="已选择: 0 个项目")
        ttk.Label(selection_frame, textvariable=self.selected_count_var).pack(side=tk.LEFT, padx=(20, 0))

        # 配置获取结果信息栏（替代弹窗）
        self.config_status_var = tk.StringVar(value="")
        ttk.Label(selection_frame, textvariable=self.config_status_var, foreground="blue").pack(side=tk.LEFT, padx=(20, 0))

    def setup_data_tab(self, parent):
        """设置检测数据标签页 - 添加标准溶液提交功能"""
        # 第一行 - 当前方法ID + 切换方法ID
        method_row = ttk.Frame(parent)
        method_row.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(method_row, text="当前方法ID:").pack(side=tk.LEFT)
        self.current_method_id_var = tk.StringVar(value="未获取")
        ttk.Label(method_row, textvariable=self.current_method_id_var, foreground="blue").pack(side=tk.LEFT, padx=(5, 20))
        ttk.Label(method_row, text="切换方法ID:").pack(side=tk.LEFT)
        self.switch_method_id_var = tk.StringVar()
        ttk.Entry(method_row, textvariable=self.switch_method_id_var, width=15).pack(side=tk.LEFT, padx=(5, 5))
        ttkb.Button(method_row, text="切换方法", command=self.switch_method_id, bootstyle="secondary").pack(side=tk.LEFT)
        # 可切换方法ID提示（获取动态配置时填充，提示当前方法下可输入哪些ID）
        self.switchable_methods_var = tk.StringVar(value="")
        ttk.Label(method_row, textvariable=self.switchable_methods_var, foreground="gray").pack(side=tk.LEFT, padx=(15, 0))

        # 检测数据区域
        data_frame = ttk.LabelFrame(parent, text="检测数据", padding="10")
        data_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 创建主框架来包含动态字段、备注字段和标准溶液
        main_data_frame = ttk.Frame(data_frame)
        main_data_frame.pack(fill=tk.BOTH, expand=True)

        # 动态字段框架 - 添加滚动条
        dynamic_container = ttk.Frame(main_data_frame)
        dynamic_container.pack(fill=tk.BOTH, expand=True)

        # 添加滚动条
        canvas = tk.Canvas(dynamic_container, height=200)  # 限制高度
        scrollbar = ttk.Scrollbar(dynamic_container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.dynamic_fields_frame = scrollable_frame

        # 初始化默认字段（在获取动态配置前显示默认字段）
        self.setup_default_data_fields()

        # 按钮区域
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=10)

        ttkb.Button(button_frame, text="结果计算", command=self.calc_results, bootstyle="info").pack(side=tk.LEFT, padx=(0, 10))
        ttkb.Button(button_frame, text="提交数据", command=self.submit_data, bootstyle="primary").pack(side=tk.LEFT, padx=(0, 10))
        ttkb.Button(button_frame, text="清空数据", command=self.clear_data, bootstyle="secondary").pack(side=tk.LEFT, padx=(0, 10))
        ttkb.Button(button_frame, text="数据采集", command=self.acquire_data, bootstyle="success").pack(side=tk.LEFT, padx=(0, 10))
        ttkb.Button(button_frame, text="采集设置", command=self.edit_acquisition_settings, bootstyle="secondary").pack(side=tk.LEFT, padx=(0, 10))

        # 添加分组信息显示
        self.group_info_var = tk.StringVar(value="")
        group_info_label = ttk.Label(button_frame, textvariable=self.group_info_var, foreground="blue")
        group_info_label.pack(side=tk.LEFT, padx=(20, 0))

        # 实验编号显示（保留在检测数据标签页）
        self.experiment_code_var = tk.StringVar(value="实验编号: 未生成")
        ttk.Label(button_frame, textvariable=self.experiment_code_var, foreground="blue").pack(side=tk.LEFT, padx=(20, 0))

    def setup_solution_section(self, parent):
        """设置标准溶液提交区域 - 显示实验编号信息"""
        # 标准溶液框架
        solution_frame = ttk.LabelFrame(parent, text="标准溶液配置", padding="10")
        solution_frame.pack(fill=tk.X, pady=(10, 0))

        # 配置序号输入
        solution_row1 = ttk.Frame(solution_frame)
        solution_row1.pack(fill=tk.X, pady=5)

        ttk.Label(solution_row1, text="配置序号:").pack(side=tk.LEFT)

        # 确保正确初始化
        if not hasattr(self, 'solution_type_var') or self.solution_type_var is None:
            self.solution_type_var = tk.StringVar(value="")  # 初始化为空字符串

        self.solution_entry = ttk.Entry(solution_row1, textvariable=self.solution_type_var, width=30)
        self.solution_entry.pack(side=tk.LEFT, padx=(5, 10))

        # 添加示例提示
        ttk.Label(solution_row1, text="例如: D-6688", foreground="gray").pack(side=tk.LEFT, padx=(5, 0))

        # 溶液提交状态
        if not hasattr(self, 'solution_status_var') or self.solution_status_var is None:
            self.solution_status_var = tk.StringVar(value="")
        ttk.Label(solution_frame, textvariable=self.solution_status_var, foreground="blue").pack(anchor=tk.W, pady=5)

    def update_experiment_code_display(self):
        """更新实验编号显示"""
        if hasattr(self, 'experiment_codes') and self.experiment_codes:
            # 显示第一个实验编号作为示例
            first_code = list(self.experiment_codes.values())[0]
            count = len(self.experiment_codes)
            display_text = f"实验编号: {first_code} (共 {count} 个)"
            self.experiment_code_var.set(display_text)

            # 删除成功日志
            # 只更新状态栏，不输出日志
            self.update_status(f"实验编号已生成: {first_code}", "green")
        else:
            self.experiment_code_var.set("实验编号: 未生成")

    def setup_spectrum_tab(self, parent):
        """设置谱图上传标签页"""
        # 信息栏（第一行，顶部）
        info_frame = ttk.Frame(parent)
        info_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))
        self.spectrum_status_var = tk.StringVar(value="未刷新")
        ttk.Label(info_frame, textvariable=self.spectrum_status_var, foreground="blue").pack(side=tk.LEFT)

        toolbar = ttk.Frame(parent)
        toolbar.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0))

        ttkb.Button(toolbar, text="刷新选中项目", command=self.refresh_spectrum_list, bootstyle="secondary").pack(side=tk.LEFT, padx=(0, 5))
        ttkb.Button(toolbar, text="选择谱图(按样品编号匹配)", command=self.pick_spectrum_match, bootstyle="secondary").pack(side=tk.LEFT, padx=(0, 5))
        ttkb.Button(toolbar, text="同一谱图应用到全部", command=self.pick_spectrum_apply_all, bootstyle="secondary").pack(side=tk.LEFT, padx=(0, 5))
        ttkb.Button(toolbar, text="开始上传", command=self.upload_all_spectra, bootstyle="primary").pack(side=tk.LEFT, padx=(0, 5))

        # 滚动列表
        list_outer = ttk.Frame(parent)
        list_outer.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(list_outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_outer, orient="vertical", command=canvas.yview)
        self.spectrum_list_frame = ttk.Frame(canvas)

        self.spectrum_list_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self.spectrum_list_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self._render_spectrum_rows()

    def _render_spectrum_rows(self):
        """根据 self.spectrum_rows 重建行控件"""
        for widget in self.spectrum_list_frame.winfo_children():
            widget.destroy()

        if not self.spectrum_rows:
            ttk.Label(self.spectrum_list_frame,
                      text="请先在「样品信息」页选中项目，再点「刷新选中项目」",
                      foreground="gray").pack(pady=20)
            return

        # 表头
        header = ttk.Frame(self.spectrum_list_frame)
        header.pack(fill=tk.X, pady=(0, 3))
        ttk.Label(header, text="样品编号", width=16).pack(side=tk.LEFT)
        ttk.Label(header, text="项目", width=16).pack(side=tk.LEFT)
        ttk.Label(header, text="样品名称", width=14).pack(side=tk.LEFT)
        ttk.Label(header, text="谱图文件", width=26).pack(side=tk.LEFT)
        ttk.Label(header, text="状态", width=12).pack(side=tk.LEFT)
        ttk.Separator(self.spectrum_list_frame, orient='horizontal').pack(fill=tk.X, pady=3)

        for row in self.spectrum_rows:
            proj = row["project"]
            line = ttk.Frame(self.spectrum_list_frame)
            line.pack(fill=tk.X, pady=1)
            ttk.Label(line, text=proj.get('sampleCode', ''), width=16).pack(side=tk.LEFT)
            ttk.Label(line, text=proj.get('projectName', ''), width=16).pack(side=tk.LEFT)
            ttk.Label(line, text=proj.get('sampleName', ''), width=14).pack(side=tk.LEFT)
            ttk.Label(line, textvariable=row["file_disp"], width=26, foreground="blue").pack(side=tk.LEFT)
            ttk.Label(line, textvariable=row["status_var"], width=12).pack(side=tk.LEFT)
            ttkb.Button(line, text="重选", width=4, bootstyle="secondary",
                       command=lambda r=row: self._repick_row(r)).pack(side=tk.LEFT, padx=(4, 0))

    def refresh_spectrum_list(self):
        """按样品信息页勾选的项目重建谱图行列表"""
        if not self.login_system.current_user:
            messagebox.showerror("错误", "请先登录系统")
            return

        selected_indices = [i for i, var in enumerate(self.project_vars) if var.get()]
        if not selected_indices:
            messagebox.showerror("错误", "请先在「样品信息」页至少选择一个项目")
            return

        # 按 projectId 保留已选文件/已上传记录
        old_by_pid = {r["project"].get("projectId"): r for r in self.spectrum_rows}
        self.spectrum_rows = []
        for index in selected_indices:
            if index < len(self.filtered_projects):
                project = self.filtered_projects[index]
                prev = old_by_pid.get(project.get("projectId"))
                if prev:
                    row = {
                        "project": project,
                        "file_path": prev["file_path"],
                        "file_disp": prev["file_disp"],
                        "status_var": prev["status_var"],
                        "record": prev["record"],
                    }
                else:
                    row = {
                        "project": project,
                        "file_path": None,
                        "file_disp": tk.StringVar(value="未选择"),
                        "status_var": tk.StringVar(value="待上传"),
                        "record": None,
                    }
                self.spectrum_rows.append(row)

        self._render_spectrum_rows()
        self._sync_uploaded()
        self._update_spectrum_status()

    def pick_spectrum_match(self):
        """选择多个文件，按样品编号匹配到对应行"""
        if not self._ensure_spectrum_rows():
            return
        paths = filedialog.askopenfilenames(
            title="选择谱图文件（按样品编号匹配）",
            filetypes=[("PDF文件", "*.pdf"), ("图片", "*.jpg *.jpeg *.png"), ("所有文件", "*.*")]
        )
        if not paths:
            return

        matched_any = False
        unmatched = []
        for path in paths:
            stem = os.path.splitext(os.path.basename(path))[0]
            hits = [r for r in self.spectrum_rows
                    if stem and stem in r["project"].get('sampleCode', '')]
            if hits:
                for r in hits:
                    r["file_path"] = path
                    r["file_disp"].set(os.path.basename(path))
                    if r["record"] is None:
                        r["status_var"].set("待上传")
                matched_any = True
            else:
                unmatched.append(os.path.basename(path))

        self._render_spectrum_rows()
        self._update_spectrum_status()
        parts = []
        if matched_any:
            parts.append("已匹配")
        if unmatched:
            parts.append("未匹配: " + ", ".join(unmatched))
        self.spectrum_status_var.set("；".join(parts) if parts else "无匹配")

    def pick_spectrum_apply_all(self):
        """选一个文件，赋给所有行"""
        if not self._ensure_spectrum_rows():
            return
        path = filedialog.askopenfilename(
            title="选择一个谱图文件（应用到全部）",
            filetypes=[("PDF文件", "*.pdf"), ("图片", "*.jpg *.jpeg *.png"), ("所有文件", "*.*")]
        )
        if not path:
            return
        for r in self.spectrum_rows:
            r["file_path"] = path
            r["file_disp"].set(os.path.basename(path))
            if r["record"] is None:
                r["status_var"].set("待上传")
        self._render_spectrum_rows()
        self._update_spectrum_status()

    def _repick_row(self, row):
        path = filedialog.askopenfilename(
            title="为该行选择谱图文件",
            filetypes=[("PDF文件", "*.pdf"), ("图片", "*.jpg *.jpeg *.png"), ("所有文件", "*.*")]
        )
        if path:
            row["file_path"] = path
            row["file_disp"].set(os.path.basename(path))
            row["record"] = None
            row["status_var"].set("待上传")
            self._sync_uploaded()
            self._update_spectrum_status()

    def upload_all_spectra(self):
        """上传所有已选文件且未上传的行"""
        if not self.login_system.current_user:
            messagebox.showerror("错误", "请先登录系统")
            return
        if not self.spectrum_rows:
            messagebox.showerror("错误", "请先点「刷新选中项目」")
            return

        todo = [r for r in self.spectrum_rows if r["file_path"] and r["record"] is None]
        if not todo:
            messagebox.showinfo("提示", "没有待上传的谱图（已选文件均已上传，或未选择文件）")
            return

        # 按文件路径去重：同一文件只上传一次，结果共享给所有引用该路径的行（否则 N 个项目=N 份重复上传）
        unique_paths = list(dict.fromkeys(r["file_path"] for r in todo))

        self.update_status(f"正在上传 {len(unique_paths)} 个谱图文件...")
        self.root.update_idletasks()

        success_count = 0
        fail_count = 0
        for path in unique_paths:
            rows = [r for r in todo if r["file_path"] == path]
            for r in rows:
                r["status_var"].set("上传中...")
            self.root.update_idletasks()
            project_id = rows[0]["project"].get("projectId")
            sample_id = rows[0]["project"].get("sampleId")
            self.log(f"上传 {os.path.basename(path)}: sampleId={sample_id!r}, projectId(resultCheckInId)={project_id}（共享 {len(rows)} 行）")
            success, result = self.api.upload_spectrum_file(path, str(project_id) if project_id else "", self.log)
            if success:
                record = {
                    "fileId": result.get("id"),
                    "fileName": result.get("orgName") or result.get("name"),
                    "url": result.get("url"),
                }
                for r in rows:
                    r["record"] = record
                    r["status_var"].set(f"成功 {record['fileName']}")
                success_count += 1
            else:
                for r in rows:
                    r["status_var"].set(f"失败: {result}")
                fail_count += 1

        self._sync_uploaded()
        self._update_spectrum_status()
        self.update_status(f"谱图上传完成: 成功 {success_count}, 失败 {fail_count}",
                           "green" if fail_count == 0 else "orange")
        if fail_count:
            messagebox.showwarning("部分失败", f"成功 {success_count} 个，失败 {fail_count} 个，请查看状态列")

    def _ensure_spectrum_rows(self):
        if not self.spectrum_rows:
            messagebox.showerror("错误", "请先点「刷新选中项目」加载项目列表")
            return False
        return True

    def _sync_uploaded(self):
        """spectrum_uploaded 与当前行的已上传记录保持一致（按 fileId 去重：多项目共享同一文件只算一份）"""
        seen = set()
        out = []
        for r in self.spectrum_rows:
            rec = r.get("record")
            if rec and rec.get("fileId") not in seen:
                seen.add(rec.get("fileId"))
                out.append(rec)
        self.spectrum_uploaded = out

    def _update_spectrum_status(self):
        total = len(self.spectrum_rows)
        assigned = sum(1 for r in self.spectrum_rows if r["file_path"])
        uploaded = sum(1 for r in self.spectrum_rows if r["record"])
        self.spectrum_status_var.set(f"共 {total} 项 | 已选文件 {assigned} | 已上传 {uploaded}")

    def calc_results(self):
        """结果计算：按当前输入算出报告值并输出到日志（不提交）。
        复用 build_grouped_experiment_data 的取数与 calcTheValue 计算，仅省略提交动作。"""
        if not self.login_system.current_user:
            messagebox.showerror("错误", "请先登录系统")
            return
        selected_indices = [i for i, var in enumerate(self.project_vars) if var.get()]
        if not selected_indices:
            messagebox.showerror("错误", "请至少选择一个检测项目")
            return
        if not self.dynamic_columns or not getattr(self, 'experiment_config', None):
            messagebox.showerror("错误", "请先获取动态配置")
            return
        selected_projects = [self.filtered_projects[i] for i in selected_indices if i < len(self.filtered_projects)]
        if not selected_projects:
            messagebox.showerror("错误", "未找到选中的项目")
            return

        method_name = getattr(self, 'actual_method_name', None) or selected_projects[0].get('standardNo', '')
        experiment_code = self.api.get_experiment_code_for_method(method_name, self.log) or "preview"
        self.log(f"==== 结果计算 开始 (方法: {method_name}) ====")
        try:
            experiment_data = self.build_grouped_experiment_data(selected_projects, experiment_code, method_name)
        except Exception as e:
            self.log(f"结果计算失败: {e}")
            return
        try:
            records = json.loads(experiment_data.get('ocAnalysisRecordSaveList') or '[]')
        except Exception:
            records = []
        for r in records:
            rv, cv = r.get('reportValue'), r.get('calculatedValue')
            self.log(f"{r.get('sampleCode', '')} #{r.get('serialNumber', '')}  "
                     f"报告值={rv if rv is not None else '空'}  计算值={cv if cv is not None else '空'}"
                     + (f"  [{r.get('other')}]" if r.get('other') else ''))
        self.log(f"==== 结果计算 完成 ({len(records)} 条) ====")

    # ---------------- 谱图数据采集（本地解析）----------------

    _ACQ_SETTINGS_FILE = 'acquisition_settings.json'
    # 方法/项目名 -> 谱图文件名类别关键词（一样多报告时挑选用）
    _CATEGORY_HINTS = (
        ('PAE', ('PAE', '邻苯')), ('PAHS', ('PAHS', '多环', 'PAH')),
        ('PCN', ('PCN', '氯萘')), ('BXW', ('BXW', '苯系物')),
        ('氯苯', ('氯苯',)), ('SCCP', ('SCCP',)), ('SVHC', ('SVHC',)),
        ('AZO', ('AZO', '偶氮')), ('OT', ('OT',)),
        ('ICP', ('ICP', '元素', '重金属')),
    )

    def _acq_settings_path(self):
        return os.path.join(paths.data_dir(), self._ACQ_SETTINGS_FILE)

    def _load_acquisition_settings(self):
        try:
            with open(self._acq_settings_path(), 'r', encoding='utf-8') as f:
                return json.load(f) or {}
        except Exception:
            return {}

    def _save_acquisition_settings(self, settings):
        try:
            with open(self._acq_settings_path(), 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False)
        except Exception as e:
            self.log(f"保存采集设置失败: {e}")

    def edit_acquisition_settings(self):
        """选择谱图 PDF 所在的本地/网络文件夹并持久化。"""
        cur = self._load_acquisition_settings().get('pdf_folder', '')
        folder = filedialog.askdirectory(title="选择谱图 PDF 文件夹", initialdir=cur or None)
        if folder:
            self._save_acquisition_settings({'pdf_folder': folder})
            self.log(f"采集设置已保存：{folder}")

    def _find_pdf_for_sample(self, folder, sample_code):
        """返回目录中文件名以 sampleCode 开头(大小写不敏感)的 PDF 列表。
        直匹配空时兜底：按去小号前缀(如 TN26070620001→TN26070620)匹配——
        ICP 等设备文件名只用 8 位流水、无小号(001)。"""
        sc = (sample_code or '').lower()
        out = []
        try:
            fns = [fn for fn in os.listdir(folder) if fn.lower().endswith('.pdf')]
        except Exception as e:
            self.log(f"读取谱图文件夹失败: {e}")
            return out
        out = [os.path.join(folder, fn) for fn in fns if fn.lower().startswith(sc)]
        if out:
            return out
        stripped = _strip_parallel_suffix(sample_code).lower()
        if stripped and stripped != sc:
            out = [os.path.join(folder, fn) for fn in fns if fn.lower().startswith(stripped)]
            if out:
                self.log(f"样品 {sample_code} 按去小号前缀 {stripped} 匹配到谱图（文件名无小号）")
        return out

    def _pick_pdf(self, pdfs, method_hint):
        """一样多报告时按方法/项目名类别关键词挑选；仍歧义取首个并告警。"""
        if len(pdfs) == 1:
            return pdfs[0]
        text = (method_hint or '').upper()
        hint = ''
        for cat, keys in self._CATEGORY_HINTS:
            if any(k.upper() in text for k in keys):
                hint = cat
                break
        if hint:
            matched = [p for p in pdfs if hint.lower() in os.path.basename(p).lower()]
            if matched:
                return matched[0]
        self.log(f"存在多个谱图，取首个（候选：{[os.path.basename(p) for p in pdfs]}）")
        return pdfs[0]

    def _find_result_column(self, pdf_headers=()):
        """挑数据采集结果列：
        1) equipRelativeTitle 精确匹配 PDF 报告表头(原系统机制，非硬编码：浓度/校准浓度等都能对上)；
           表头命中多列时优先结果列(表头/列名含 浓度/计算值/报告值...)，避免化合物/峰面积等非结果列抢中；
        2) 兜底：列名关键词(计算值/报告值/浓度...)。多命中取 columeOrder 最小。"""
        KW = ('计算值', '报告值', '测定值', '浓度', '结果值', '含量')

        def _ok(c):
            return not self._is_merge_column(c) and not self._is_component_column(c)

        def _desc(c):
            return "(%s|%s|表头=%s|o%s%s%s)" % (
                c.get('columeCode'), c.get('columeName'), c.get('equipRelativeTitle'), c.get('columeOrder'),
                '|合并' if self._is_merge_column(c) else '', '|组分' if self._is_component_column(c) else '')

        self.log("全部动态列: %s" % [_desc(c) for c in self.dynamic_columns])

        pick, path = None, ""
        if pdf_headers:
            cands = [c for c in self.dynamic_columns if _ok(c)
                     and (c.get('equipRelativeTitle') or '').strip() in pdf_headers]
            if cands:
                # 表头命中多列：优先结果列(表头/列名含结果关键词)，否则取 columeOrder 最小
                kw_hits = [c for c in cands
                           if any(k in ((c.get('equipRelativeTitle') or '') + (c.get('columeName') or ''))
                                  for k in KW)]
                pool = kw_hits or cands
                pick = min(pool, key=lambda c: c.get('columeOrder', 9999))
                path = "表头命中%s %s" % ("(结果列优先)" if kw_hits else "", [_desc(c) for c in cands])
        if pick is None:
            cands = [c for c in self.dynamic_columns if _ok(c)
                     and any(k in (c.get('columeName') or '') for k in KW)]
            if cands:
                pick = min(cands, key=lambda c: c.get('columeOrder', 9999))
                path = "列名关键词兜底 %s" % [_desc(c) for c in cands]
        if pick is not None:
            self.log("选列诊断: PDF表头=%s | %s → 选 %s" % (list(pdf_headers), path, _desc(pick)))
            return pick.get('columeCode')
        self.log("选列诊断: PDF表头=%s | 未找到结果列(非合并/非组分列中无表头命中、列名也无关键词)"
                 % list(pdf_headers))
        return None

    def acquire_data(self):
        """数据采集：本地解析谱图 PDF -> 按项目别名求值 -> 回填检测数据结果列。"""
        if not self.login_system.current_user:
            messagebox.showerror("错误", "请先登录系统")
            return
        selected_indices = [i for i, var in enumerate(self.project_vars) if var.get()]
        if not selected_indices:
            messagebox.showerror("错误", "请至少选择一个检测项目")
            return
        if not self.dynamic_columns or not getattr(self, 'experiment_config', None):
            messagebox.showerror("错误", "请先获取动态配置")
            return
        selected_projects = [self.filtered_projects[i] for i in selected_indices if i < len(self.filtered_projects)]
        if not selected_projects:
            messagebox.showerror("错误", "未找到选中的项目")
            return

        sample_code = selected_projects[0].get('sampleCode', '')
        folder = self._load_acquisition_settings().get('pdf_folder', '')
        if not folder:
            messagebox.showwarning("数据采集", "请先点「采集设置」指定谱图 PDF 文件夹")
            return
        pdfs = self._find_pdf_for_sample(folder, sample_code)
        if not pdfs:
            messagebox.showwarning("数据采集", f"在文件夹未找到样品 {sample_code} 的谱图 PDF")
            return
        method_hint = getattr(self, 'actual_method_name', '') or selected_projects[0].get('projectName', '')
        # 分类正常/稀释(-NNX)报告：超线性组分改用稀释报告读数
        normal_pdfs = [p for p in pdfs if _dilution_factor(p) == 1.0]
        dil_pdfs = [p for p in pdfs if _dilution_factor(p) != 1.0]
        pdf_path = self._pick_pdf(normal_pdfs or pdfs, method_hint)
        dil_path = self._pick_pdf(dil_pdfs, method_hint) if dil_pdfs else None

        # 收集各选中项目的 detectionProjectId（取别名用）
        project_specs = [{'projectId': p.get('projectId'), 'projectName': p.get('projectName', ''),
                          'detectionProjectId': p.get('detectionProjectId')} for p in selected_projects]
        project_specs = [s for s in project_specs if s['detectionProjectId']]
        if not project_specs:
            messagebox.showwarning("数据采集", "未取到检测项目定义ID(detectionProjectId)，无法取项目别名")
            return

        self.update_status(f"正在解析谱图 {os.path.basename(pdf_path)} ...", "blue")
        threading.Thread(target=self._acquire_worker, args=(pdf_path, project_specs, dil_path, sample_code), daemon=True).start()

    def _acquire_worker(self, pdf_path, project_specs, dil_path=None, sample_code=''):
        """子线程：解析正常 PDF（+ 可选稀释报告），逐样品/逐项目取别名求值，回主线程填表。
        ICP 一 PDF 多样品(A/B 平行样)：按样品序号 slot 分别求值，回填时落对应平行槽。"""
        try:
            samples, headers = parse_pdf_report_multi(pdf_path)
            if not samples:
                self.root.after(0, lambda: messagebox.showwarning("数据采集", "未从谱图解析到化合物，请检查报告格式"))
                return
            samples = filter_samples_by_code(samples, sample_code)  # ICP 多报验批共一份 PDF：只取当前样品段
            diluted = None
            dilution_factor = 1.0
            if dil_path:
                diluted, _ = parse_pdf_report(dil_path)
                dilution_factor = _dilution_factor(dil_path)
            plan = []
            for spec in project_specs:
                alias, detail = self.api.get_project_alias(spec['detectionProjectId'], self.log)
                if not alias:
                    self.log(f"项目 {spec['projectName']} 别名取不到({detail})，跳过")
                    continue
                item = {'projectName': spec['projectName'], 'projectId': spec['projectId'], 'results': []}
                for slot, (_sid, compounds) in enumerate(samples):
                    for r in evaluate_alias(alias, compounds, diluted_compounds=diluted):
                        r['slot'] = slot
                        item['results'].append(r)
                plan.append(item)
            self.root.after(0, self._apply_acquired_results, plan, headers, dilution_factor, len(samples))
        except Exception as e:
            err = str(e)
            self.root.after(0, lambda: messagebox.showerror("数据采集失败", err))

    def _apply_acquired_results(self, plan, pdf_headers=(), dilution_factor=1.0, n_samples=1):
        """把求值结果填入结果列对应行/组分块。超线性(稀释)组分同时填稀释列。
        多样品(ICP A/B)：result 携带 slot，落 pos = b*N + slot 平行槽。"""
        res_col = self._find_result_column(pdf_headers)
        if not res_col:
            messagebox.showerror("数据采集", "未找到结果列(计算值/报告值/浓度)，请检查动态列")
            return
        target = self.data_fields.get(res_col)
        if not isinstance(target, list):
            messagebox.showerror("数据采集", "结果列不可填（非普通列）")
            return
        N = getattr(self, 'test_run_count', 1) or 1
        comp_col = next((c.get('columeCode') for c in self.dynamic_columns if self._is_component_column(c)), None)
        comp_vars = self.data_fields.get(comp_col) if comp_col else None
        dil_col = next((c.get('columeCode') for c in self.dynamic_columns
                        if (c.get('equipRelativeTitle') or '').strip() == '稀释'), None)
        records = (getattr(self, 'experiment_config', {}) or {}).get('ocAnalysisRecordList') or []
        # 无组分列时按 projectId 定位行；有组分列时按组分名匹配
        pid_to_block = {}
        if comp_col is None:
            for _r in records:
                _pid = _r.get('projectId')
                if _pid is not None and _pid not in pid_to_block:
                    pid_to_block[_pid] = len(pid_to_block)

        filled, total, slot_skip = 0, 0, 0
        for item in plan:
            for r in item['results']:
                total += 1
                slot = r.get('slot', 0)
                if slot >= N:  # 平行槽不足：N=1 时放不下第2个样品
                    slot_skip += 1
                    continue
                comp_name = r.get('lims_component')
                b = None
                if comp_name and isinstance(comp_vars, list):
                    for idx, v in enumerate(comp_vars):
                        if v.get() == comp_name:
                            b = idx // N
                            break
                if b is None:  # 无组分列：按 projectId 落到对应项目行
                    b = pid_to_block.get(item.get('projectId'), 0)
                pos = b * N + slot
                if pos < len(target):
                    target[pos].set(r['value'])
                    if (dilution_factor != 1.0 and dil_col is not None
                            and r.get('raw', {}).get('diluted')):
                        dil_target = self.data_fields.get(dil_col)
                        if isinstance(dil_target, list) and pos < len(dil_target):
                            dil_target[pos].set(f"{dilution_factor:g}")
                    filled += 1
                else:
                    self.log(f"无法定位填入位置：{item['projectName']} = {r['value']}")
        if filled == 0:
            messagebox.showwarning("数据采集", "未填入任何值，请检查项目别名配置与结果列")
            self.update_status("数据采集未填入值", "orange")
            return
        self.update_status(f"数据采集完成：填入 {filled}/{total} 个值（结果列 {res_col}）", "green")
        if n_samples > 1 and N > 1:
            self.log(f"ICP 平行样：A/B 已分别填入平行槽 0/1（N={N}）")
        if slot_skip:
            self.log(f"提示：{slot_skip} 个值因平行槽不足(N={N}<{n_samples})被跳过")

    def submit_data(self):
        """提交数据 - 对选中的复选框进行操作，并自动提交标准溶液"""
        if not self.login_system.current_user:
            messagebox.showerror("错误", "请先登录系统")
            return

        # 检查是否选择了项目
        selected_indices = [i for i, var in enumerate(self.project_vars) if var.get()]
        if not selected_indices:
            messagebox.showerror("错误", "请至少选择一个检测项目")
            return

        # 检查是否已获取动态配置
        if not self.dynamic_columns:
            messagebox.showerror("错误", "请先获取动态配置")
            return

        # 验证温湿度数据
        try:
            temperature = float(self.temperature_var.get())
            humidity = float(self.humidity_var.get())
        except ValueError:
            messagebox.showerror("错误", "温湿度必须为数字")
            return

        # 验证质量数据
        try:
            # 动态获取质量字段
            mass_field = self.find_mass_field()
            if mass_field:
                # 非合并列可能有多次测试的质量值，逐一校验
                field = self.data_fields[mass_field]
                vals = field if isinstance(field, list) else [field]
                for v in vals:
                    mass_value = float(v.get() or "0")
                    if mass_value <= 0:
                        messagebox.showerror("错误", "质量必须大于0")
                        return
        except ValueError:
            messagebox.showerror("错误", "质量必须为数字")
            return

        # 验证日期格式
        try:
            datetime.strptime(self.start_date_var.get(), "%Y-%m-%d")
            datetime.strptime(self.end_date_var.get(), "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("错误", "日期格式不正确，请使用 YYYY-MM-DD 格式")
            return

        # 获取选中的项目并按检测方法分组
        selected_projects = []
        for index in selected_indices:
            if index < len(self.filtered_projects):
                project = self.filtered_projects[index]
                selected_projects.append(project)

        # 按检测方法分组
        method_groups = {}
        for project in selected_projects:
            method_name = project.get('standardNo', '未知方法')
            if method_name not in method_groups:
                method_groups[method_name] = []
            method_groups[method_name].append(project)

        # 检查是否填写了配置序号
        configure_order = None
        if hasattr(self, 'solution_type_var') and self.solution_type_var is not None:
            raw_value = self.solution_type_var.get()
            if raw_value and raw_value.strip() and raw_value.strip().lower() != 'none':
                configure_order = raw_value.strip()

        # 为每个方法组提交数据
        success_groups = []
        failed_groups = []
        solution_success_groups = []
        solution_failed_groups = []
        solution_unaudited_groups = []

        for method_name, projects in method_groups.items():
            # 提交实验数据和标准溶液
            experiment_success, solution_status = self.submit_single_method_group(method_name, projects,
                                                                                  configure_order)

            if experiment_success:
                success_groups.append(method_name)

                # 检查标准溶液提交状态
                if configure_order and solution_status is not None:
                    if solution_status is True:
                        solution_success_groups.append(method_name)
                    elif solution_status == "未审核":
                        solution_unaudited_groups.append(method_name)
                    else:
                        solution_failed_groups.append(method_name)
            else:
                failed_groups.append(method_name)

        # 结果写入状态栏（不弹窗）
        parts = [f"{g}（{len(method_groups[g])}项）" for g in success_groups]
        status_text = "提交完成：" + "、".join(parts) if parts else "实验数据提交失败"
        if failed_groups:
            status_text += f"；实验失败 {len(failed_groups)} 组"
        if configure_order:
            if solution_success_groups:
                status_text += f"；标准溶液成功 {len(solution_success_groups)}"
            if solution_unaudited_groups:
                status_text += f"；标准溶液未审核 {len(solution_unaudited_groups)}"
            if solution_failed_groups:
                status_text += f"；标准溶液失败 {len(solution_failed_groups)}"
        else:
            status_text += "；标准溶液未填配置序号"
        self.update_status(status_text)

        # 仅全部成功（无实验失败、无标准溶液失败/未审核）时清空数据
        if not failed_groups and not solution_failed_groups and not solution_unaudited_groups:
            self.clear_data()

    def get_dynamic_config(self):
        """获取选中项目的动态配置"""
        if not self.login_system.current_user:
            messagebox.showerror("错误", "请先登录系统")
            return

        # 检查是否选择了项目
        selected_indices = [i for i, var in enumerate(self.project_vars) if var.get()]
        if not selected_indices:
            messagebox.showerror("错误", "请至少选择一个检测项目")
            return

        # 获取选中的项目并按检测方法分组
        selected_projects = []
        for index in selected_indices:
            if index < len(self.filtered_projects):
                project = self.filtered_projects[index]
                selected_projects.append(project)

        # 按检测方法分组
        method_groups = {}
        for project in selected_projects:
            method_name = project.get('standardNo', '未知方法')
            if method_name not in method_groups:
                method_groups[method_name] = []
            method_groups[method_name].append(project)

        # 为每个方法组获取配置
        success_groups = []
        failed_groups = []

        for method_name, projects in method_groups.items():
            try:
                success = self.get_single_method_group_config(method_name, projects)
            except Exception as e:
                import traceback
                messagebox.showerror("获取配置异常", f"方法组 {method_name}:\n{e}\n\n{traceback.format_exc()}")
                success = False
            if success:
                success_groups.append(method_name)
            else:
                failed_groups.append(method_name)

        # 配置获取成功后，自动刷新「谱图上传」标签页的选中项目，并自动查询设备
        if success_groups:
            self.refresh_spectrum_list()
            self.query_equipment_choices()  # 配置成功后自动拉取设备可选列表（切换方法走同一刷新路径，同样触发）

        # 显示结果
        result_message = f"配置获取完成！\n\n"
        result_message += f"总项目数: {len(selected_projects)}\n\n"

        if success_groups:
            result_message += f"成功组 ({len(success_groups)}):\n"
            for group in success_groups:
                project_count = len(method_groups[group])
                result_message += f"  ✓ {group} ({project_count}个项目)\n"

        if failed_groups:
            result_message += f"\n失败组 ({len(failed_groups)}):\n"
            for group in failed_groups:
                project_count = len(method_groups[group])
                result_message += f"  ✗ {group} ({project_count}个项目)\n"

        # 信息栏显示配置结果（替代成功弹窗）
        if success_groups:
            parts = [f"{g}（{len(method_groups[g])}项）" for g in success_groups]
            status_text = "配置完成：" + "、".join(parts)
            if failed_groups:
                status_text += f"；失败 {len(failed_groups)} 组"
            self.config_status_var.set(status_text)
        else:
            self.config_status_var.set(f"配置失败：{len(failed_groups)} 组")

        # 仅在存在失败组时弹窗提示
        if failed_groups:
            messagebox.showwarning("部分成功", result_message)

    def get_single_method_group_config(self, method_name, projects):
        """获取单个检测方法组的配置 - 简化版本"""
        try:
            project_ids = [str(p.get('projectId')) for p in projects if p.get('projectId')]
            project_names = [p.get('projectName', '') for p in projects]

            if not project_ids:
                return False

            # 获取所有配置信息（包含动态列处理）
            sample_project_ids_str = ",".join(project_ids)
            first_project = projects[0]
            sample_id = first_project.get('sampleId')

            # 先取实验配置，解析方法ID并立即显示——确保即使后续 get_all_configs 的子步骤
            # （动态列/设备等）抛异常，"当前方法ID" 也能显示出来
            experiment_config = self.api.get_experiment_config(
                sample_project_ids_str, method_name, "", sample_id, self.log, None
            )
            method_id = (experiment_config or {}).get('ocMethodSettings', {}).get('methodId')
            actual_method_id = method_id
            description = None
            if method_id and str(method_id) in self.api.sub_method_map:
                sub_method_info = self.api.sub_method_map[str(method_id)]
                actual_method_id = sub_method_info['sub_method_id']
                description = sub_method_info.get('description')
            # 查询当前方法下可切换的方法ID（整合进获取动态配置流程）
            switchable = []
            if method_id and project_names:
                switchable = self.api.get_switchable_methods(project_names[0], method_id, self.log)
            self.switchable_methods = switchable  # 供 switch_method_id 切换后同步本地 standardNo

            # 当前方法ID 附带名称（从可切换列表取名称，空则显示「主方法」）
            if actual_method_id is not None:
                cur = next((m for m in switchable if str(m.get('decideProjectMethodId')) == str(actual_method_id)), None)
                cur_name = (cur.get('decideProjectMethodName') if cur else "") or "主方法"
                self.current_method_id_var.set(f"{actual_method_id}（{cur_name}）")
            else:
                self.current_method_id_var.set("未获取")

            # 排除当前已显示的方法ID（避免与「当前方法ID」重复）
            others = [m for m in switchable if str(m.get('decideProjectMethodId')) != str(actual_method_id)]
            hint = "、".join(
                f"{m.get('decideProjectMethodId')}（{m.get('decideProjectMethodName') or '主方法'}）"
                for m in others
            )
            self.switchable_methods_var.set(f"可切换ID: {hint}" if hint else "")

            # 使用统一的方法获取配置（传入已解析的 method_id，对齐 submit_data 流程）
            all_configs = self.api.get_all_configs(
                sample_project_ids_str,
                method_name,
                "",  # result_checkin_ids
                sample_id,
                self.log,
                method_id,  # 已解析的方法ID
                project_names  # 项目名称列表
            )

            if not all_configs:
                error_msg = "无法获取配置信息，请检查网络连接或系统状态"
                raise Exception(error_msg)

            # 提取各个配置
            self.experiment_config = all_configs.get('experiment', {})
            self.equipment_config = all_configs.get('equipment', {})
            self._refresh_equipment_display()
            self.units_config = all_configs.get('units', [])
            self.round_methods_config = all_configs.get('round_methods', {})
            self.calc_methods_config = all_configs.get('calc_methods', {})
            self.dynamic_columns = all_configs.get('dynamic_columns', [])

            # 测试次数 N：单项目在 ocAnalysisRecordList 中的记录数（默认 1，平行样为 2）
            self.test_run_count = self._compute_test_run_count(self.experiment_config)

            # 实际方法名称（可能经过子方法切换）
            actual_method_name = all_configs.get('actual_method_name', method_name)
            display_method_name = f"{actual_method_name} {description}" if description else actual_method_name

            # 保存实际的方法名称和方法ID用于实验数据
            self.actual_method_name = actual_method_name
            self.actual_method_id = actual_method_id

            # 更新UI动态字段
            self.setup_dynamic_data_fields(self.dynamic_columns)

            # 更新状态栏显示当前检测方法
            self.project_method_var.set(f"检测方法: {display_method_name}")

            return True

        except Exception as e:
            import traceback
            raise Exception(f"获取检测方法组 {method_name} 配置失败: {str(e)}")

    def build_grouped_experiment_data(self, projects, experiment_code, method_name):
        """构建分组实验数据 - 委托给 detection_entry_api 共享实现（阶段0抽取，零行为变化）"""
        return build_grouped_experiment_data(self, projects, experiment_code, method_name)

    def on_project_selected(self):
        """当项目被选择时自动更新状态"""
        self.update_selected_count()

        # 检查是否只选择了一个检测方法
        selected_indices = [i for i, var in enumerate(self.project_vars) if var.get()]
        if len(selected_indices) == 1:
            # 单个项目被选中时，自动获取配置
            index = selected_indices[0]
            if index < len(self.filtered_projects):
                project = self.filtered_projects[index]
                method_name = project.get('standardNo', '未知方法')
        else:
            # 多个项目被选中时，检查是否是同一个检测方法
            selected_projects = []
            for index in selected_indices:
                if index < len(self.filtered_projects):
                    project = self.filtered_projects[index]
                    selected_projects.append(project)

            # 检查所有选中项目的检测方法是否相同
            methods = set(p.get('standardNo', '') for p in selected_projects)
            if len(methods) == 1:
                method_name = list(methods)[0]

    def setup_fixed_remark_field(self):
        """设置固定的备注字段"""
        # 创建备注字段框架
        remark_frame = ttk.LabelFrame(self.dynamic_fields_frame, text="备注信息", padding="10")
        remark_frame.pack(fill=tk.X, pady=10)

        # 备注标签和输入框
        ttk.Label(remark_frame, text="备注:").pack(anchor=tk.W, pady=5)

        # 创建滚动文本框用于备注
        text_frame = ttk.Frame(remark_frame)
        text_frame.pack(fill=tk.X, pady=5)

        # 创建文本框和滚动条
        self.remark_text = tk.Text(text_frame, height=4, width=80, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.remark_text.yview)
        self.remark_text.configure(yscrollcommand=scrollbar.set)

        self.remark_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def get_select_values(self, col_id, col_code):
        """获取选择框的可选值"""

        if col_code == 'dynamic1692':  # 体积V（ml）
            return ['10', '25', '50', '100']
        elif col_code == 'dynamic1693':  # 样品空白C0（μg/mL）
            return ['<0.10', '<0.05', '<0.01', '0.00']
        elif col_code == 'dynamic1694':  # 样品浓度C1（μg/mL）
            return ['<0.10', '<0.05', '<0.01', '0.00', '0.10', '0.50', '1.00']
        elif col_code == 'dynamic1695':  # 稀释因子F
            return ['1', '2', '5', '10']
        else:
            # 默认返回空列表，使用文本框
            return []

    def setup_env_tab(self, parent):
        """设置环境与设备标签页（环境条件 + 当前方法ID下的设备信息）"""
        # 环境条件区域
        env_frame = ttk.LabelFrame(parent, text="环境条件", padding="10")
        env_frame.pack(fill=tk.X, pady=(0, 10))

        # 温湿度和时间
        env_row1 = ttk.Frame(env_frame)
        env_row1.pack(fill=tk.X, pady=10)

        ttk.Label(env_row1, text="温度(℃):").pack(side=tk.LEFT)
        self.temperature_var = tk.StringVar(value="21.0")
        ttk.Entry(env_row1, textvariable=self.temperature_var, width=10).pack(side=tk.LEFT, padx=(5, 20))

        ttk.Label(env_row1, text="湿度(%RH):").pack(side=tk.LEFT)
        self.humidity_var = tk.StringVar(value="51.0")
        ttk.Entry(env_row1, textvariable=self.humidity_var, width=10).pack(side=tk.LEFT, padx=(5, 20))

        # 日期信息
        env_row2 = ttk.Frame(env_frame)
        env_row2.pack(fill=tk.X, pady=10)

        ttk.Label(env_row2, text="开始日期:").pack(side=tk.LEFT)
        self.start_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        ttk.Entry(env_row2, textvariable=self.start_date_var, width=12).pack(side=tk.LEFT, padx=5)

        ttk.Label(env_row2, text="结束日期:").pack(side=tk.LEFT)
        self.end_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        ttk.Entry(env_row2, textvariable=self.end_date_var, width=12).pack(side=tk.LEFT, padx=5)

        # 设置日期按钮
        env_row3 = ttk.Frame(env_frame)
        env_row3.pack(fill=tk.X, pady=10)

        ttkb.Button(env_row3, text="设置当前日期", command=self.set_current_date, bootstyle="secondary").pack(side=tk.LEFT)

        # 设备信息区域：复选框多选（主检支持多台），点"查询设备"按当前样品项目拉取可选列表
        eq_frame = ttk.LabelFrame(parent, text="设备信息", padding="10")
        eq_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        eq_top = ttk.Frame(eq_frame)
        eq_top.pack(fill=tk.X, pady=(0, 5))
        ttkb.Button(eq_top, text="查询设备", command=self.query_equipment_choices, bootstyle="secondary").pack(side=tk.LEFT)
        self.eq_query_status_var = tk.StringVar(value="未查询")
        ttk.Label(eq_top, textvariable=self.eq_query_status_var, foreground="blue").pack(side=tk.LEFT, padx=(10, 0))

        eq_cols = ttk.Frame(eq_frame)
        eq_cols.pack(fill=tk.BOTH, expand=True)

        main_col = ttk.LabelFrame(eq_cols, text="主检设备", padding="5")
        main_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        self.main_eq_inner = self._build_checklist(main_col)
        self.main_eq_vars = []  # [(BooleanVar, item), ...]

        weighing_col = ttk.LabelFrame(eq_cols, text="称样设备", padding="5")
        weighing_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.weighing_eq_inner = self._build_checklist(weighing_col)
        self.weighing_eq_vars = []

        # 标准溶液提交区域（置于环境与设备标签页最下方）
        self.setup_solution_section(parent)

    def _build_checklist(self, parent):
        """在 parent 内构建可滚动复选框容器，返回内部 Frame（用于放置 Checkbutton）"""
        canvas = tk.Canvas(parent, height=140, bg=ttkb.Style().colors.bg, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        return inner

    def _populate_checklist(self, inner, var_list, items, defaults):
        """用 items 重建复选框；label 命中 defaults 任一字符串的默认勾选"""
        for w in inner.winfo_children():
            w.destroy()
        var_list.clear()
        for it in items:
            label = it.get('label', '')
            if not label:
                continue
            checked = any(d and d in label for d in defaults)
            var = tk.BooleanVar(value=checked)
            ttk.Checkbutton(inner, text=label, variable=var, style=self.large_cb_style).pack(anchor=tk.W, fill=tk.X)
            var_list.append((var, it))

    @staticmethod
    def _split_names(joined):
        """分号连接的设备串拆成集合，用于默认勾选匹配"""
        return {p.strip() for p in (joined or '').split(';') if p.strip()}

    def _refresh_equipment_display(self):
        """配置加载后用方法默认设备初始化复选框（点"查询设备"后会替换为完整可选列表，默认项仍勾选）"""
        if not hasattr(self, 'main_eq_inner'):
            return
        eq = getattr(self, 'equipment_config', None) or {}
        main_defaults = self._split_names(eq.get('mainEquipment', ''))
        weighing_defaults = self._split_names(eq.get('weighingEquipment', ''))
        self._populate_checklist(self.main_eq_inner, self.main_eq_vars,
                                 [{'label': n, 'id': ''} for n in main_defaults], main_defaults)
        self._populate_checklist(self.weighing_eq_inner, self.weighing_eq_vars,
                                 [{'label': n, 'id': ''} for n in weighing_defaults], weighing_defaults)
        self.eq_query_status_var.set("已显示默认设备，点查询设备可拉取全部")

    def _first_selected_project_id(self):
        """返回第一个选中检测项目的 projectId，无则 None"""
        project_vars = getattr(self, 'project_vars', [])
        filtered = getattr(self, 'filtered_projects', [])
        for i, var in enumerate(project_vars):
            if var.get() and i < len(filtered):
                pid = filtered[i].get('projectId')
                if pid:
                    return pid
        return None

    def query_equipment_choices(self):
        """查询当前选中项目的主检/称样设备可选列表，重建复选框（对齐网页端 ocMultipleChoicePage / ocChoicePage）"""
        if not self.login_system.current_user:
            messagebox.showerror("错误", "请先登录系统")
            return
        sample_project_id = self._first_selected_project_id()
        if not sample_project_id:
            messagebox.showerror("错误", "请先选择检测项目")
            return
        self.log(f"查询设备可选列表: SampleProjectId={sample_project_id}")
        main_list = self.api.get_main_equipment_choices(sample_project_id, self.log)
        weighing_list = self.api.get_weighing_equipment_choices(sample_project_id, self.log)
        eq = getattr(self, 'equipment_config', None) or {}
        self._populate_checklist(self.main_eq_inner, self.main_eq_vars,
                                 main_list, self._split_names(eq.get('mainEquipment', '')))
        self._populate_checklist(self.weighing_eq_inner, self.weighing_eq_vars,
                                 weighing_list, self._split_names(eq.get('weighingEquipment', '')))
        self.eq_query_status_var.set(f"主检 {len(main_list)} 项, 称样 {len(weighing_list)} 项")
        self.log(f"设备可选列表已更新: 主检 {len(main_list)} 项, 称样 {len(weighing_list)} 项")

    def setup_default_data_fields(self):
        """设置默认的数据字段（在获取动态配置前使用）- 包括固定备注字段"""
        # 清空现有字段
        for widget in self.dynamic_fields_frame.winfo_children():
            widget.destroy()

        self.data_fields = {}

        # 显示提示信息
        info_label = ttk.Label(self.dynamic_fields_frame, text="请先选择检测项目并获取动态配置", foreground="gray")
        info_label.pack(pady=20)

        # 即使没有动态配置，也显示固定备注字段
        self.setup_fixed_remark_field()

    def setup_dynamic_data_fields(self, dynamic_columns):
        """根据动态列配置设置数据字段 - 表格形式，支持平行样多次测试与列合并

        平行样测试次数 > 1 时渲染多行；isColumnMerge=1 的列跨行合并（共享一个值）。
        data_fields[col_code]：合并列 -> StringVar；非合并列 -> list[StringVar]（每行一个）。
        """
        # 保存当前用户输入的值，避免重新创建字段时丢失
        current_values = {}
        if hasattr(self, 'data_fields'):
            for col_code, var in self.data_fields.items():
                if isinstance(var, list):
                    current_values[col_code] = [v.get() for v in var]
                else:
                    current_values[col_code] = var.get()

        # 清空现有字段
        for widget in self.dynamic_fields_frame.winfo_children():
            widget.destroy()

        self.data_fields = {}
        self.merge_columns = set()

        if not dynamic_columns:
            # 显示无配置信息
            info_label = ttk.Label(self.dynamic_fields_frame,
                                   text="未获取到动态列配置，请检查网络连接或系统配置",
                                   foreground="orange")
            info_label.pack(pady=20)

            self.setup_fixed_remark_field()
            return

        sorted_columns = sorted(dynamic_columns, key=lambda x: x.get('columeOrder', 0))

        # 记录合并列
        for col in sorted_columns:
            if self._is_merge_column(col):
                self.merge_columns.add(col.get('columeCode', f'dynamic{col.get("id")}'))

        # 组分列 + 维度：多组分方法 M 个组分块、每块 N 行平行；单组分 M=1（退化为原平行样表格）
        comp_col_code = next(
            (c.get('columeCode', f'dynamic{c.get("id")}') for c in sorted_columns
             if self._is_component_column(c)), None)
        M, N, comp_names = self._compute_grid_dims(comp_col_code)
        self.test_run_count = N
        total_rows = M * N
        # 实验配置里的分析记录（按 record_index 顺序），用于回填随组分/方法固有的列值
        records = (getattr(self, 'experiment_config', {}) or {}).get('ocAnalysisRecordList') or []

        # 项目名称（只读列）：按组分块取该块 projectId -> projectName，便于核对采集回填归属
        name_by_pid = {p.get('projectId'): p.get('projectName', '')
                       for p in getattr(self, 'filtered_projects', []) if isinstance(p, dict)}
        _block_groups = {}
        if comp_col_code and records:
            for _r in records:
                _block_groups.setdefault(_r.get(comp_col_code), []).append(_r)
            block_pids = [g[0].get('projectId') for g in _block_groups.values()]
        else:
            block_pids = []
            for _r in records:
                _pid = _r.get('projectId')
                if _pid is not None and _pid not in block_pids:
                    block_pids.append(_pid)
        block_project_names = [str(name_by_pid.get(pid) or '') for pid in block_pids]

        # ---- 表格 ----
        table = ttk.Frame(self.dynamic_fields_frame)
        table.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        table.columnconfigure(0, weight=0)
        table.columnconfigure(1, weight=0)  # 项目名称（只读）
        for c_idx in range(2, len(sorted_columns) + 2):
            table.columnconfigure(c_idx, weight=1)

        # 表头：次数 + 项目名称 + 各动态列
        ttk.Label(table, text="次数", width=6).grid(row=0, column=0, padx=3, pady=(0, 4), sticky='w')
        ttk.Label(table, text="项目名称", width=14).grid(row=0, column=1, padx=3, pady=(0, 4), sticky='w')
        for c_idx, col in enumerate(sorted_columns, start=2):
            ttk.Label(table, text=col.get('columeName', ''), width=10).grid(
                row=0, column=c_idx, padx=3, pady=(0, 4), sticky='w')

        # 次数列（只读）：每个组分块内 1..N
        for i in range(total_rows):
            ttk.Label(table, text=str((i % N) + 1), width=6).grid(
                row=i + 1, column=0, padx=3, pady=2, sticky='w')

        # 项目名称列（只读，不写入 data_fields、不参与提交）：每个组分块跨 N 行
        for b in range(M):
            _pname = block_project_names[b] if b < len(block_project_names) else ''
            _lbl = ttk.Label(table, text=_pname, width=14, anchor='w')
            _lbl.grid(row=b * N + 1, column=1, rowspan=N, padx=3, pady=2, sticky='w')
            if _pname:
                _bind_tooltip(_lbl, _pname)

        # 各动态列控件
        for c_idx, col in enumerate(sorted_columns, start=2):
            col_id = col.get('id')
            col_code = col.get('columeCode', f'dynamic{col_id}')
            edit_type = col.get('editType', 'EDIT_TYPE_TEXT')
            default_val = col.get('defaultVal', '')
            select_values = self.get_select_values(col_id, col_code) if edit_type == 'EDIT_TYPE_SELECT' else []

            if col_code in self.merge_columns:
                # 合并列：跨整表所有行共享一个值
                cur = current_values.get(col_code)
                var = tk.StringVar(value=cur if isinstance(cur, str) else default_val)
                cell = ttk.Frame(table)
                cell.grid(row=1, column=c_idx, rowspan=total_rows, padx=3, pady=2, sticky='nsew')
                self._create_cell_widget(cell, var, select_values)
                self.data_fields[col_code] = var
            elif col_code == comp_col_code:
                # 组分列：M 个只读格，每块跨 N 行；值取自记录（方法固有，不从 current_values 还原）。
                # 同一块内 N 个位置共享同一 StringVar，按 record_index 提交即可取到该组分名。
                flat = []
                for b in range(M):
                    name = comp_names[b] if b < len(comp_names) else default_val
                    var = tk.StringVar(value=name)
                    cell = ttk.Frame(table)
                    cell.grid(row=b * N + 1, column=c_idx, rowspan=N, padx=3, pady=2, sticky='nsew')
                    self._create_readonly_cell(cell, var)
                    flat.extend([var] * N)
                self.data_fields[col_code] = flat
            else:
                # 普通列：每行一个值（扁平 M*N，与 record_index 对齐）
                cur_list = current_values.get(col_code) if isinstance(current_values.get(col_code), list) else None
                vars_list = []
                for i in range(total_rows):
                    # 优先用户已输入值；其次回填记录里的值（如称重因子a/相对分子质量b/定量限等
                    # 随组分变化的固有值，来自 ocAnalysisRecordList）；最后用列默认值
                    if cur_list and i < len(cur_list) and cur_list[i] not in (None, ''):
                        val = cur_list[i]
                    elif i < len(records):
                        rec_val = records[i].get(col_code)
                        val = rec_val if rec_val not in (None, '') else default_val
                    else:
                        val = default_val
                    v = tk.StringVar(value=val)
                    cell = ttk.Frame(table)
                    cell.grid(row=i + 1, column=c_idx, padx=3, pady=2, sticky='ew')
                    self._create_cell_widget(cell, v, select_values)
                    vars_list.append(v)
                self.data_fields[col_code] = vars_list

        # 添加固定备注字段（在动态列之后）
        self.setup_fixed_remark_field()

    def _is_merge_column(self, col):
        """该动态列是否跨测试次数合并（来自方法设置 isColumnMerge）"""
        try:
            return int(col.get('isColumnMerge') or 0) == 1
        except (TypeError, ValueError):
            return False

    def _compute_test_run_count(self, experiment_config):
        """测试次数 N = 单个项目在 ocAnalysisRecordList 中的记录数（默认 1）"""
        records = (experiment_config or {}).get('ocAnalysisRecordList') or []
        per_project = {}
        for r in records:
            pid = r.get('projectId')
            if pid is None:
                continue
            per_project[pid] = per_project.get(pid, 0) + 1
        return max(per_project.values()) if per_project else 1

    def _is_component_column(self, col):
        """该动态列是否为多组分方法的组分列（来自方法设置 isMutiPolyColume）"""
        try:
            return int(col.get('isMutiPolyColume') or 0) == 1
        except (TypeError, ValueError):
            return False

    def _compute_grid_dims(self, comp_col_code):
        """计算表格维度 (M 组分数, N 平行数, comp_names 组分名列表)。
        多组分：按组分列值对 ocAnalysisRecordList 分组（保序）；
        无组分列：M=不同项目数(每项目一块)，N=单项目最大记录数(平行数)。"""
        records = (getattr(self, 'experiment_config', {}) or {}).get('ocAnalysisRecordList') or []
        if comp_col_code and records:
            groups = {}
            for r in records:
                groups.setdefault(r.get(comp_col_code), []).append(r)
            if groups:
                return len(groups), max(len(g) for g in groups.values()), list(groups.keys())
        pids = []
        for r in records:
            pid = r.get('projectId')
            if pid is not None and pid not in pids:
                pids.append(pid)
        return len(pids) or 1, self._compute_test_run_count(getattr(self, 'experiment_config', {})), []

    def _create_cell_widget(self, parent, var, select_values):
        """在表格单元格里创建输入控件：有下拉选项用 Combobox，否则 Entry"""
        if select_values:
            cb = ttk.Combobox(parent, textvariable=var, values=select_values, width=10)
            cb.pack(fill=tk.BOTH, expand=True)
            if var.get():
                cb.set(var.get())
            return cb
        entry = ttk.Entry(parent, textvariable=var, width=10)
        entry.pack(fill=tk.BOTH, expand=True)
        return entry

    def _create_readonly_cell(self, parent, var):
        """只读单元格（组分名等固有值）：只读 Entry，纵向填满合并区域"""
        entry = ttk.Entry(parent, textvariable=var, width=10, state='readonly')
        entry.pack(fill=tk.BOTH, expand=True)
        return entry

    def update_status(self, message, color="black"):
        """更新状态栏"""
        self.status_var.set(f"状态: {message}")

    def log(self, message):
        """记录日志"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        print(f"[{timestamp}] {message}")
    def set_current_date(self):
        """设置当前日期为开始和结束日期"""
        current_date = datetime.now().strftime("%Y-%m-%d")
        self.start_date_var.set(current_date)
        self.end_date_var.set(current_date)

    def select_all_projects(self):
        """选择所有项目"""
        for var in self.project_vars:
            var.set(True)
        self.update_selected_count()

    def deselect_all_projects(self):
        """取消选择所有项目"""
        for var in self.project_vars:
            var.set(False)
        self.update_selected_count()

    def update_selected_count(self):
        """更新已选择项目数量显示"""
        selected_count = sum(1 for var in self.project_vars if var.get())
        self.selected_count_var.set(f"已选择: {selected_count} 个项目")

        # 同时更新分组信息
        self.update_group_info()

    def update_group_info(self):
        """更新分组信息显示"""
        selected_indices = [i for i, var in enumerate(self.project_vars) if var.get()]
        if not selected_indices:
            self.group_info_var.set("")
            return

        # 获取选中的项目并按检测方法分组
        selected_projects = []
        for index in selected_indices:
            if index < len(self.filtered_projects):
                project = self.filtered_projects[index]
                selected_projects.append(project)

        # 按检测方法分组
        method_groups = {}
        for project in selected_projects:
            method_name = project.get('standardNo', '未知方法')
            if method_name not in method_groups:
                method_groups[method_name] = []
            method_groups[method_name].append(project)

        if method_groups:
            group_info = f"检测到 {len(method_groups)} 个方法组"
            self.group_info_var.set(group_info)
        else:
            self.group_info_var.set("")

    def query_samples(self):
        """根据条件查询样品信息 - 优化响应速度"""
        # 获取查询条件
        sample_code = self.sample_code_var.get().strip()
        project_name = self.project_name_var.get().strip()
        method_name = self.method_name_var.get().strip()
        # 使用复选框状态
        retest_checked = self.retest_var.get()

        # 清空现有的结果
        for widget in self.search_results_frame.winfo_children():
            widget.destroy()
        self.project_vars = []
        self.project_id_to_var = {}
        self.filtered_projects = []

        # 显示查询状态
        self.update_status("正在查询样品信息...", "blue")

        # 强制更新界面，显示状态变化
        self.root.update_idletasks()

        # 使用API查询样品信息 - API已经过滤了未登记数据
        all_projects = self.api.query_samples_by_conditions(
            sample_code=sample_code,
            project_name=project_name,
            method_name=method_name,
            retest_checked=retest_checked,
            log_func=self.log
        )

        # 移除冗余的客户端过滤，因为API已经确保了只返回未登记数据
        filtered_projects = all_projects

        # 应用额外的客户端过滤（确保完全匹配）
        filtered_projects = self.apply_client_filters(filtered_projects, sample_code, project_name,
                                                      retest_checked)
        self.filtered_projects = filtered_projects

        # 显示查询结果
        self.display_search_results(filtered_projects)

        # 更新状态
        self.update_status(f"查询完成: 找到 {len(filtered_projects)} 条记录", "green")
        self.result_count_var.set(f"查询结果: {len(filtered_projects)} 条")

    def apply_client_filters(self, projects, sample_code, project_name, retest_checked):
        """在客户端应用额外的过滤条件（确保完全匹配）"""
        filtered = projects

        # 样品编号过滤（模糊匹配）- 只有当输入了样品编号时才应用
        if sample_code:
            filtered = [p for p in filtered if sample_code.lower() in p.get('sampleCode', '').lower()]

        # 项目名称过滤（模糊匹配）- 只有当输入了项目名称时才应用
        if project_name:
            # 使用更宽松的匹配方式，确保包含关键字的项目都能被找到
            filtered = [p for p in filtered if project_name.lower() in p.get('projectName', '').lower()]

        # 检测方法过滤已移除：服务端 query_samples_by_conditions 已按 decideProjectMethodName
        # 过滤且结果精准，客户端再按 standardNo 子串过滤会误杀（standardNo 不含 subMethodName
        # 如「单组份」，导致完整方法名查询 389 条全被删）。

        # 注销复测过滤 - 基于 oldSampleProjectId 字段:勾选=仅复测(有值),未勾选=仅非复测(为空)
        if retest_checked:
            # 只显示 oldSampleProjectId 有值的记录（注销复测）
            filtered = [p for p in filtered if p.get('oldSampleProjectId') is not None]
            if not filtered:
                pass
        else:
            # 未勾选：只显示非复测记录（oldSampleProjectId 为空）
            filtered = [p for p in filtered if p.get('oldSampleProjectId') is None]

        return filtered

    def display_search_results(self, projects):
        """显示查询结果"""
        if not projects:
            ttk.Label(self.search_results_frame, text="未找到符合条件的记录",
                      foreground="gray").grid(row=0, column=0, columnspan=6, pady=20)
            return

        # 表头与数据行共用同一 grid，列宽由 grid 统一分配，表头与各行自动对齐
        for c, (text, w) in enumerate(zip(
                ["选择", "样品编号", "样品名称", "项目", "检测方法", "注销复测"],
                [8, 15, 15, 25, 20, 10])):
            ttk.Label(self.search_results_frame, text=text, width=w).grid(row=0, column=c, sticky='w')
        ttk.Separator(self.search_results_frame, orient='horizontal').grid(
            row=1, column=0, columnspan=6, sticky='ew', pady=5)

        for i, project in enumerate(projects):
            r = i + 2
            var = tk.BooleanVar()
            self.project_vars.append(var)
            self.project_id_to_var[project.get('projectId')] = var

            checkbox = ttk.Checkbutton(self.search_results_frame, variable=var, style=self.large_cb_style)
            checkbox.grid(row=r, column=0, sticky='w')
            checkbox.bind('<Button-1>', lambda e, v=var: self.on_project_selected())

            ttk.Label(self.search_results_frame, text=project.get('sampleCode', ''), width=15).grid(row=r, column=1, sticky='w')
            ttk.Label(self.search_results_frame, text=project.get('sampleName', ''), width=15).grid(row=r, column=2, sticky='w')
            proj_lbl = ttk.Label(self.search_results_frame, text=project.get('projectName', ''), width=25)
            proj_lbl.grid(row=r, column=3, sticky='w')
            _bind_tooltip(proj_lbl, project.get('projectName', ''))
            meth_lbl = ttk.Label(self.search_results_frame, text=project.get('standardNo', ''), width=20)
            meth_lbl.grid(row=r, column=4, sticky='w')
            _bind_tooltip(meth_lbl, project.get('standardNo', ''))
            retest_text = "是" if project.get('isRetest', False) else "否"
            ttk.Label(self.search_results_frame, text=retest_text, width=10).grid(row=r, column=5, sticky='w')

    def clear_query_conditions(self):
        """清空查询条件"""
        self.sample_code_var.set("")
        self.project_name_var.set("")
        self.method_name_var.set("")
        # 重置复选框状态
        self.retest_var.set(False)

        # 清空查询结果
        for widget in self.search_results_frame.winfo_children():
            widget.destroy()
        self.project_vars = []
        self.project_id_to_var = {}
        self.filtered_projects = []

        # 重置统计信息
        self.result_count_var.set("查询结果: 0 条")
        self.selected_count_var.set("已选择: 0 个项目")
        self.group_info_var.set("")

    def find_mass_field(self):
        """查找质量字段"""
        for field_code, var in self.data_fields.items():
            # 合并列=单个 StringVar；非合并列=list（每次测试一个）
            vals = var if isinstance(var, list) else [var]
            for v in vals:
                field_value = v.get()
                # 简单的启发式方法：如果字段值看起来像质量值
                if field_value and field_value.replace('.', '').isdigit():
                    if float(field_value) > 0:
                        return field_code
        return None

    def submit_single_method_group(self, method_name, projects, configure_order=None):
        """提交单个检测方法组的数据 - 保存实验编号并自动提交标准溶液"""
        try:
            # 直接从界面获取配置序号，不依赖参数
            actual_configure_order = None
            if hasattr(self, 'solution_type_var') and self.solution_type_var is not None:
                raw_value = self.solution_type_var.get()
                if raw_value and raw_value.strip() and raw_value.strip().lower() != 'none':
                    actual_configure_order = raw_value.strip()

                    # 提前检查标准溶液审核状态 - 如果未审核，完全阻止提交
                    configure_id, error_msg = self.api.get_solution_configure_id(actual_configure_order, self.log)

                    if not configure_id:
                        self.log(f"标准溶液未审核: {error_msg}，完全阻止数据提交")
                        return False, error_msg

            project_ids = [str(p.get('projectId')) for p in projects if p.get('projectId')]
            project_names = [p.get('projectName', '') for p in projects]

            if not project_ids:
                return False, None

            # 关键修复：确保清空时传递所有项目ID
            sample_project_ids_str = ",".join(project_ids)

            # 先清除暂存数据 - 传递所有项目ID
            clear_success = self.api.clear_experiment_cache(sample_project_ids_str, self.log)

            if not clear_success:
                self.log(f"警告: 清除暂存数据失败，项目ID: {sample_project_ids_str}")

            # 重新获取配置信息，确保使用子方法
            first_project = projects[0]
            sample_id = first_project.get('sampleId')
            result_checkin_ids = ""

            # 获取方法ID
            initial_config = self.api.get_experiment_config(
                sample_project_ids_str,
                method_name,
                result_checkin_ids,
                sample_id,
                self.log
            )

            if not initial_config:
                error_msg = f"无法获取实验配置，请检查网络连接或项目状态。方法: {method_name}, 项目ID: {sample_project_ids_str}"
                self.log(error_msg)
                success = self.fallback_config_retrieval(method_name, projects, sample_project_ids_str, sample_id)
                if not success:
                    raise Exception(error_msg)
                else:
                    initial_config = self.api.get_experiment_config(
                        sample_project_ids_str,
                        method_name,
                        result_checkin_ids,
                        sample_id,
                        self.log
                    )
                    if not initial_config:
                        raise Exception(error_msg)

            oc_method_settings = initial_config.get('ocMethodSettings', {})
            method_id = oc_method_settings.get('methodId')

            # 检查子方法切换
            actual_method_id = method_id
            actual_method_name = method_name

            if method_id and str(method_id) in self.api.sub_method_map:
                sub_method_info = self.api.sub_method_map[str(method_id)]
                actual_method_id = sub_method_info['sub_method_id']

                # 关键步骤：更新检测方法 - 传递项目名称
                update_success = self.api.update_method(
                    sample_project_ids_str,
                    actual_method_id,
                    [p.get('projectName', '') for p in projects],
                    self.log
                )

                # 获取子方法标准号
                sub_method_standard_no = self.api.get_method_standard_no_by_id(actual_method_id, self.log)
                if sub_method_standard_no:
                    actual_method_name = sub_method_standard_no

            all_configs = self.api.get_all_configs(
                sample_project_ids_str,
                method_name,
                result_checkin_ids,
                sample_id,
                self.log,
                method_id,  # 原始方法ID
                [p.get('projectName', '') for p in projects]  # 项目名称列表
            )

            if not all_configs:
                error_msg = "无法获取配置信息，请检查网络连接或系统状态"
                raise Exception(error_msg)

            # 提取各个配置
            self.experiment_config = all_configs.get('experiment')
            self.equipment_config = all_configs.get('equipment')
            # 注意：此处不再 _refresh_equipment_display() —— 提交流程重绘会清空用户在
            # 「查询设备」里勾选的真实设备（含 id/raw），导致主检设备提交拿不到数据。
            # 设备清单的默认展示由「获取动态配置」流程负责。
            self.units_config = all_configs.get('units')
            self.round_methods_config = all_configs.get('round_methods')
            self.calc_methods_config = all_configs.get('calc_methods')
            self.dynamic_columns = all_configs.get('dynamic_columns')

            # 从配置中获取实际使用的方法信息
            actual_method_name = all_configs.get('actual_method_name', method_name)
            actual_method_id = all_configs.get('actual_method_id')

            # 保存实际的方法信息
            self.actual_method_name = actual_method_name
            self.actual_method_id = actual_method_id

            # 使用 API 的统一实验编号生成方法（替换原来的 generate_experiment_code）
            experiment_code = self.api.generate_experiment_code(
                method_name=actual_method_name,
                log_func=self.log
            )

            # 更新UI动态字段 - 避免丢失用户输入
            if not hasattr(self, 'data_fields') or not self.data_fields:
                self.setup_dynamic_data_fields(self.dynamic_columns)
            else:
                self.dynamic_columns = all_configs.get('dynamic_columns', [])

            # 称样设备：用「查询设备」勾选的设备覆盖方法默认（称样设备随 saveOcExperiment 提交，
            # 实验数据里的称样字段取自 equipment_config，故提交前注入用户选择；模型为单台，取第 1 台）
            sel_weigh = [it for var, it in getattr(self, 'weighing_eq_vars', []) if var.get()]
            if sel_weigh and isinstance(self.equipment_config, dict):
                raw = sel_weigh[0].get('raw') or {}
                wid = str(raw.get('id') or sel_weigh[0].get('id') or '')
                wno = (raw.get('no') or '').strip()
                wname = (raw.get('name') or '').strip()
                cod = raw.get('checkOutDate')
                wdate = cod[:10] if isinstance(cod, str) else (str(cod)[:10] if cod else '')
                self.equipment_config['weighingEquipmentId'] = wid
                self.equipment_config['weighingEquipmentBaseName'] = wname
                self.equipment_config['weighingEquipment'] = f"{wno},{wname},{wdate}" if wno else wname
                if len(sel_weigh) > 1:
                    self.log(f"称样设备仅取第 1 台（实验数据为单台），共勾选 {len(sel_weigh)} 台")
                self.log(f"称样设备注入实验数据: id={wid}, {self.equipment_config['weighingEquipment']}")

            # 构建实验数据 - 使用修复后的方法
            experiment_data = self.build_grouped_experiment_data(projects, experiment_code, actual_method_name)

            # 注入已上传的谱图关联（谱图上传标签页产出的记录）
            if getattr(self, "spectrum_uploaded", None):
                experiment_data["fileIds"] = ",".join(str(f["fileId"]) for f in self.spectrum_uploaded)
                group_pids = ",".join(str(p.get("projectId")) for p in projects if p.get("projectId"))
                experiment_data["spectrumJsonList"] = json.dumps([
                    {"fileId": f["fileId"], "fileName": f["fileName"], "projectId": group_pids}
                    for f in self.spectrum_uploaded
                ])

            # 提交数据
            success, _ = self.api.submit_experiment_data(experiment_data, actual_method_name, self.log)

            if success:
                # 实验编号由服务端生成；saveOcExperiment 响应只回项目id，故另查 getOcExperiment 取真实编号
                real_code = experiment_code
                experiment_id = 0
                first_pid = str(project_ids[0]) if project_ids else ""
                if first_pid:
                    try:
                        cfg = self.api.get_experiment_config(first_pid, "", "", "", self.log)
                        if isinstance(cfg, dict):
                            real_code = self.api.extract_experiment_code(cfg, experiment_code)
                            try:
                                experiment_id = int(cfg.get('id') or 0)
                            except (TypeError, ValueError):
                                experiment_id = 0
                    except Exception as e:
                        if self.log:
                            self.log(f"读取真实实验编号失败: {e}")
                # 保存实验编号
                if not hasattr(self, 'experiment_codes'):
                    self.experiment_codes = {}

                for project_id in project_ids:
                    self.experiment_codes[project_id] = real_code

                # 更新实验编号显示
                self.update_experiment_code_display()

                # 提交主检设备（saveMainEqubment）—— 集成进提交数据流程
                selected_main = [it for var, it in getattr(self, 'main_eq_vars', []) if var.get()]
                if not selected_main:
                    self.log("主检设备未提交：未勾选任何主检设备（环境与设备页）")
                elif not experiment_id:
                    self.log(f"主检设备未提交：experimentId=0（{len(selected_main)} 台已勾选，但实验编号未生成）")
                else:
                    eq_ok, _ = self.api.save_main_equipment(experiment_id, selected_main, self.log)
                    self.log(f"主检设备{'提交成功' if eq_ok else '提交失败'}: "
                             f"{','.join(it.get('id', '') for it in selected_main)}")

                # 自动提交标准溶液（如果配置序号已填写）
                if actual_configure_order:
                    # 提交标准溶液 - 现在返回两个值
                    solution_success, solution_error = self.api.submit_solution_with_experiment(
                        project_ids,
                        actual_configure_order,
                        self.experiment_codes,
                        self.log
                    )

                    if solution_success:
                        if hasattr(self, 'solution_status_var'):
                            self.solution_status_var.set("标准溶液自动提交成功")
                        return True, True
                    else:
                        self.log(f"标准溶液自动提交失败: {actual_configure_order}, 错误: {solution_error}")
                        if hasattr(self, 'solution_status_var'):
                            self.solution_status_var.set("标准溶液自动提交失败")

                        # 检查是否为未审核错误
                        if "未审核" in str(solution_error):
                            messagebox.showwarning(
                                "标准溶液未审核",
                                f"配置序号 {actual_configure_order} 未审核，无法关联，请审核后操作。\n\n错误详情: {solution_error}"
                            )
                            return True, "未审核"
                        else:
                            messagebox.showerror(
                                "标准溶液提交错误",
                                f"标准溶液自动提交失败: {solution_error}"
                            )
                            return True, False
                else:
                    return True, None

            return False, None

        except Exception as e:
            import traceback
            self.log(f"提交检测方法组 {method_name} 时发生异常: {str(e)}")
            self.log(f"异常详情: {traceback.format_exc()}")
            return False, None

    def clear_data(self):
        """清空数据 - 包括固定备注字段"""
        # 清空环境条件
        self.temperature_var.set("21.0")
        self.humidity_var.set("51.0")
        self.set_current_date()

        # 重置动态字段值为空值
        if hasattr(self, 'dynamic_columns') and self.dynamic_columns:
            for col in self.dynamic_columns:
                col_code = col.get('columeCode', '')
                default_val = col.get('defaultVal', '')
                if col_code in self.data_fields:
                    # 组分列是方法固有值（预填组分名），清空时不重置
                    if self._is_component_column(col):
                        continue
                    field = self.data_fields[col_code]
                    if isinstance(field, list):
                        for v in field:
                            v.set(default_val)
                    else:
                        field.set(default_val)

        # 清空固定备注字段
        if hasattr(self, 'remark_text'):
            self.remark_text.delete('1.0', tk.END)

    def switch_method_id(self):
        """手动切换选中项目的方法ID - 对选中项目调用 update_method 切到输入的目标方法ID"""
        if not self.login_system.current_user:
            messagebox.showerror("错误", "请先登录系统")
            return

        target_id = self.switch_method_id_var.get().strip()
        if not target_id:
            messagebox.showerror("错误", "请输入要切换到的目标方法ID")
            return

        # 取选中的项目（与提交流程一致的选中逻辑）
        selected_indices = [i for i, var in enumerate(self.project_vars) if var.get()]
        if not selected_indices:
            messagebox.showerror("错误", "请至少选择一个检测项目")
            return

        selected_projects = [self.filtered_projects[i] for i in selected_indices if i < len(self.filtered_projects)]
        project_ids = [str(p.get('projectId')) for p in selected_projects if p.get('projectId')]
        if not project_ids:
            messagebox.showerror("错误", "选中的项目没有有效的项目ID")
            return

        sample_project_ids_str = ",".join(project_ids)
        project_names = [p.get('projectName', '') for p in selected_projects]

        self.log(f"开始切换方法ID: 项目[{sample_project_ids_str}] -> 方法ID {target_id}")
        success = self.api.update_method(sample_project_ids_str, target_id, project_names, self.log)

        if success:
            # update_method 只改服务端 decideProjectMethodId，本地 standardNo 仍是旧值，
            # 会导致 get_dynamic_config 按旧方法取配置。先同步选中项目到目标方法的 standardNo。
            tgt = next((m for m in getattr(self, 'switchable_methods', [])
                        if str(m.get('decideProjectMethodId')) == target_id), None)
            if tgt and tgt.get('standardNo'):
                for p in selected_projects:
                    p['standardNo'] = tgt['standardNo']
            self.log(f"已切换到方法ID: {target_id}，自动刷新动态配置...")
            self.get_dynamic_config()  # 切换后自动刷新动态输入字段（随新方法ID变化）
        else:
            self.log(f"方法ID切换失败: 目标ID={target_id}")
            messagebox.showerror("失败", "方法ID切换失败，请检查目标方法ID是否正确或网络连接")

def main():
    parser = argparse.ArgumentParser(description='检测数据录入系统')
    parser.add_argument('--result-checkin-ids', help='结果录入ID')
    parser.add_argument('--sample-id', help='样品ID')
    parser.add_argument('--sample-project-ids', help='样品项目ID')

    args = parser.parse_args()

    root = ttkb.Window(themename="sandstone-light")
    app = DetectionEntrySystem(
        root,
        result_checkin_ids=args.result_checkin_ids,
        sample_id=args.sample_id,
        sample_project_ids=args.sample_project_ids
    )

    root.mainloop()


if __name__ == "__main__":
    main()