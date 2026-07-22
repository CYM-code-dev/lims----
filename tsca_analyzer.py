# detection_entry_enhanced.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext
import requests
import json
import time
import pandas as pd
from datetime import datetime, timedelta
import urllib3
from urllib3.exceptions import InsecureRequestWarning
import os
import sys
import argparse
import webbrowser

# 禁用SSL警告
urllib3.disable_warnings(InsecureRequestWarning)

# 导入登录系统
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from login import MultiUserLoginSystem


class DetectionEntrySystem:
    def __init__(self, root, result_checkin_ids=None, sample_id=None, sample_project_ids=None):
        self.root = root
        self.root.title("检测数据录入系统")
        self.root.geometry("800x600")  # 增加窗口高度以容纳多选列表

        # 使用登录系统的会话
        self.login_system = MultiUserLoginSystem()

        # 动态参数（从URL获取）
        self.current_result_checkin_id = result_checkin_ids
        self.current_sample_id = sample_id
        self.current_sample_project_ids = sample_project_ids

        # 存储从API获取的样品信息
        self.sample_info = {}

        # 存储查询到的项目列表
        self.detected_projects = []

        # 存储项目选择状态
        self.project_vars = []

        # 存储项目ID到选择状态的映射
        self.project_id_to_var = {}

        # 存储动态列配置
        self.dynamic_columns = []

        # 初始化手动输入相关的变量
        self.detection_no_var = None
        self.sample_small_no_var = None
        self.detection_project_var = None
        self.standard_no_var = None
        self.is_retest_var = None
        self.old_sample_project_id_var = None
        self.old_sample_project_id_entry = None

        # 创建界面
        self.setup_ui()

        # 检查登录状态
        self.check_login_status()

    def check_login_status(self):
        """检查登录状态"""
        if self.login_system.load_session() and self.login_system.verify_session():
            self.update_status(f"已登录: {self.login_system.current_user}", "green")
            self.log(f"自动登录成功: {self.login_system.current_user}")
        else:
            self.update_status("未登录", "red")
            self.log("请先登录系统")

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

        # 样品信息标签页
        sample_frame = ttk.Frame(notebook, padding="10")
        notebook.add(sample_frame, text="样品信息")

        # 环境条件标签页
        env_frame = ttk.Frame(notebook, padding="10")
        notebook.add(env_frame, text="环境条件")

        # 标准物质信息标签页
        reference_frame = ttk.Frame(notebook, padding="10")
        notebook.add(reference_frame, text="标准物质信息")

        # 检测数据标签页
        data_frame = ttk.Frame(notebook, padding="10")
        notebook.add(data_frame, text="检测数据")

        # 设置各个标签页
        self.setup_sample_tab(sample_frame)
        self.setup_env_tab(env_frame)
        self.setup_reference_tab(reference_frame)
        self.setup_data_tab(data_frame)

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

        # 操作按钮
        button_frame = ttk.Frame(status_frame)
        button_frame.pack(side=tk.RIGHT)

        ttk.Button(button_frame, text="重新登录", command=self.relogin).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="刷新状态", command=self.check_login_status).pack(side=tk.LEFT)

    def setup_sample_tab(self, parent):
        """设置样品信息标签页"""
        # 样品信息区域
        sample_frame = ttk.LabelFrame(parent, text="样品信息", padding="10")
        sample_frame.pack(fill=tk.X, pady=(0, 10))

        # 第一行 - 手动输入样品编号信息
        row1_frame = ttk.Frame(sample_frame)
        row1_frame.pack(fill=tk.X, pady=5)

        ttk.Label(row1_frame, text="检测编号(detectionNo):").pack(side=tk.LEFT)
        self.detection_no_var = tk.StringVar(value="TS25100593")  # 设置默认值
        ttk.Entry(row1_frame, textvariable=self.detection_no_var, width=15).pack(side=tk.LEFT, padx=(5, 10))

        ttk.Label(row1_frame, text="样品小号(sampleSmallNo):").pack(side=tk.LEFT)
        self.sample_small_no_var = tk.StringVar(value="011")  # 设置默认值
        ttk.Entry(row1_frame, textvariable=self.sample_small_no_var, width=15).pack(side=tk.LEFT, padx=(5, 10))

        # 第二行 - 检测项目列表（多选）
        row2_frame = ttk.Frame(sample_frame)
        row2_frame.pack(fill=tk.X, pady=5)

        ttk.Label(row2_frame, text="检测项目:").pack(side=tk.LEFT)

        # 创建项目选择框架
        projects_frame = ttk.LabelFrame(sample_frame, text="请选择检测项目")
        projects_frame.pack(fill=tk.X, pady=10)

        # 创建滚动框架用于项目选择
        canvas = tk.Canvas(projects_frame, height=150)
        scrollbar = ttk.Scrollbar(projects_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 第三行 - 注销复测信息
        row3_frame = ttk.Frame(sample_frame)
        row3_frame.pack(fill=tk.X, pady=5)

        self.is_retest_var = tk.BooleanVar()
        ttk.Checkbutton(row3_frame, text="注销复测", variable=self.is_retest_var,
                        command=self.toggle_retest_fields).pack(side=tk.LEFT)

        ttk.Label(row3_frame, text="原样品项目ID(oldSampleProjectId):").pack(side=tk.LEFT, padx=(20, 5))
        self.old_sample_project_id_var = tk.StringVar()
        self.old_sample_project_id_entry = ttk.Entry(row3_frame, textvariable=self.old_sample_project_id_var, width=15,
                                                     state="disabled")
        self.old_sample_project_id_entry.pack(side=tk.LEFT, padx=5)

        # 手动输入按钮
        manual_button_frame = ttk.Frame(sample_frame)
        manual_button_frame.pack(fill=tk.X, pady=5)

        ttk.Button(manual_button_frame, text="查询样品信息", command=self.query_sample_info).pack(side=tk.LEFT,
                                                                                                  padx=(0, 10))
        ttk.Button(manual_button_frame, text="应用手动输入信息", command=self.apply_manual_input).pack(side=tk.LEFT,
                                                                                                       padx=(0, 10))
        ttk.Button(manual_button_frame, text="清空手动输入", command=self.clear_manual_input).pack(side=tk.LEFT)

        # 项目选择按钮
        project_button_frame = ttk.Frame(sample_frame)
        project_button_frame.pack(fill=tk.X, pady=5)

        ttk.Button(project_button_frame, text="全选项目", command=self.select_all_projects).pack(side=tk.LEFT,
                                                                                                 padx=(0, 10))
        ttk.Button(project_button_frame, text="取消全选", command=self.deselect_all_projects).pack(side=tk.LEFT)

    def setup_env_tab(self, parent):
        """设置环境条件标签页"""
        # 环境条件区域
        env_frame = ttk.LabelFrame(parent, text="环境条件", padding="10")
        env_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 温湿度和时间
        env_row1 = ttk.Frame(env_frame)
        env_row1.pack(fill=tk.X, pady=10)

        ttk.Label(env_row1, text="温度(℃):").pack(side=tk.LEFT)
        self.temperature_var = tk.StringVar(value="25")
        ttk.Entry(env_row1, textvariable=self.temperature_var, width=10).pack(side=tk.LEFT, padx=(5, 20))

        ttk.Label(env_row1, text="湿度(%RH):").pack(side=tk.LEFT)
        self.humidity_var = tk.StringVar(value="50")
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

        ttk.Button(env_row3, text="设置当前日期", command=self.set_current_date).pack(side=tk.LEFT)

    def setup_reference_tab(self, parent):
        """设置标准物质信息标签页"""
        # 标准物质信息区域
        reference_frame = ttk.LabelFrame(parent, text="标准物质信息", padding="10")
        reference_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 参考物质
        ref_row = ttk.Frame(reference_frame)
        ref_row.pack(fill=tk.X, pady=5)

        ttk.Label(ref_row, text="参考物质:").pack(side=tk.LEFT)
        self.reference_material_var = tk.StringVar()
        ttk.Entry(ref_row, textvariable=self.reference_material_var, width=50).pack(side=tk.LEFT, padx=(5, 20))

        # 其他标准物质信息可以在这里添加
        other_ref_frame = ttk.Frame(reference_frame)
        other_ref_frame.pack(fill=tk.X, pady=5)

        ttk.Label(other_ref_frame, text="标准菌株:").pack(side=tk.LEFT)
        self.standard_strain_var = tk.StringVar()
        ttk.Entry(other_ref_frame, textvariable=self.standard_strain_var, width=50).pack(side=tk.LEFT, padx=(5, 20))

    def setup_data_tab(self, parent):
        """设置检测数据标签页"""
        # 检测数据区域
        data_frame = ttk.LabelFrame(parent, text="检测数据", padding="10")
        data_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 动态字段
        self.data_fields = {}
        field_config = [
            ("试样描述", "dynamic3578", ""),
            ("空白浓度C0(μg/mL)", "dynamic3579", "<0.10"),
            ("样品浓度C1(μg/mL)", "dynamic3580", "<0.10"),
            ("体积V(mL)", "dynamic3581", "25.0"),
            ("质量m(g)", "dynamic3582", ""),
            ("乘积因子F", "dynamic3583", "1")
        ]

        for i, (label, field, default) in enumerate(field_config):
            frame = ttk.Frame(data_frame)
            frame.pack(fill=tk.X, pady=5)

            ttk.Label(frame, text=label, width=15).pack(side=tk.LEFT)
            var = tk.StringVar(value=default)
            entry = ttk.Entry(frame, textvariable=var, width=20)
            entry.pack(side=tk.LEFT, padx=5)
            self.data_fields[field] = var

        # 按钮区域
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="提交数据", command=self.submit_single_data).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="清空数据", command=self.clear_single_data).pack(side=tk.LEFT, padx=(0, 10))

    def update_status(self, message, color="black"):
        """更新状态栏"""
        self.status_var.set(f"状态: {message}")

    def log(self, message):
        """记录日志（简化版）"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        print(f"[{timestamp}] {message}")

    def relogin(self):
        """重新登录"""
        self.check_login_status()

    def set_current_date(self):
        """设置当前日期为开始和结束日期"""
        current_date = datetime.now().strftime("%Y-%m-%d")
        self.start_date_var.set(current_date)
        self.end_date_var.set(current_date)
        self.log("已设置当前日期为开始和结束日期")

    def select_all_projects(self):
        """选择所有项目"""
        for var in self.project_vars:
            var.set(True)
        self.update_selected_count()
        self.log("已选择所有项目")

    def deselect_all_projects(self):
        """取消选择所有项目"""
        for var in self.project_vars:
            var.set(False)
        self.update_selected_count()
        self.log("已取消选择所有项目")

    def update_selected_count(self):
        """更新已选择项目数量显示"""
        selected_count = sum(1 for var in self.project_vars if var.get())
        project_info = f"已选择 {selected_count} 个项目"
        self.project_method_var.set(project_info)

    def get_dynamic_columns(self, sample_project_ids):
        """获取动态列配置 - 修复：为每个项目单独获取动态列"""
        if not self.login_system.current_user:
            self.log("请先登录系统")
            return []

        try:
            # 调用动态列配置API
            response = self.login_system.session.post(
                f"{self.login_system.base_url}/detectionManager/manager/ocExperiment/getDynamicColumns",
                data={
                    "sampleProjectIds": sample_project_ids,
                    "experimentCode": "",
                    "pid": self.login_system.users.get(self.login_system.current_user, {}).get('pid', '377'),
                    "pname": self.login_system.current_user,
                    "loginId": self.login_system.users.get(self.login_system.current_user, {}).get('pid', '377')
                },
                headers={
                    'Accept': 'application/json, text/javascript, */*; q=0.01',
                    'Referer': f'{self.login_system.base_url}/web/detectionResultCheckInCalc.html'
                },
                verify=False
            )

            if response.status_code == 200:
                result = response.json()
                if result.get('success'):
                    dynamic_columns = result.get('resultData', [])
                    self.log(f"获取到 {len(dynamic_columns)} 个动态列配置")
                    return dynamic_columns
                else:
                    self.log(f"获取动态列配置失败: {result.get('errorCtx', {}).get('errorMsg', '未知错误')}")
                    return []
            else:
                self.log(f"获取动态列配置失败: HTTP {response.status_code}")
                return []

        except Exception as e:
            self.log(f"获取动态列配置异常: {str(e)}")
            return []

    def query_sample_by_code(self, sample_code):
        """通过样品编号查询样品信息 - 修复：正确获取实验室编号"""
        if not self.login_system.current_user:
            self.log("请先登录系统")
            return None

        try:
            # 调用样品查询API
            response = self.login_system.session.get(
                f"{self.login_system.base_url}/detectionManager/manager/resultCheckIn/pagePCObjAndSample",
                params={
                    "_search": "false",
                    "nd": int(time.time() * 1000),
                    "pageSize": 30,
                    "pageNo": 1,
                    "sidx": "",
                    "sord": "asc",
                    "acceptStartDate": (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d"),
                    "acceptEndDate": datetime.now().strftime("%Y-%m-%d"),
                    "retrievalWay": "",
                    "decideProjectOrgId": "23",
                    "checkInStatus": "CHECK_IN_STATUS_NO",
                    "screenTime": "",
                    "sort": "0",
                    "keyword": sample_code,
                    "pid": self.login_system.users.get(self.login_system.current_user, {}).get('pid', '377'),
                    "pname": self.login_system.current_user,
                    "loginId": self.login_system.users.get(self.login_system.current_user, {}).get('pid', '377')
                },
                headers={
                    'Accept': 'application/json, text/javascript, */*; q=0.01',
                    'Referer': f'{self.login_system.base_url}/web/detectionResultCheckInListMgt.html'
                },
                verify=False
            )

            if response.status_code == 200:
                result = response.json()
                self.log(f"样品查询结果状态: {result.get('success', False)}")

                # 解析返回的样品信息
                if result.get('success') and result.get('resultData', {}).get('voList'):
                    vo_list = result['resultData']['voList']
                    self.log(f"找到 {len(vo_list)} 条记录")

                    # 提取所有检测项目信息
                    projects = []
                    sample_info = None

                    # 遍历所有记录，找到匹配的样品
                    for sample_data in vo_list:
                        detection_no = sample_data.get('detectionNo', '')
                        small_no = sample_data.get('smallNo', '')
                        current_sample_code = f"{detection_no}{small_no}"

                        # 检查是否匹配查询的样品编号
                        if current_sample_code == sample_code:
                            self.log(f"找到匹配的样品: {current_sample_code}")

                            # 如果是第一个匹配的样品，保存基本信息
                            if sample_info is None:
                                sample_info = {
                                    'sampleId': sample_data.get('sampleId'),
                                    'projectId': sample_data.get('id'),
                                    'sampleCode': current_sample_code,
                                    'sampleName': sample_data.get('sampleName', ''),
                                    'detectionNo': detection_no,
                                    'sampleSmallNo': small_no,
                                    'oldSampleProjectId': sample_data.get('oldSampleProjectId'),
                                    'laboratoryNo': sample_data.get('laboratoryNo', '未知')
                                }

                            # 添加项目信息
                            project_name = sample_data.get('decideProjectName', '')
                            standard_no = sample_data.get('standardNo', '')

                            if project_name:
                                projects.append({
                                    'name': project_name,
                                    'standardNo': standard_no,
                                    'projectId': sample_data.get('id'),
                                    'laboratoryNo': sample_data.get('laboratoryNo',
                                                                    sample_info.get('laboratoryNo', '未知'))
                                })

                    if sample_info and projects:
                        sample_info['projects'] = projects
                        return sample_info
                    else:
                        self.log(f"未找到完全匹配的样品编号: {sample_code}")
                        return None
                else:
                    self.log(f"API返回数据格式异常或没有数据")
                    return None
            else:
                self.log(f"样品查询失败: HTTP {response.status_code}")
                return None

        except Exception as e:
            self.log(f"样品查询异常: {str(e)}")
            return None

    def query_sample_info(self):
        """查询样品信息"""
        detection_no = self.detection_no_var.get().strip()
        sample_small_no = self.sample_small_no_var.get().strip()

        if not detection_no:
            messagebox.showwarning("警告", "请输入检测编号(detectionNo)")
            return

        if sample_small_no:
            full_sample_code = f"{detection_no}{sample_small_no}"
        else:
            full_sample_code = detection_no

        self.log(f"正在查询样品信息: {full_sample_code}")
        sample_info = self.query_sample_by_code(full_sample_code)

        if sample_info:
            # 保存查询到的项目列表
            self.detected_projects = sample_info.get('projects', [])

            # 清空现有的项目选择框
            for widget in self.scrollable_frame.winfo_children():
                widget.destroy()
            self.project_vars = []
            self.project_id_to_var = {}

            # 按实验室编号分组项目
            projects_by_lab = {}
            for project in self.detected_projects:
                lab_no = project.get('laboratoryNo', '未知实验室')
                if lab_no not in projects_by_lab:
                    projects_by_lab[lab_no] = []
                projects_by_lab[lab_no].append(project)

            # 更新检测项目选择框，按实验室编号分组显示
            for lab_no, projects in projects_by_lab.items():
                # 添加实验室编号标题
                lab_frame = ttk.Frame(self.scrollable_frame)
                lab_frame.pack(fill=tk.X, pady=(5, 2))
                ttk.Label(lab_frame, text=f"实验室编号: {lab_no}", font=('Arial', 9, 'bold')).pack(side=tk.LEFT)

                for i, project in enumerate(projects):
                    project_name = project.get('name', '')
                    standard_no = project.get('standardNo', '')
                    project_id = project.get('projectId')

                    # 创建选择变量
                    var = tk.BooleanVar()
                    self.project_vars.append(var)
                    self.project_id_to_var[project_id] = var

                    # 创建选择框
                    frame = ttk.Frame(self.scrollable_frame)
                    frame.pack(fill=tk.X, pady=1, padx=(20, 0))

                    checkbox = ttk.Checkbutton(frame, variable=var, text=f"{project_name} ({standard_no})")
                    checkbox.pack(side=tk.LEFT)
                    # 绑定选择状态变化事件
                    checkbox.bind('<Button-1>', lambda e, v=var: self.update_selected_count())

            # 自动填充其他信息
            if sample_info.get('oldSampleProjectId'):
                self.is_retest_var.set(True)
                self.old_sample_project_id_var.set(sample_info.get('oldSampleProjectId', ''))
                self.old_sample_project_id_entry.config(state="normal")

            # 更新选择数量显示
            self.update_selected_count()
        else:
            messagebox.showwarning("警告",
                                   f"未找到样品编号为 {full_sample_code} 的记录\n请检查样品编号是否正确，或使用手动输入信息")

    def toggle_retest_fields(self):
        """切换注销复测字段的可用状态"""
        if self.is_retest_var.get():
            self.old_sample_project_id_entry.config(state="normal")
            self.log("已启用注销复测模式")
        else:
            self.old_sample_project_id_entry.config(state="disabled")
            self.old_sample_project_id_var.set("")
            self.log("已禁用注销复测模式")

    def apply_manual_input(self):
        """应用手动输入的信息 - 修复：保留已选择的项目"""
        # 保存当前已选择的项目ID
        selected_project_ids = set()
        for project_id, var in self.project_id_to_var.items():
            if var.get():
                selected_project_ids.add(project_id)

        # 构建完整的样品编号
        detection_no = self.detection_no_var.get().strip()
        sample_small_no = self.sample_small_no_var.get().strip()

        if detection_no:
            if sample_small_no:
                full_sample_code = f"{detection_no}{sample_small_no}"
            else:
                full_sample_code = detection_no

            # 先通过API查询样品信息
            self.log(f"正在查询样品信息: {full_sample_code}")
            sample_info = self.query_sample_by_code(full_sample_code)

            if sample_info:
                # 使用查询到的样品信息
                self.sample_info = sample_info
                self.log(f"成功获取样品信息: {sample_info}")

                # 更新检测项目选择框
                projects = sample_info.get('projects', [])

                # 清空现有的项目选择框
                for widget in self.scrollable_frame.winfo_children():
                    widget.destroy()
                self.project_vars = []
                self.project_id_to_var = {}

                # 按实验室编号分组项目
                projects_by_lab = {}
                for project in projects:
                    lab_no = project.get('laboratoryNo', '未知实验室')
                    if lab_no not in projects_by_lab:
                        projects_by_lab[lab_no] = []
                    projects_by_lab[lab_no].append(project)

                # 更新检测项目选择框，按实验室编号分组显示
                for lab_no, lab_projects in projects_by_lab.items():
                    # 添加实验室编号标题
                    lab_frame = ttk.Frame(self.scrollable_frame)
                    lab_frame.pack(fill=tk.X, pady=(5, 2))
                    ttk.Label(lab_frame, text=f"实验室编号: {lab_no}", font=('Arial', 9, 'bold')).pack(side=tk.LEFT)

                    for i, project in enumerate(lab_projects):
                        project_name = project.get('name', '')
                        standard_no = project.get('standardNo', '')
                        project_id = project.get('projectId')

                        # 创建选择变量
                        var = tk.BooleanVar()

                        # 如果这个项目之前被选中过，保持选中状态
                        if project_id in selected_project_ids:
                            var.set(True)

                        self.project_vars.append(var)
                        self.project_id_to_var[project_id] = var

                        # 创建选择框
                        frame = ttk.Frame(self.scrollable_frame)
                        frame.pack(fill=tk.X, pady=1, padx=(20, 0))

                        checkbox = ttk.Checkbutton(frame, variable=var, text=f"{project_name} ({standard_no})")
                        checkbox.pack(side=tk.LEFT)
                        # 绑定选择状态变化事件
                        checkbox.bind('<Button-1>', lambda e, v=var: self.update_selected_count())

                # 获取动态列配置
                sample_project_ids = ",".join([str(project.get('projectId')) for project in projects])
                self.get_dynamic_columns(sample_project_ids)
            else:
                # 如果没有查询到，使用手动输入的信息
                self.sample_info = {
                    'sampleCode': full_sample_code,
                    'sampleName': f"手动输入样品-{full_sample_code}",
                    'detectionNo': detection_no,
                    'sampleSmallNo': sample_small_no,
                    'oldSampleProjectId': self.old_sample_project_id_var.get() if self.is_retest_var.get() else '',
                    'isRetest': self.is_retest_var.get(),
                    'sampleId': None,
                    'projectId': None
                }
                self.log("未查询到样品信息，使用手动输入数据")

            # 更新状态栏
            retest_status = " (注销复测)" if self.sample_info.get('isRetest', False) else ""
            sample_info_text = f"当前样品: {self.sample_info.get('sampleCode', '')}{retest_status}"
            self.sample_info_var.set(sample_info_text)

            # 更新选择数量显示
            self.update_selected_count()

            # 显示ID获取状态
            if self.sample_info.get('sampleId') and self.sample_info.get('projectId'):
                id_status = f"ID状态: 已获取真实ID (样品ID: {self.sample_info.get('sampleId')})"
            else:
                id_status = f"ID状态: 使用默认ID"
            self.log(id_status)

            self.log(f"已应用手动输入信息: {sample_info_text}")
            self.log(f"已选择 {sum(1 for var in self.project_vars if var.get())} 个项目")

            if self.is_retest_var.get():
                self.log(f"注销复测ID: {self.sample_info.get('oldSampleProjectId', '')}")
        else:
            messagebox.showwarning("警告", "请输入检测编号(detectionNo)")

    def clear_manual_input(self):
        """清空手动输入字段"""
        self.detection_no_var.set("TS25100593")
        self.sample_small_no_var.set("011")

        # 清空项目选择框
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.project_vars = []
        self.project_id_to_var = {}

        self.is_retest_var.set(False)
        self.old_sample_project_id_var.set("")
        self.old_sample_project_id_entry.config(state="disabled")

        # 清空状态栏
        self.sample_info_var.set("")
        self.project_method_var.set("")

        self.log("已清空手动输入信息")

    def generate_experiment_code(self):
        """生成统一的实验代码"""
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        user_prefix = self.login_system.current_user[:3].lower() if self.login_system.current_user else "unk"
        return f"{user_prefix}{timestamp}"

    def submit_single_data(self):
        """提交数据 - 使用相同的实验代码关联多个项目"""
        if not self.login_system.current_user:
            messagebox.showerror("错误", "请先登录系统")
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
            mass = float(self.data_fields["dynamic3582"].get() or "0")
            if mass <= 0:
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

        # 检查样品信息
        if not self.sample_info:
            messagebox.showerror("错误", "没有可用的样品信息，请先应用手动输入信息")
            return

        # 检查是否选择了项目
        selected_indices = [i for i, var in enumerate(self.project_vars) if var.get()]
        if not selected_indices:
            messagebox.showerror("错误", "请至少选择一个检测项目")
            return

        # 生成统一的实验代码
        experiment_code = self.generate_experiment_code()

        # 按实验室编号分组项目
        projects_by_lab = {}
        for index in selected_indices:
            if index < len(self.detected_projects):
                project = self.detected_projects[index]
                lab_no = project.get('laboratoryNo', '未知实验室')
                if lab_no not in projects_by_lab:
                    projects_by_lab[lab_no] = []
                projects_by_lab[lab_no].append(project)

        # 为每个实验室分组提交数据，使用相同的实验代码
        success_count = 0
        error_count = 0
        submission_results = []

        for lab_no, lab_projects in projects_by_lab.items():
            self.log(f"正在提交实验室 {lab_no} 的项目组，共 {len(lab_projects)} 个项目，实验代码: {experiment_code}")

            lab_success = 0
            lab_error = 0

            for project in lab_projects:
                project_name = project.get('name', '')
                standard_no = project.get('standardNo', '')
                project_id = project.get('projectId')

                self.log(f"正在提交项目: {project_name} (方法: {standard_no})")

                # 构建实验数据，使用相同的实验代码
                experiment_data = self.build_experiment_data(
                    self.sample_info.get('sampleCode', ''),
                    self.sample_info.get('sampleName', ''),
                    project_name,
                    standard_no,
                    project_id,
                    experiment_code  # 传递统一的实验代码
                )

                # 提交数据
                if self.submit_experiment_data(experiment_data, project_name):
                    lab_success += 1
                    success_count += 1
                else:
                    lab_error += 1
                    error_count += 1

            submission_results.append(f"实验室 {lab_no}: 成功 {lab_success}, 失败 {lab_error}")

        # 显示提交结果
        result_message = f"提交完成！\n总成功: {success_count} 个项目\n总失败: {error_count} 个项目\n实验代码: {experiment_code}\n\n"
        result_message += "\n".join(submission_results)

        if error_count == 0:
            messagebox.showinfo("成功", result_message)
            self.clear_single_data()

            # 提供查看结果的选项
            self.prompt_view_results_with_code(experiment_code)
        else:
            messagebox.showwarning("部分成功", result_message)

    def submit_experiment_data(self, experiment_data, project_name):
        """提交单个实验数据"""
        try:
            # 设置请求头，明确指定JSON格式
            headers = {
                'Content-Type': 'application/json; charset=utf-8',
                'Accept': 'application/json',
                'Referer': f'{self.login_system.base_url}/web/detectionResultCheckInCalc.html',
                'Origin': self.login_system.base_url
            }

            response = self.login_system.session.post(
                f"{self.login_system.base_url}/detectionManager/manager/ocExperiment/saveOcExperiment",
                json=experiment_data,
                headers=headers,
                verify=False,
                timeout=30  # 添加超时设置
            )

            self.log(f"HTTP状态码: {response.status_code}")

            if response.status_code == 200:
                try:
                    result = response.json()
                    self.log(f"服务器返回的完整响应: {json.dumps(result, ensure_ascii=False, indent=2)}")

                    if result.get('success'):
                        self.log(f"数据提交成功: {project_name}")
                        return True
                    else:
                        error_msg = result.get('errorCtx', {}).get('errorMsg', '未知错误')
                        error_code = result.get('errorCtx', {}).get('errorCode', '未知错误码')
                        self.log(f"数据提交失败 - 错误码: {error_code}, 错误信息: {error_msg}")
                        messagebox.showerror("错误",
                                             f"项目 {project_name} 提交失败: {error_msg} (错误码: {error_code})")
                        return False
                except json.JSONDecodeError:
                    self.log(f"服务器返回非JSON响应: {response.text}")
                    messagebox.showerror("错误", f"项目 {project_name} 提交失败: 服务器返回异常响应")
                    return False
            else:
                # 更详细的错误信息
                error_detail = response.text
                self.log(f"数据提交失败: HTTP {response.status_code}, 响应: {error_detail}")
                messagebox.showerror("错误", f"项目 {project_name} 提交失败: HTTP {response.status_code}")
                return False

        except requests.exceptions.Timeout:
            self.log("请求超时")
            messagebox.showerror("错误", f"项目 {project_name} 提交失败: 请求超时")
            return False
        except requests.exceptions.ConnectionError:
            self.log("连接错误")
            messagebox.showerror("错误", f"项目 {project_name} 提交失败: 网络连接错误")
            return False
        except Exception as e:
            self.log(f"数据提交失败: {str(e)}")
            messagebox.showerror("错误", f"项目 {project_name} 提交失败: {str(e)}")
            return False

    def build_experiment_data(self, sample_code, sample_name, project_name, standard_no, project_id,
                              experiment_code=None):
        """构建实验数据 - 使用统一的实验代码"""
        # 获取当前用户信息
        current_user = self.login_system.current_user
        user_info = self.login_system.users.get(current_user, {})
        pid = user_info.get('pid', '377')

        # 确保pid是数字类型
        try:
            pid_int = int(pid)
        except (ValueError, TypeError):
            pid_int = 377  # 默认值

        # 如果没有提供实验代码，生成一个
        if not experiment_code:
            experiment_code = self.generate_experiment_code()

        # 使用查询到的样品ID和项目ID，如果没有则使用默认值
        sample_id = self.sample_info.get('sampleId')

        # 如果ID为空，使用基于样品编号的默认ID
        if not sample_id:
            sample_id = f"sample_{sample_code}"
        if not project_id:
            project_id = f"project_{sample_code}"

        # 获取注销复测ID
        old_sample_project_id = self.sample_info.get('oldSampleProjectId', '')

        # 获取检测编号和样品小号
        detection_no = self.sample_info.get('detectionNo', self.detection_no_var.get())
        sample_small_no = self.sample_info.get('sampleSmallNo', self.sample_small_no_var.get())

        # 修复数据类型问题：确保所有数字字段都是正确的类型
        try:
            mass_value = float(self.data_fields["dynamic3582"].get() or "0")
        except ValueError:
            mass_value = 0

        # 根据项目名称确定正确的动态列配置
        dynamic_fields = self.get_project_dynamic_fields(project_name)

        # 从项目信息中获取实验室编号
        laboratory_no = "未知"
        for project in self.detected_projects:
            if project.get('projectId') == project_id:
                laboratory_no = project.get('laboratoryNo', '未知')
                break

        # 构建分析记录数据
        analysis_record = {
            "id": 1,
            "createDatetime": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "ocExperiment": None,
            "serialNumber": 1,
            "sampleId": int(sample_id) if sample_id and str(sample_id).isdigit() else 1,
            "sampleCode": sample_code,
            "sampleName": sample_name,
            "projectId": int(project_id) if project_id and str(project_id).isdigit() else 1,
            "projectCode": str(project_id),
            "projectName": project_name,
            "originalNumber": "1",
            "detectionLimit": "0.001",
            "standardValue": "0.001",
            "reportValue": "0.001",
            "reportUnit": 1,
            "reportUnitName": "%",
            "other": "",
            "spectrumId": None,
            "description": self.data_fields["dynamic3578"].get() or "试样",
            "calculatedValue": "0.001",
            "calculatedUnit": 1,
            "accuracy": "STANDARD_DEVIATION",
            "remark": "",
            "precisionCalcEnumMap": '{"PROCESS_STATUS_NEEDLESS":"不需要","STANDARD_DEVIATION":"标准偏差","RELATIVE_LABEL_DEVIATION":"相对标准偏差(>2)","RELATIVE_LABEL_DEVIATION_TWO":"相对标准偏差(≤2)","RELATIVE_DIFFERENCE":"相对相差(极差)","ABSOLUTE_DEVIATION":"绝对偏差","RELATIVE_DEVIATION":"相对偏差","ABSOLUTE_DIFFERENCE":"绝对差值"}',
            "ocAnalysisExperimentList": None,
            "ocCurve": None,
            "ocCurveSave": None,
            "ocCurveValueSaveList": None,
            "sampleProjectId": int(project_id) if project_id and str(project_id).isdigit() else 1,
            "detectionLimitType": "检出限",
            "originalNo": None,
            "standardAndWarningVal": None,
            "creatorName": current_user,
            "creatorId": pid_int,
            "modifierName": current_user,
            "modifyDatetime": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "dynamic3578": self.data_fields["dynamic3578"].get() or None,
            "dynamic3579": self.data_fields["dynamic3579"].get() or "<0.10",
            "dynamic3580": self.data_fields["dynamic3580"].get() or "<0.10",
            "dynamic3581": self.data_fields["dynamic3581"].get() or "25.0",
            "dynamic3582": str(mass_value) if mass_value > 0 else None,
            "dynamic3583": self.data_fields["dynamic3583"].get() or "1",
            "selectmap": dynamic_fields.get('selectmap', {}),
            "detectionNo": detection_no,
            "sampleSmallNo": sample_small_no,
            "detectionProjectName": project_name,
            "standardNo": standard_no,
            "laboratoryNo": laboratory_no,
            "oldSampleProjectId": int(old_sample_project_id) if old_sample_project_id and str(
                old_sample_project_id).isdigit() else None,
            "detectionNoAndSmallNo": sample_code,
            "samplePhotos": False,
            "sampleEvaluate": None,
            "sampleRemarkContent": "",
            "allSelectStatus": False,
            "picColorStatus": False,
            "cid": None,
            "_X_ID": "row_8"
        }

        # 基础模板 - 使用统一的实验代码
        template = {
            "sampleProjectIds2": str(project_id),
            "newSampleProjectIds": str(project_id),
            "experimentCode": experiment_code,  # 使用统一的实验代码
            "fileCode": sample_code,
            "referenceMaterial": self.reference_material_var.get() or None,
            "standardStrain": self.standard_strain_var.get() or None,
            "reagent": None,
            "cultureMedium": None,
            "temperature": float(self.temperature_var.get()),
            "humidity": float(self.humidity_var.get()),
            "detectionMethod": {
                "standardNo": standard_no,
                "id": 2431,
                "subMethodName": None
            },
            "computingFormula": "X=(C1-C0)*V*F/(m*10000)",
            "experimentProcess": None,
            "methodDescription": None,
            "weighingEquipment": {"name": ""},
            "weighingEquipmentJson": None,
            "cultivationEquipmentJson": "",
            "titrationEquipmentJson": "",
            "cultivationEquipment": "",
            "titrationEquipment": "",
            "reagentId": "",
            "titrationSolution": None,
            "userId": "",
            "description": None,
            "otherDescription": None,
            "mainEquipmentNames": "",
            "detectionMethodEquipmentBills": [],
            "mainEquipmentIds": "",
            "ocCurve": {
                "id": "",
                "equation": "",
                "correlationCoefficient": "",
                "slope": "",
                "intercept": ""
            },
            "roundMethod": "ROUND_METHOD_ENUM_2,3,ROUND_KEEP_ENUM_2",
            "roundMethodLevelJson": None,
            "calcMethod": "CALC_METHOD_ENUM_AVG",
            "ocCurveSaveList": "[]",
            "fileIds": "",
            "ocMethodSettings": {
                "id": 2438,
                "curveCalcMethod": "STANDARD_CALC",
                "curveXCalc": "X_CALC_AVG",
                "curveYCalc": "Y_CALC_AVG",
                "curveCalcXyReverse": None,
                "isCalculatedValue": 1,
                "sampleEvaluate": 0,
                "reportUnit": 1,
                "isOptionalTable": 0,
                "calculatedValueWidth": None,
                "calculatedUnitWidth": None,
                "reportValueWidth": None,
                "reportUnitWidth": None,
                "isRepeatabilityLimit": 0,
                "associatedRecords": 0
            },
            "spectrumJsonList": "",
            "recoverySpectrumJsonList": "",
            "consumableReceiveIds": None,
            "titrationConsumableReceiveIds": None,
            "emptyColonyCountList": None,
            "compareColonyCountList": None,
            "selfColonyCountList": None,
            "emptySampleColonyCountList": None,
            "negativeColonyCountList": None,
            "keepSampleColonyCountList": None,
            "curveCheckColonyCountList": None,
            "jcxTestColonyCountList": None,
            "materialColonyCountList": None,
            "methodCheckCode": None,
            "resultRoundMethod": None,
            "resultRoundMethodLevelJson": None,
            "mainEquipment": "[]",
            "dynamicColumns": json.dumps(dynamic_fields.get('columns', []), ensure_ascii=False),
            "startTime": self.start_date_var.get() + " 00:00:00",
            "endTime": self.end_date_var.get() + " 00:00:00",
            "KeyTime": "",
            "ocAnalysisRecordSaveList": json.dumps([analysis_record], ensure_ascii=False),
            "configureJsonList": "[]",
            "configureDJsonList": "[]",
            "configureSJsonList": "[]",
            "selfRecoverySaveList": "[]",
            "emptyRecoverySaveList": "[]",
            "negativeRecoverySaveList": "[]",
            "npecompareRecoverySaveList": "[]",
            "materialRecoverySaveList": "[]",
            "curveCheckRecoverySaveList": "[]",
            "emptyTestRecoverySaveList": "[]",
            "keepSampleRecoverySaveList": "[]",
            "jcxTestRecoverySaveList": "[]",
            "oldSampleProjectId": int(old_sample_project_id) if old_sample_project_id and str(
                old_sample_project_id).isdigit() else None,
            "detectionProjectName": project_name,
            "pid": pid_int,
            "pname": current_user,
            "loginId": pid_int
        }

        return template

    def get_project_dynamic_fields(self, project_name):
        """根据项目名称获取对应的动态字段配置"""
        # 这里需要根据实际项目配置不同的动态列
        # 目前使用相同的配置，但可以扩展为项目特定的配置

        base_selectmap = {
            "dynamic3579": [
                {
                    "id": 1462,
                    "createDatetime": "2023-12-07 11:44:29",
                    "methodSettingsId": 2438,
                    "methodTitleSettingId": 3579,
                    "sort": 1,
                    "isDefault": 1,
                    "selectValueName": "<0.10",
                    "creatorName": "系统管理员",
                    "creatorId": 1,
                    "modifierName": "系统管理员",
                    "modifyDatetime": "2023-12-07 11:44:29"
                }
            ],
            "dynamic3583": [
                {
                    "id": 1465,
                    "createDatetime": "2023-12-07 11:44:29",
                    "methodSettingsId": 2438,
                    "methodTitleSettingId": 3583,
                    "sort": 1,
                    "isDefault": 1,
                    "selectValueName": "1",
                    "creatorName": "系统管理员",
                    "creatorId": 1,
                    "modifierName": "系统管理员",
                    "modifyDatetime": "2023-12-07 11:44:29"
                }
            ],
            "dynamic3580": [
                {
                    "id": 1463,
                    "createDatetime": "2023-12-07 11:44:29",
                    "methodSettingsId": 2438,
                    "methodTitleSettingId": 3580,
                    "sort": 1,
                    "isDefault": 1,
                    "selectValueName": "<0.10",
                    "creatorName": "系统管理员",
                    "creatorId": 1,
                    "modifierName": "系统管理员",
                    "modifyDatetime": "2023-12-07 11:44:29"
                }
            ],
            "dynamic3581": [
                {
                    "id": 1464,
                    "createDatetime": "2023-12-07 11:44:29",
                    "methodSettingsId": 2438,
                    "methodTitleSettingId": 3581,
                    "sort": 1,
                    "isDefault": 1,
                    "selectValueName": "25.0",
                    "creatorName": "系统管理员",
                    "creatorId": 1,
                    "modifierName": "系统管理员",
                    "modifyDatetime": "2023-12-07 11:44:29"
                }
            ]
        }

        return {
            'selectmap': base_selectmap,
            'columns': self.dynamic_columns  # 使用获取的动态列配置
        }

    def prompt_view_results_with_code(self, experiment_code):
        """提示用户按实验代码查看结果"""
        response = messagebox.askyesno(
            "查看结果",
            f"数据提交成功！是否要查看分析记录？\n\n实验代码: {experiment_code}"
        )

        if response:
            # 构建查看结果的URL
            base_url = self.login_system.base_url

            # 这里需要根据实际系统的URL结构进行调整
            result_url = f"{base_url}/web/detectionResultCheckInListMgt.html"

            try:
                webbrowser.open(result_url)
                self.log(f"已打开结果查看页面: {result_url}")
                self.log(f"请在结果页面中使用实验代码筛选: {experiment_code}")
            except Exception as e:
                self.log(f"打开结果页面失败: {str(e)}")
                messagebox.showinfo(
                    "查看结果",
                    f"请在浏览器中手动访问:\n{result_url}\n\n查看时请使用实验代码筛选:\n实验代码: {experiment_code}"
                )

    def clear_single_data(self):
        """清空单条数据（保留样品信息）"""
        # 清空环境条件
        self.temperature_var.set("25")
        self.humidity_var.set("50")
        self.set_current_date()  # 重置日期为当前日期

        # 清空标准物质信息
        self.reference_material_var.set("")
        self.standard_strain_var.set("")

        # 重置字段值
        self.data_fields["dynamic3578"].set("")
        self.data_fields["dynamic3579"].set("<0.10")
        self.data_fields["dynamic3580"].set("<0.10")
        self.data_fields["dynamic3581"].set("25.0")
        self.data_fields["dynamic3582"].set("")
        self.data_fields["dynamic3583"].set("1")

        self.log("已清空检测数据")


def main():
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='检测数据录入系统')
    parser.add_argument('--result-checkin-ids', help='结果录入ID')
    parser.add_argument('--sample-id', help='样品ID')
    parser.add_argument('--sample-project-ids', help='样品项目ID')

    args = parser.parse_args()

    root = tk.Tk()
    app = DetectionEntrySystem(
        root,
        result_checkin_ids=args.result_checkin_ids,
        sample_id=args.sample_id,
        sample_project_ids=args.sample_project_ids
    )
    root.mainloop()


if __name__ == "__main__":
    main()