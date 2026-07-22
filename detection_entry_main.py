# detection_entry_main.py
import tkinter as tk
from tkinter import ttk, messagebox
import json
import time
from datetime import datetime, timedelta
import webbrowser
import argparse
import os
import sys

# 导入登录系统和API模块
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from login import MultiUserLoginSystem
from detection_entry_api import DetectionAPI, build_grouped_experiment_data


class DetectionEntrySystem:
    def __init__(self, root, result_checkin_ids=None, sample_id=None, sample_project_ids=None, url_params=None):
        self.root = root
        self.root.title("检测数据录入系统")
        self.root.geometry("800x700")
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

        # 创建界面
        self.setup_ui()

        # 检查登录状态
        self.check_login_status()

    def check_login_status(self):
        """检查登录状态"""
        if self.login_system.load_session() and self.login_system.verify_session():
            self.update_status(f"已登录: {self.login_system.current_user}", "green")
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

        # 样品信息标签页
        sample_frame = ttk.Frame(notebook, padding="10")
        notebook.add(sample_frame, text="样品信息")

        # 环境条件标签页
        env_frame = ttk.Frame(notebook, padding="10")
        notebook.add(env_frame, text="环境条件")

        # 检测数据标签页
        data_frame = ttk.Frame(notebook, padding="10")
        notebook.add(data_frame, text="检测数据")

        # 设置各个标签页
        self.setup_sample_tab(sample_frame)
        self.setup_env_tab(env_frame)
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

    def setup_sample_tab(self, parent):
        """设置样品信息标签页 - 4个输入框+模糊查询"""
        # 查询条件区域
        query_frame = ttk.LabelFrame(parent, text="查询条件", padding="10")
        query_frame.pack(fill=tk.X, pady=(0, 10))

        # 第一行 - 样品编号和项目
        row1_frame = ttk.Frame(query_frame)
        row1_frame.pack(fill=tk.X, pady=5)

        ttk.Label(row1_frame, text="样品编号:").pack(side=tk.LEFT)
        self.sample_code_var = tk.StringVar()
        sample_code_entry = ttk.Entry(row1_frame, textvariable=self.sample_code_var, width=20)
        sample_code_entry.pack(side=tk.LEFT, padx=(5, 20))

        ttk.Label(row1_frame, text="项目:").pack(side=tk.LEFT)
        self.project_name_var = tk.StringVar()
        project_name_entry = ttk.Entry(row1_frame, textvariable=self.project_name_var, width=20)
        project_name_entry.pack(side=tk.LEFT, padx=(5, 20))

        # 第二行 - 检测方法和注销复测
        row2_frame = ttk.Frame(query_frame)
        row2_frame.pack(fill=tk.X, pady=5)

        ttk.Label(row2_frame, text="检测方法:").pack(side=tk.LEFT)
        self.method_name_var = tk.StringVar()
        method_name_entry = ttk.Entry(row2_frame, textvariable=self.method_name_var, width=20)
        method_name_entry.pack(side=tk.LEFT, padx=(5, 20))

        # 注销复测改为复选框
        self.retest_var = tk.BooleanVar(value=False)
        retest_checkbox = ttk.Checkbutton(row2_frame, text="注销复测", variable=self.retest_var)
        retest_checkbox.pack(side=tk.LEFT, padx=(5, 0))

        # 查询按钮区域
        button_frame = ttk.Frame(query_frame)
        button_frame.pack(fill=tk.X, pady=10)

        ttk.Button(button_frame, text="查询", command=self.query_samples).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="清空条件", command=self.clear_query_conditions).pack(side=tk.LEFT, padx=(0, 10))

        # 查询结果统计
        self.result_count_var = tk.StringVar(value="查询结果: 0 条")
        ttk.Label(button_frame, textvariable=self.result_count_var).pack(side=tk.LEFT, padx=(20, 0))

        # 查询结果区域
        results_frame = ttk.LabelFrame(parent, text="查询结果", padding="10")
        results_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 创建滚动框架用于结果显示
        canvas = tk.Canvas(results_frame)
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

        ttk.Button(selection_frame, text="全选", command=self.select_all_projects).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(selection_frame, text="取消全选", command=self.deselect_all_projects).pack(side=tk.LEFT,
                                                                                              padx=(0, 10))

        # 新增：获取动态配置按钮
        ttk.Button(selection_frame, text="获取动态配置", command=self.get_dynamic_config).pack(side=tk.LEFT,
                                                                                               padx=(0, 10))

        # 已选择项目统计
        self.selected_count_var = tk.StringVar(value="已选择: 0 个项目")
        ttk.Label(selection_frame, textvariable=self.selected_count_var).pack(side=tk.LEFT, padx=(20, 0))

    def setup_data_tab(self, parent):
        """设置检测数据标签页 - 添加标准溶液提交功能"""
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

        # 标准溶液提交区域
        self.setup_solution_section(main_data_frame)

        # 按钮区域
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=10)

        ttk.Button(button_frame, text="提交数据", command=self.submit_data).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="清空数据", command=self.clear_data).pack(side=tk.LEFT, padx=(0, 10))

        # 添加分组信息显示
        self.group_info_var = tk.StringVar(value="")
        group_info_label = ttk.Label(button_frame, textvariable=self.group_info_var, foreground="blue")
        group_info_label.pack(side=tk.LEFT, padx=(20, 0))

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

        # 实验编号信息
        self.experiment_code_var = tk.StringVar(value="实验编号: 未生成")
        experiment_label = ttk.Label(solution_frame, textvariable=self.experiment_code_var, foreground="blue")
        experiment_label.pack(anchor=tk.W, pady=(0, 5))

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
                mass_value = float(self.data_fields[mass_field].get() or "0")
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

        # 构建结果消息
        result_message = f"批量提交完成！\n\n"
        result_message += f"总项目数: {len(selected_projects)}\n"
        result_message += f"方法组数: {len(method_groups)}\n\n"

        # 实验数据提交结果
        if success_groups:
            result_message += f"实验数据成功组 ({len(success_groups)}):\n"
            for group in success_groups:
                project_count = len(method_groups[group])
                result_message += f"  ✓ {group} ({project_count}个项目)\n"

        if failed_groups:
            result_message += f"\n实验数据失败组 ({len(failed_groups)}):\n"
            for group in failed_groups:
                project_count = len(method_groups[group])
                result_message += f"  ✗ {group} ({project_count}个项目)\n"

        # 标准溶液提交结果
        if configure_order:
            result_message += f"\n标准溶液提交 ({configure_order}):\n"
            if solution_success_groups:
                result_message += f"  成功组 ({len(solution_success_groups)}):\n"
                for group in solution_success_groups:
                    project_count = len(method_groups[group])
                    result_message += f"    ✓ {group} ({project_count}个项目)\n"

            if solution_unaudited_groups:
                result_message += f"  未审核组 ({len(solution_unaudited_groups)}):\n"
                for group in solution_unaudited_groups:
                    project_count = len(method_groups[group])
                    result_message += f"    ⚠ {group} ({project_count}个项目) - 物质未审核\n"

            if solution_failed_groups:
                result_message += f"  失败组 ({len(solution_failed_groups)}):\n"
                for group in solution_failed_groups:
                    project_count = len(method_groups[group])
                    result_message += f"    ✗ {group} ({project_count}个项目)\n"

            if not solution_success_groups and not solution_unaudited_groups and not solution_failed_groups:
                result_message += "  未自动提交（实验数据提交失败）\n"
        else:
            result_message += f"\n标准溶液: 未填写配置序号\n"

        # 显示结果
        if failed_groups or solution_failed_groups:
            messagebox.showwarning("部分成功", result_message)
        elif solution_unaudited_groups and not failed_groups and not solution_failed_groups:
            messagebox.showinfo("提交完成",
                                f"实验数据提交成功！\n\n"
                                f"标准溶液 {configure_order} 未审核，无法关联，请审核后操作。")
        else:
            messagebox.showinfo("成功", result_message)

            # 只有全部成功时才清空数据
            if not failed_groups:
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
            success = self.get_single_method_group_config(method_name, projects)
            if success:
                success_groups.append(method_name)
            else:
                failed_groups.append(method_name)

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

        if failed_groups:
            messagebox.showwarning("部分成功", result_message)
        else:
            messagebox.showinfo("成功", result_message)

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

            # 使用统一的方法获取配置
            all_configs = self.api.get_all_configs(
                sample_project_ids_str,
                method_name,
                "",  # result_checkin_ids
                sample_id,
                self.log,
                None,  # 方法ID将在API内部获取
                project_names  # 项目名称列表
            )

            if not all_configs:
                error_msg = "无法获取配置信息，请检查网络连接或系统状态"
                raise Exception(error_msg)

            # 提取各个配置
            self.experiment_config = all_configs.get('experiment', {})
            self.equipment_config = all_configs.get('equipment', {})
            self.units_config = all_configs.get('units', [])
            self.round_methods_config = all_configs.get('round_methods', {})
            self.calc_methods_config = all_configs.get('calc_methods', {})
            self.dynamic_columns = all_configs.get('dynamic_columns', [])

            # 从配置中获取实际使用的方法信息
            actual_method_name = all_configs.get('actual_method_name', method_name)
            actual_method_id = all_configs.get('actual_method_id')

            # 保存实际的方法名称和方法ID用于实验数据
            self.actual_method_name = actual_method_name
            self.actual_method_id = actual_method_id

            # 构建显示名称
            method_id = self.experiment_config.get('ocMethodSettings', {}).get('methodId')
            if method_id and str(method_id) in self.api.sub_method_map:
                sub_method_info = self.api.sub_method_map[str(method_id)]
                description = sub_method_info['description']
                display_method_name = f"{actual_method_name} {description}"
            else:
                display_method_name = actual_method_name

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
        """设置环境条件标签页"""
        # 环境条件区域
        env_frame = ttk.LabelFrame(parent, text="环境条件", padding="10")
        env_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

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

        ttk.Button(env_row3, text="设置当前日期", command=self.set_current_date).pack(side=tk.LEFT)

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
        """根据动态列配置设置数据字段 - 修复数据绑定问题"""
        # 保存当前用户输入的值，避免重新创建字段时丢失
        current_values = {}
        if hasattr(self, 'data_fields'):
            for col_code, var in self.data_fields.items():
                current_values[col_code] = var.get()

        # 清空现有字段
        for widget in self.dynamic_fields_frame.winfo_children():
            widget.destroy()

        self.data_fields = {}

        if not dynamic_columns:
            # 显示无配置信息
            info_label = ttk.Label(self.dynamic_fields_frame,
                                   text="未获取到动态列配置，请检查网络连接或系统配置",
                                   foreground="orange")
            info_label.pack(pady=20)

            self.setup_fixed_remark_field()
            return

        sorted_columns = sorted(dynamic_columns, key=lambda x: x.get('columeOrder', 0))

        # 只处理动态列
        for col in sorted_columns:
            col_id = col.get('id')
            col_name = col.get('columeName', '')
            col_code = col.get('columeCode', f'dynamic{col_id}')
            edit_type = col.get('editType', 'EDIT_TYPE_TEXT')
            default_val = col.get('defaultVal', '')

            frame = ttk.Frame(self.dynamic_fields_frame)
            frame.pack(fill=tk.X, pady=5)

            ttk.Label(frame, text=col_name, width=20).pack(side=tk.LEFT)

            # 优先使用用户当前输入的值，如果没有则使用默认值
            current_value = current_values.get(col_code, default_val)

            if edit_type == 'EDIT_TYPE_SELECT':
                # 选择框类型
                var = tk.StringVar(value=current_value)

                # 获取可选值
                select_values = self.get_select_values(col_id, col_code)
                if select_values:
                    # 有可选值，创建下拉框
                    combobox = ttk.Combobox(frame, textvariable=var, values=select_values, width=20)
                    combobox.pack(side=tk.LEFT, padx=5)
                    if current_value:
                        combobox.set(current_value)
                else:
                    # 没有可选值，使用文本框
                    entry = ttk.Entry(frame, textvariable=var, width=20)
                    entry.pack(side=tk.LEFT, padx=5)

                self.data_fields[col_code] = var
            else:
                # 文本类型
                var = tk.StringVar(value=current_value)
                entry = ttk.Entry(frame, textvariable=var, width=20)
                entry.pack(side=tk.LEFT, padx=5)
                self.data_fields[col_code] = var

        # 添加固定备注字段（在动态列之后）
        self.setup_fixed_remark_field()

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
        filtered_projects = self.apply_client_filters(filtered_projects, sample_code, project_name, method_name,
                                                      retest_checked)
        self.filtered_projects = filtered_projects

        # 显示查询结果
        self.display_search_results(filtered_projects)

        # 更新状态
        self.update_status(f"查询完成: 找到 {len(filtered_projects)} 条记录", "green")
        self.result_count_var.set(f"查询结果: {len(filtered_projects)} 条")

    def apply_client_filters(self, projects, sample_code, project_name, method_name, retest_checked):
        """在客户端应用额外的过滤条件（确保完全匹配）"""
        filtered = projects

        # 样品编号过滤（模糊匹配）- 只有当输入了样品编号时才应用
        if sample_code:
            filtered = [p for p in filtered if sample_code.lower() in p.get('sampleCode', '').lower()]

        # 项目名称过滤（模糊匹配）- 只有当输入了项目名称时才应用
        if project_name:
            # 使用更宽松的匹配方式，确保包含关键字的项目都能被找到
            filtered = [p for p in filtered if project_name.lower() in p.get('projectName', '').lower()]

        # 检测方法过滤（模糊匹配）- 只有当输入了检测方法时才应用
        if method_name:
            filtered = [p for p in filtered if method_name.lower() in p.get('standardNo', '').lower()]

        # 注销复测过滤 - 基于oldSampleProjectId字段 - 只有当复选框被选中时才应用
        if retest_checked:
            # 只显示 oldSampleProjectId 有值的记录（注销复测）
            filtered = [p for p in filtered if p.get('oldSampleProjectId') is not None]
            if not filtered:
                pass
        else:
            # 复选框未选中时，显示所有记录（包括注销复测和非注销复测）
            pass  # 不过滤

        return filtered

    def display_search_results(self, projects):
        """显示查询结果"""
        if not projects:
            # 显示无结果消息
            no_result_label = ttk.Label(self.search_results_frame, text="未找到符合条件的记录", foreground="gray")
            no_result_label.pack(pady=20)
            return

        # 创建表头
        header_frame = ttk.Frame(self.search_results_frame)
        header_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(header_frame, text="选择", width=8).pack(side=tk.LEFT)
        ttk.Label(header_frame, text="样品编号", width=15).pack(side=tk.LEFT)
        ttk.Label(header_frame, text="样品名称", width=20).pack(side=tk.LEFT)
        ttk.Label(header_frame, text="项目", width=25).pack(side=tk.LEFT)
        ttk.Label(header_frame, text="检测方法", width=20).pack(side=tk.LEFT)
        ttk.Label(header_frame, text="注销复测", width=10).pack(side=tk.LEFT)

        # 添加分隔线
        separator = ttk.Separator(self.search_results_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=5)

        # 显示每个项目
        for i, project in enumerate(projects):
            project_frame = ttk.Frame(self.search_results_frame)
            project_frame.pack(fill=tk.X, pady=2)

            # 选择框
            var = tk.BooleanVar()
            self.project_vars.append(var)
            self.project_id_to_var[project.get('projectId')] = var

            checkbox = ttk.Checkbutton(project_frame, variable=var, width=8)
            checkbox.pack(side=tk.LEFT)
            checkbox.bind('<Button-1>', lambda e, v=var: self.on_project_selected())

            # 样品编号
            ttk.Label(project_frame, text=project.get('sampleCode', ''), width=15).pack(side=tk.LEFT)

            # 样品名称
            ttk.Label(project_frame, text=project.get('sampleName', ''), width=20).pack(side=tk.LEFT)

            # 项目名称
            ttk.Label(project_frame, text=project.get('projectName', ''), width=25).pack(side=tk.LEFT)

            # 检测方法
            ttk.Label(project_frame, text=project.get('standardNo', ''), width=20).pack(side=tk.LEFT)

            # 注销复测
            retest_text = "是" if project.get('isRetest', False) else "否"
            ttk.Label(project_frame, text=retest_text, width=10).pack(side=tk.LEFT)

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
            field_value = var.get()
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

            # 构建实验数据 - 使用修复后的方法
            experiment_data = self.build_grouped_experiment_data(projects, experiment_code, actual_method_name)

            # 提交数据
            success = self.api.submit_experiment_data(experiment_data, actual_method_name, self.log)

            if success:
                # 保存实验编号
                if not hasattr(self, 'experiment_codes'):
                    self.experiment_codes = {}

                for project_id in project_ids:
                    self.experiment_codes[project_id] = experiment_code

                # 更新实验编号显示
                self.update_experiment_code_display()

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
                    self.data_fields[col_code].set(default_val)

        # 清空固定备注字段
        if hasattr(self, 'remark_text'):
            self.remark_text.delete('1.0', tk.END)

def main():
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