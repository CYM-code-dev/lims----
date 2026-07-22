# detection_entry_api.py
import requests
import json
import time
from datetime import datetime, timedelta
import urllib3
from urllib3.exceptions import InsecureRequestWarning
import pandas as pd
import os
from openpyxl import load_workbook

# 禁用SSL警告
urllib3.disable_warnings(InsecureRequestWarning)

class DetectionAPI:
    def __init__(self, login_system):
        self.login_system = login_system
        # 缓存配置数据
        self.cached_units = None
        self.cached_round_methods = None
        self.cached_calc_methods = None
        self.current_sample_id = None
        self.sub_method_map = {}
        self.method_id_to_standard_no = {}
        self.cached_solution_types = None

        # 实验编号缓存
        self.experiment_code_cache = {}  # 缓存已生成的实验编号
        self.user_prefix_cache = {}  # 缓存用户前缀

        self.load_sub_method_config()

    def generate_experiment_code(self, method_name="", log_func=None):
        """生成实验编号 - 统一实现，避免重复"""
        try:
            user_name = self.get_user_pname()

            # 构建缓存键
            cache_key = f"{user_name}_{method_name}" if method_name else user_name

            # 如果已经为该用户和方法生成过编号，直接返回
            if cache_key in self.experiment_code_cache:
                return self.experiment_code_cache[cache_key]

            # 生成新的实验编号
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")

            # 缓存用户前缀
            if user_name not in self.user_prefix_cache:
                user_prefix = user_name[:3].lower() if user_name else "unk"
                self.user_prefix_cache[user_name] = user_prefix
            else:
                user_prefix = self.user_prefix_cache[user_name]

            experiment_code = f"{user_prefix}{timestamp}"

            # 缓存生成的编号
            self.experiment_code_cache[cache_key] = experiment_code

            return experiment_code

        except Exception as e:
            if log_func:
                log_func(f"生成实验编号异常: {str(e)}")
            return f"exp_{int(time.time())}"

    def get_experiment_code_for_method(self, method_name, log_func=None):
        """获取特定方法对应的实验编号（如果已生成）"""
        try:
            user_name = self.get_user_pname()
            cache_key = f"{user_name}_{method_name}" if method_name else user_name
            code = self.experiment_code_cache.get(cache_key)

            if code and log_func:
                log_func(f"获取已存在的实验编号: {code} (方法: {method_name})")

            return code

        except Exception as e:
            if log_func:
                log_func(f"获取实验编号异常: {str(e)}")
            return None

    def clear_experiment_codes(self, user_name=None, method_name=None, log_func=None):
        """清理实验编号缓存"""
        try:
            if user_name is None:
                self.experiment_code_cache.clear()
                self.user_prefix_cache.clear()
            elif method_name is None:
                # 清理该用户的所有缓存
                keys_to_remove = [k for k in self.experiment_code_cache.keys() if k.startswith(f"{user_name}_")]
                for key in keys_to_remove:
                    del self.experiment_code_cache[key]
                if user_name in self.user_prefix_cache:
                    del self.user_prefix_cache[user_name]
            else:
                # 清理特定用户和方法的缓存
                cache_key = f"{user_name}_{method_name}"
                if cache_key in self.experiment_code_cache:
                    del self.experiment_code_cache[cache_key]

            if log_func:
                if user_name and method_name:
                    log_func(f"已清理实验编号缓存: {user_name} - {method_name}")
                elif user_name:
                    log_func(f"已清理用户 {user_name} 的所有实验编号缓存")
                else:
                    log_func("已清理所有实验编号缓存")

        except Exception as e:
            if log_func:
                log_func(f"清理实验编号缓存异常: {str(e)}")

    def get_user_pid(self):
        """动态获取当前用户的pid"""
        if not self.login_system.current_user:
            return '377'  # 默认值

        user_info = self.login_system.users.get(self.login_system.current_user, {})
        return user_info.get('pid', '377')

    def get_user_login_id(self):
        """动态获取当前用户的loginId"""
        return self.get_user_pid()

    def get_user_pname(self):
        """获取当前用户名"""
        return self.login_system.current_user or ""

    def load_sub_method_config(self):
        """加载子方法切换配置"""
        try:
            # 假设Excel文件在当前目录
            excel_file = "原始记录登记.xlsx"
            if os.path.exists(excel_file):
                # 使用openpyxl读取Excel
                wb = load_workbook(excel_file)
                if '子方法切换' in wb.sheetnames:
                    sheet = wb['子方法切换']
                    # 从第二行开始读取（跳过标题行）
                    for row in sheet.iter_rows(min_row=2, values_only=True):
                        if row[0] and row[1]:  # 确保主方法ID和子方法ID都不为空
                            main_method_id = str(row[0])
                            sub_method_id = str(row[1])
                            description = row[2] if len(row) > 2 else ""
                            self.sub_method_map[main_method_id] = {
                                'sub_method_id': sub_method_id,
                                'description': description
                            }
        except Exception as e:
            pass

    def get_solution_configure_id_dynamic(self, configure_order, log_func=None):
        """动态查询配置序号对应的配置ID - 严格检查审核状态"""
        try:
            params = {
                "_search": "false",
                "nd": int(time.time() * 1000),
                "pageSize": "100",
                "pageNo": "1",
                "sidx": "",
                "sord": "asc",
                "state": "0",
                "keyword": "",
                "type": "SOLUTION_TYPE_D",
                "pid": self.get_user_pid(),
                "pname": self.get_user_pname(),
                "loginId": self.get_user_login_id()
            }

            headers = {
                'Accept': 'application/json, text/javascript, */*; q=0.01',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36',
                'Referer': f"{self.login_system.base_url}/web/detectionResultCheckInCalc.html",
                'X-Requested-With': 'XMLHttpRequest'
            }

            response = self.login_system.session.get(
                f"{self.login_system.base_url}/detectionManager/manager/ocExperiment/getSolutionAdata",
                params=params,
                headers=headers,
                verify=False,
                timeout=30
            )

            if response.status_code == 200:
                result = response.json()
                if result.get('success'):
                    vo_list = result.get('resultData', {}).get('voList', [])

                    # 查找匹配的配置序号
                    for solution in vo_list:
                        solution_order = solution.get('configureOrder')
                        if solution_order == configure_order:
                            configure_id = solution.get('id')
                            audit_user_name = solution.get('auditUserName')

                            # 严格的审核状态检查
                            is_audited = audit_user_name is not None and audit_user_name != ""

                            # 如果未审核，直接返回错误
                            if not is_audited:
                                return None, "未审核"

                            if configure_id:
                                return str(configure_id), None

                    return None, "未找到配置序号"
                else:
                    error_msg = result.get('errorCtx', {}).get('errorMsg', '未知错误')
                    return None, error_msg
            else:
                return None, f"HTTP {response.status_code}"

        except Exception as e:
            if log_func:
                log_func(f"动态查询配置序号异常: {str(e)}")
            return None, str(e)


    def get_solution_configure_id(self, configure_order, log_func=None):
        """根据配置序号查询配置ID - 严格审核检查版本，完全阻止未审核提交"""
        try:
            # 只使用动态查询（包含严格审核检查）
            configure_id, error_msg = self.get_solution_configure_id_dynamic(configure_order, log_func)

            # 如果检测到未审核错误，立即返回
            if error_msg == "未审核":
                if log_func:
                    log_func(f"配置序号 {configure_order} 未审核")
                return None, error_msg

            # 如果动态查询成功返回配置ID，返回配置ID
            if configure_id:
                return configure_id, None

            # 其他所有情况都阻止提交
            return None, error_msg

        except Exception as e:
            if log_func:
                log_func(f"查询配置序号异常: {str(e)}")
            return None, str(e)

    def submit_solution_with_experiment(self, sample_project_ids, configure_order, experiment_codes, log_func=None):
        """提交标准溶液配置 - 严格审核检查，未审核时完全阻止提交"""
        try:
            # 确保sample_project_ids是逗号分隔的字符串
            if isinstance(sample_project_ids, list):
                sample_project_ids_str = ",".join(str(pid) for pid in sample_project_ids)
            else:
                sample_project_ids_str = str(sample_project_ids)

            # 动态查询配置ID并严格检查审核状态
            configure_id, error_msg = self.get_solution_configure_id(configure_order, log_func)

            # 严格的审核状态检查 - 只有明确返回配置ID才允许提交
            if not configure_id:
                return False, f"标准溶液 {configure_order} {error_msg}，请先审核后再提交"

            # 获取当前用户信息
            pid = self.get_user_pid()
            pname = self.get_user_pname()
            login_id = self.get_user_login_id()

            # 获取实验编号
            first_project_id = sample_project_ids[0] if isinstance(sample_project_ids, list) else \
            sample_project_ids_str.split(',')[0]
            experiment_code = experiment_codes.get(first_project_id, "")

            if not experiment_code:
                return False, "无法获取实验编号"

            # 获取实验ID
            experiment_id = None
            try:
                config = self.get_experiment_config(
                    first_project_id,
                    "",
                    "",
                    "",
                    log_func
                )
                if config and 'id' in config:
                    experiment_id = config['id']
            except Exception:
                pass

            if not experiment_id:
                experiment_id = 0

            # 构建请求数据
            request_data = {
                "id": int(experiment_id) if experiment_id and str(experiment_id).isdigit() else 0,
                "configureJsonList": json.dumps([
                    {
                        "configureIds": str(configure_id),
                        "configureType": "SOLUTION_TYPE_D",
                        "status": 0
                    }
                ]),
                "pid": pid,
                "pname": pname,
                "loginId": login_id
            }

            # 构建Referer头
            referer_url = f"{self.login_system.base_url}/web/detectionResultCheckInCalc.html?ids=&decideProjectOrgIds=23&type=0&recordNumber=null&checkInStatus=CHECK_IN_STATUS_NO&verifyStatus=&auditStatus=&resultCheckInIds={sample_project_ids_str.split(',')[0] if sample_project_ids_str else ''}&souce=checkIn&pid={pid}&pname={pname}&loginId={login_id}"

            headers = {
                'Content-Type': 'application/json;charset=UTF-8',
                'Accept': 'application/json, text/javascript, */*; q=0.01',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36',
                'Origin': self.login_system.base_url,
                'Referer': referer_url,
                'Accept-Encoding': 'gzip, deflate',
                'Accept-Language': 'zh-CN,zh;q=0.9'
            }

            # 尝试不同的API端点
            endpoints = [
                "/detectionManager/manager/ocExperiment/saveSolution",
                "/detectionManager/manager/ocExperiment/saveSolutionL"
            ]

            for endpoint in endpoints:
                try:
                    response = self.login_system.session.post(
                        f"{self.login_system.base_url}{endpoint}",
                        json=request_data,
                        headers=headers,
                        verify=False,
                        timeout=30
                    )

                    if response.status_code == 200:
                        try:
                            result = response.json()
                            if result.get('success'):
                                return True, ""
                            else:
                                error_msg = result.get('errorCtx', {}).get('errorMsg', '未知错误')
                                continue
                        except json.JSONDecodeError:
                            continue
                    else:
                        continue

                except Exception:
                    continue

            return False, "所有标准溶液提交端点都失败"

        except Exception as e:
            if log_func:
                log_func(f"提交标准溶液异常: {str(e)}")
            return False, f"提交标准溶液异常: {str(e)}"

    def get_all_solution_types(self, log_func=None):
        """获取所有溶液类型配置 - 处理字典格式，返回键和值"""
        try:
            # 如果已缓存，直接返回
            if self.cached_solution_types:
                return self.cached_solution_types

            params = {
                "pid": self.get_user_pid(),
                "pname": self.get_user_pname(),
                "loginId": self.get_user_login_id(),
                "_": str(int(time.time() * 1000))
            }

            if log_func:
                log_func(f"请求溶液类型URL: /detectionManager/manager/dtSolutionConfigure/getAllSolutionType")
                log_func(f"请求参数: {params}")

            response = self.login_system.session.get(
                f"{self.login_system.base_url}/detectionManager/manager/dtSolutionConfigure/getAllSolutionType",
                params=params,
                headers={
                    'Accept': 'application/json, text/javascript, */*; q=0.01',
                    'Referer': f'{self.login_system.base_url}/web/detectionResultCheckInCalc.html',
                    'X-Requested-With': 'XMLHttpRequest'
                },
                verify=False,
                timeout=30
            )

            if log_func:
                log_func(f"溶液类型响应状态: {response.status_code}")
                log_func(f"溶液类型响应文本: {response.text}")

            if response.status_code == 200:
                result = response.json()
                if result.get('success'):
                    solution_types = result.get('resultData', {})

                    # 处理字典格式
                    if isinstance(solution_types, dict):
                        # 将字典的键和值都作为选项
                        type_options = []
                        for key, value in solution_types.items():
                            type_options.append(key)  # 添加键
                            type_options.append(value)  # 添加值
                        self.cached_solution_types = type_options
                        return type_options
                    elif isinstance(solution_types, list):
                        # 如果是列表，直接返回
                        self.cached_solution_types = solution_types
                        return solution_types
                    else:
                        if log_func:
                            log_func(f"未知的溶液类型数据格式: {type(solution_types)}")
                        return []
                else:
                    error_msg = result.get('errorCtx', {}).get('errorMsg', '未知错误')
                    if log_func:
                        log_func(f"获取溶液类型配置失败: {error_msg}")
                    return []
            else:
                if log_func:
                    log_func(f"获取溶液类型配置失败: HTTP {response.status_code}")
                return []

        except Exception as e:
            if log_func:
                log_func(f"获取溶液类型配置异常: {str(e)}")
            return []

    def ensure_sub_method_and_update(self, sample_project_ids, method_name, method_id, project_names, log_func=None):
        """统一处理子方法切换和更新 - 返回实际使用的方法ID和标准号"""
        try:
            actual_method_id = method_id
            actual_method_name = method_name

            # 检查子方法切换
            if method_id and str(method_id) in self.sub_method_map:
                sub_method_info = self.sub_method_map[str(method_id)]
                actual_method_id = sub_method_info['sub_method_id']
                description = sub_method_info['description']

                # 更新检测方法
                update_success = self.update_method(
                    sample_project_ids,
                    actual_method_id,
                    project_names,
                    log_func
                )

                if update_success:
                    if log_func:
                        log_func(f"已切换到子方法: {description}")
                else:
                    if log_func:
                        log_func(f"警告: 检测方法更新失败，可能无法正确切换子方法")

                # 获取子方法的标准号
                sub_method_standard_no = self.get_method_standard_no_by_id(actual_method_id, log_func)
                if sub_method_standard_no:
                    actual_method_name = sub_method_standard_no
                else:
                    # 如果无法获取子方法标准号，使用原始方法标准号
                    actual_method_name = method_name

            return actual_method_id, actual_method_name

        except Exception as e:
            if log_func:
                log_func(f"子方法切换异常: {str(e)}")
            return method_id, method_name  # 发生异常时返回原始值

    def get_method_standard_no_by_id(self, method_id, log_func=None):
        """根据方法ID获取方法标准号 - 修复异常处理"""
        try:
            # 如果已经缓存，直接返回
            if str(method_id) in self.method_id_to_standard_no:
                return self.method_id_to_standard_no[str(method_id)]

            # 尝试不同的API端点
            endpoints = [
                "/detectionManager/manager/detectionMethod/getObj",
                "/detectionManager/manager/detectionMethod/pageObj"
            ]

            for endpoint in endpoints:
                try:
                    params = {
                        "id": method_id,
                        "pid": self.get_user_pid(),
                        "pname": self.get_user_pname(),
                        "loginId": self.get_user_login_id(),
                        "_": str(int(time.time() * 1000))
                    }

                    # 对于pageObj端点，需要不同的参数
                    if "pageObj" in endpoint:
                        params = {
                            "pageSize": "100",
                            "pageNo": "1",
                            "id": method_id,
                            "pid": self.get_user_pid(),
                            "pname": self.get_user_pname(),
                            "loginId": self.get_user_login_id(),
                            "_": str(int(time.time() * 1000))
                        }

                    response = self.login_system.session.get(
                        f"{self.login_system.base_url}{endpoint}",
                        params=params,
                        headers={
                            'Accept': 'application/json, text/javascript, */*; q=0.01',
                            'Referer': f'{self.login_system.base_url}/web/detectionResultCheckInCalc.html',
                            'X-Requested-With': 'XMLHttpRequest'
                        },
                        verify=False,
                        timeout=30
                    )

                    if response.status_code == 200:
                        result = response.json()
                        if result.get('success'):
                            # 处理不同的响应格式
                            method_data = None
                            if 'resultData' in result:
                                method_data = result.get('resultData', {})

                                # 处理数组响应
                                if isinstance(method_data, list) and method_data:
                                    method_data = method_data[0]

                                # 处理分页响应
                                elif isinstance(method_data, dict) and 'voList' in method_data:
                                    vo_list = method_data.get('voList', [])
                                    if vo_list:
                                        method_data = vo_list[0]

                            if method_data:
                                standard_no = method_data.get('standardNo') or method_data.get('standard_no')
                                if standard_no:
                                    self.method_id_to_standard_no[str(method_id)] = standard_no
                                    return standard_no

                except Exception as e:
                    continue

            return None

        except Exception as e:
            return None

    def submit_experiment_data(self, experiment_data, project_name, log_func=None):
        """提交实验数据 - 修复版本，确保与前端请求一致"""
        try:
            if not experiment_data:
                return False

            # 构建正确的请求头 - 与前端完全一致
            sample_project_ids = experiment_data.get("sampleProjectIds", "")
            first_project_id = sample_project_ids.split(",")[0] if sample_project_ids else ""

            headers = {
                'Content-Type': 'application/json;charset=UTF-8',
                'Accept': 'application/json, text/javascript, */*; q=0.01',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36',
                'Origin': self.login_system.base_url,
                'Referer': f'{self.login_system.base_url}/web/detectionResultCheckInCalc.html?ids=&decideProjectOrgIds=23&type=0&recordNumber=null&checkInStatus=CHECK_IN_STATUS_NO&verifyStatus=&auditStatus=&resultCheckInIds={first_project_id}&souce=checkIn&pid={experiment_data.get("pid", self.get_user_pid())}&pname={experiment_data.get("pname", self.get_user_pname())}&loginId={experiment_data.get("loginId", self.get_user_login_id())}',
                'Accept-Encoding': 'gzip, deflate',
                'Accept-Language': 'zh-CN,zh;q=0.9'
            }

            response = self.login_system.session.post(
                f"{self.login_system.base_url}/detectionManager/manager/ocExperiment/saveOcExperiment",
                json=experiment_data,
                headers=headers,
                verify=False,
                timeout=30
            )

            if response.status_code == 200:
                try:
                    result = response.json()
                    if result.get('success'):
                        # 删除成功日志
                        return True
                    else:
                        error_msg = result.get('errorCtx', {}).get('errorMsg', '未知错误')
                        if log_func:
                            log_func(f"提交实验数据失败: {error_msg}, 项目: {project_name}")
                        return False
                except json.JSONDecodeError:
                    if log_func:
                        log_func(f"提交实验数据失败: 响应不是有效的JSON格式")
                    return False
            else:
                if log_func:
                    log_func(f"提交实验数据失败: HTTP {response.status_code}, 项目: {project_name}")
                return False

        except Exception as e:
            if log_func:
                log_func(f"提交实验数据异常: {str(e)}, 项目: {project_name}")
            return False

    def clear_experiment_cache(self, sample_project_ids, log_func=None):
        """清除实验暂存数据 - 完整修复，确保传递所有项目ID"""
        try:
            # 确保sample_project_ids是逗号分隔的字符串
            if isinstance(sample_project_ids, list):
                sample_project_ids_str = ",".join(str(pid) for pid in sample_project_ids)
            else:
                sample_project_ids_str = str(sample_project_ids)

            # 使用抓包获得的正确API端点
            endpoint = "/detectionManager/manager/ocExperiment/cancleOcExperiment"

            # 使用表单数据格式，与抓包请求一致
            request_data = {
                "sampleProjectIds": sample_project_ids_str,
                "pid": self.get_user_pid(),
                "pname": self.get_user_pname(),
                "loginId": self.get_user_login_id()
            }

            # 构建正确的Referer头 - 关键修复
            referer_params = "ids=&decideProjectOrgIds=23&type=0&recordNumber=null&checkInStatus=CHECK_IN_STATUS_NO&verifyStatus=&auditStatus="

            if sample_project_ids_str:
                referer_params += f"&resultCheckInIds={sample_project_ids_str.split(',')[0] if sample_project_ids_str else ''}"

            referer_params += "&souce=checkIn"

            headers = {
                'Accept': 'application/json, text/javascript, */*; q=0.01',
                'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                'Referer': f'{self.login_system.base_url}/web/detectionResultCheckInCalc.html?{referer_params}',
                'X-Requested-With': 'XMLHttpRequest',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36'
            }

            response = self.login_system.session.post(
                f"{self.login_system.base_url}{endpoint}",
                data=request_data,
                headers=headers,
                verify=False,
                timeout=30
            )

            if response.status_code == 200:
                result = response.json()
                if result.get('success'):
                    # 删除成功日志
                    return True
                else:
                    error_msg = result.get('errorCtx', {}).get('errorMsg', '未知错误')
                    if log_func:
                        log_func(f"清除实验暂存数据失败: {error_msg}, 项目ID: {sample_project_ids_str}")
                    return False
            else:
                if log_func:
                    log_func(f"清除实验暂存数据失败: HTTP {response.status_code}, 项目ID: {sample_project_ids_str}")
                return False

        except Exception as e:
            if log_func:
                log_func(f"清除实验暂存数据异常: {str(e)}, 项目ID: {sample_project_ids}")
            return False

    def get_units_config(self, log_func=None):
        """获取单位配置 - 使用正确的API地址"""
        try:
            # 如果已缓存，直接返回
            if self.cached_units:
                return self.cached_units

            # 使用正确的API地址和参数
            params = {
                "pageSize": "999",
                "pageNo": "1",
                "reportUnit": "YES",
                "pid": self.get_user_pid(),
                "pname": self.get_user_pname(),
                "loginId": self.get_user_login_id(),
                "_": str(int(time.time() * 1000))
            }

            response = self.login_system.session.get(
                f"{self.login_system.base_url}/detectionManager/manager/meteringUnit/pageObj",
                params=params,
                headers={
                    'Accept': 'application/json, text/javascript, */*; q=0.01',
                    'Referer': f'{self.login_system.base_url}/web/detectionResultCheckInCalc.html',
                    'X-Requested-With': 'XMLHttpRequest'
                },
                verify=False,
                timeout=30
            )

            if response.status_code == 200:
                result = response.json()
                if result.get('success'):
                    units_data = result.get('resultData', {}).get('voList', [])
                    # 缓存单位配置
                    self.cached_units = units_data
                    return units_data
                else:
                    raise Exception("获取单位配置失败，API返回success=false")
            else:
                raise Exception(f"获取单位配置失败: HTTP {response.status_code}")

        except Exception as e:
            raise

    def get_round_methods_config(self, log_func=None):
        """获取修约方法配置"""
        try:
            # 如果已缓存，直接返回
            if self.cached_round_methods:
                return self.cached_round_methods

            # 从系统枚举获取修约方法
            round_methods = {
                "ROUND_METHOD_ENUM_1": "保留{}位有效位数",
                "ROUND_METHOD_ENUM_2": "保留小数点后{}位",
                "ROUND_METHOD_ENUM_3": "保留整数以{}临界点",
                "ROUND_METHOD_ENUM_4": "科学计数法表示"
            }

            round_keep_methods = {
                "ROUND_KEEP_ENUM_1": "四舍五入",
                "ROUND_KEEP_ENUM_2": "四舍六入五留双",
                "ROUND_KEEP_ENUM_3": ">0 直接进位",
                "ROUND_KEEP_ENUM_4": "直接截取"
            }

            config = {
                "round_methods": round_methods,
                "round_keep_methods": round_keep_methods
            }

            # 缓存配置
            self.cached_round_methods = config
            return config

        except Exception as e:
            raise

    def get_calc_methods_config(self, log_func=None):
        """获取计算方法配置"""
        try:
            # 如果已缓存，直接返回
            if self.cached_calc_methods:
                return self.cached_calc_methods

            # 从系统枚举获取计算方法
            calc_methods = {
                "CALC_METHOD_ENUM_AVG": "按平均值方式",
                "CALC_METHOD_ENUM_MAX": "按最大值方式",
                "CALC_METHOD_ENUM_GATHER": "按汇总方式",
                "CALC_METHOD_ENUM_SUM": "按求和方式",
                "CALC_METHOD_ENUM_MEDIAN": "按中位值方式",
                "CALC_METHOD_ENUM_MPN": "MPN检索表",
                "CALC_METHOD_MNUM_JOIN": "按+获取表头",
                "CALC_METHOD_CHOICE_CONTENT": "按勾选内容"
            }

            # 缓存配置
            self.cached_calc_methods = calc_methods
            return calc_methods

        except Exception as e:
            raise

    def get_method_specific_dynamic_columns(self, sample_project_ids, method_standard_no, method_id=None,
                                            log_func=None):
        """获取特定检测方法的动态列配置 - 修复子方法切换问题"""
        if not self.login_system.current_user:
            raise Exception("用户未登录")

        try:
            # 保存原始参数用于回退
            original_method_id = method_id
            original_standard_no = method_standard_no

            # 如果有方法ID，检查是否需要子方法切换
            actual_method_id = method_id
            if method_id:
                actual_method_id = self.check_and_switch_sub_method(method_id, log_func)

            # 使用 Fiddler 抓包获得的完整参数
            request_data = {
                "sampleProjectIds": sample_project_ids,
                "experimentCode": "",
                "detectionMethod.standardNo": method_standard_no,
                "detectionMethod.id": actual_method_id if actual_method_id else "",
                "detectionMethod.subMethodName": "",
                "pid": self.get_user_pid(),
                "pname": self.get_user_pname(),
                "loginId": self.get_user_login_id(),
                "isCalculatedValue": "1",
                "sampleEvaluate": "0",
                "reportUnit": "1",
                "isOptionalTable": "0"
            }

            # 第一次尝试：使用当前参数获取动态列
            response = self.login_system.session.post(
                f"{self.login_system.base_url}/detectionManager/manager/ocExperiment/getDynamicColumns",
                data=request_data,
                headers={
                    'Accept': 'application/json, text/javascript, */*; q=0.01',
                    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                    'Referer': f'{self.login_system.base_url}/web/detectionResultCheckInCalc.html',
                    'X-Requested-With': 'XMLHttpRequest',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36'
                },
                verify=False,
                timeout=30
            )

            # 处理响应
            if response.status_code == 200:
                result = response.json()
                if result.get('success'):
                    dynamic_columns = result.get('resultData', [])
                    return dynamic_columns
                else:
                    error_msg = result.get('errorCtx', {}).get('errorMsg', '未知错误')
                    # 如果是下拉菜单未设置选项的错误，尝试回退策略
                    if "动态列为下拉菜单，但未设置下拉菜单的选项" in error_msg:
                        return self.fallback_dynamic_columns(
                            sample_project_ids, method_standard_no, method_id,
                            original_method_id, original_standard_no,
                            request_data, log_func
                        )
                    else:
                        # 其他错误返回空数组
                        return []
            else:
                # HTTP错误时也尝试回退
                return self.fallback_dynamic_columns(
                    sample_project_ids, method_standard_no, method_id,
                    original_method_id, original_standard_no,
                    request_data, log_func
                )

        except Exception as e:
            # 异常情况下返回空数组，避免整个流程中断
            return []

    def fallback_dynamic_columns(self, sample_project_ids, method_standard_no, method_id,
                                 original_method_id, original_standard_no, request_data, log_func):
        """回退策略：尝试不同的方法获取动态列"""

        # 策略1: 如果发生了子方法切换，回退到主方法
        if method_id and original_method_id and method_id != original_method_id:
            request_data["detectionMethod.id"] = original_method_id
            response = self.retry_dynamic_columns_request(request_data, log_func)
            if response is not None:
                return response

        # 策略2: 尝试不传递方法ID
        request_data["detectionMethod.id"] = ""
        response = self.retry_dynamic_columns_request(request_data, log_func)
        if response is not None:
            return response

        # 策略3: 尝试使用原始标准号和方法ID
        request_data["detectionMethod.standardNo"] = original_standard_no
        request_data["detectionMethod.id"] = original_method_id if original_method_id else ""
        response = self.retry_dynamic_columns_request(request_data, log_func)
        if response is not None:
            return response

        # 所有策略都失败，返回空数组
        return []

    def retry_dynamic_columns_request(self, request_data, log_func):
        """重试动态列请求"""
        try:
            response = self.login_system.session.post(
                f"{self.login_system.base_url}/detectionManager/manager/ocExperiment/getDynamicColumns",
                data=request_data,
                headers={
                    'Accept': 'application/json, text/javascript, */*; q=0.01',
                    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                    'Referer': f'{self.login_system.base_url}/web/detectionResultCheckInCalc.html',
                    'X-Requested-With': 'XMLHttpRequest',
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36'
                },
                verify=False,
                timeout=30
            )

            if response.status_code == 200:
                result = response.json()
                if result.get('success'):
                    dynamic_columns = result.get('resultData', [])
                    return dynamic_columns

            return None

        except Exception as e:
            return None

    def get_experiment_config(self, sample_project_ids, method_standard_no, result_checkin_ids=None, sample_id=None,
                              log_func=None, method_id=None):
        """获取实验配置 - 确保使用子方法配置"""
        try:
            # 验证参数
            if not sample_project_ids:
                raise Exception("项目ID列表为空，无法获取实验配置")

            # 检查并切换子方法
            actual_method_id = method_id
            if method_id:
                actual_method_id = self.check_and_switch_sub_method(method_id, log_func)

            # 使用完整的项目ID列表
            request_data = {
                "sampleProjectIds": sample_project_ids,
                "calcType": "0",
                "pid": self.get_user_pid(),
                "pname": self.get_user_pname(),
                "loginId": self.get_user_login_id(),
            }

            # 只有在有值的情况下才添加可选参数
            if result_checkin_ids:
                request_data["resultCheckInIds"] = result_checkin_ids
            if sample_id:
                request_data["sampleId"] = sample_id
            # 添加方法ID参数 - 使用子方法ID
            if actual_method_id:
                request_data["detectionMethod.id"] = actual_method_id

            # 构建Referer头
            referer_params = "ids=&decideProjectOrgIds=23&type=0&recordNumber=null&checkInStatus=CHECK_IN_STATUS_NO&verifyStatus=&auditStatus="

            if result_checkin_ids:
                referer_params += f"&resultCheckInIds={result_checkin_ids}"
            if sample_id:
                referer_params += f"&sampleId={sample_id}"

            referer_params += "&souce=checkIn"

            headers = {
                'Accept': 'application/json, text/javascript, */*; q=0.01',
                'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                'Referer': f'{self.login_system.base_url}/web/detectionResultCheckInCalc.html?{referer_params}',
                'X-Requested-With': 'XMLHttpRequest',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36'
            }

            response = self.login_system.session.post(
                f"{self.login_system.base_url}/detectionManager/manager/ocExperiment/getOcExperiment",
                data=request_data,
                headers=headers,
                verify=False,
                timeout=30
            )

            if response.status_code == 200:
                result = response.json()
                if result.get('success'):
                    config_data = result.get('resultData', {})
                    return config_data
                else:
                    return {}
            else:
                return {}

        except Exception as e:
            return {}

    def query_samples_by_conditions(self, sample_code=None, project_name=None, method_name=None, retest_checked=False,
                                    log_func=None, exact_match=False):
        """通过多个条件查询样品信息 - 支持精确匹配和模糊查询"""
        if not self.login_system.current_user:
            return []

        try:
            today = datetime.now()
            one_month_ago = today - timedelta(days=30)

            # 优化分页参数
            page_size = 500
            page_no = 1
            all_projects = []

            # 循环获取所有页面的数据
            while True:
                # 直接在API参数中使用所有条件
                params = {
                    "_search": "false",
                    "nd": int(time.time() * 1000),
                    "pageSize": page_size,
                    "pageNo": page_no,
                    "acceptStartDate": one_month_ago.strftime("%Y-%m-%d"),
                    "acceptEndDate": today.strftime("%Y-%m-%d"),
                    "checkInStatus": "CHECK_IN_STATUS_NO",
                    "decideProjectOrgId": "23",
                    "sampleStatus": "one",
                    "pid": self.get_user_pid(),
                    "pname": self.get_user_pname(),
                    "loginId": self.get_user_login_id(),
                }

                # 精确匹配模式：使用样品编号作为关键字
                if exact_match and sample_code:
                    params["keyword"] = sample_code
                else:
                    # 模糊查询模式：使用各个字段分别查询
                    if sample_code:
                        params["keyword"] = sample_code
                    if project_name:
                        params["decideProjectName"] = project_name
                    if method_name:
                        params["decideProjectMethodName"] = method_name

                if retest_checked:
                    params["cancelRetesting"] = "1"

                start_time = time.time()
                response = self.login_system.session.get(
                    f"{self.login_system.base_url}/detectionManager/manager/resultCheckIn/pagePCObjAndSample",
                    params=params,
                    headers={
                        'Accept': 'application/json, text/javascript, */*; q=0.01',
                        'Referer': f'{self.login_system.base_url}/web/detectionResultCheckInListMgt.html'
                    },
                    verify=False,
                    timeout=15
                )
                request_time = time.time() - start_time

                if response.status_code == 200:
                    result = response.json()
                    if result.get('success') and result.get('resultData', {}).get('voList'):
                        vo_list = result['resultData']['voList']
                        current_page_records = len(vo_list)

                        # 处理当前页的记录
                        for sample_data in vo_list:
                            detection_no = sample_data.get('detectionNo', '')
                            small_no = sample_data.get('smallNo', '')
                            current_sample_code = f"{detection_no}{small_no}"

                            # 判断是否为复测项目
                            old_sample_project_id = sample_data.get('oldSampleProjectId')
                            is_retest = old_sample_project_id is not None

                            project_data = {
                                'sampleCode': current_sample_code,
                                'sampleName': sample_data.get('sampleName', ''),
                                'projectName': sample_data.get('decideProjectName', ''),
                                'standardNo': sample_data.get('standardNo', ''),
                                'projectId': sample_data.get('id'),
                                'sampleId': sample_data.get('sampleId'),
                                'detectionNo': detection_no,
                                'sampleSmallNo': small_no,
                                'isRetest': is_retest,
                                'oldSampleProjectId': old_sample_project_id,
                                'checkInStatus': sample_data.get('checkInStatus', '')
                            }
                            all_projects.append(project_data)

                        # 精确匹配模式：找到完全匹配的样品后立即返回
                        if exact_match and sample_code:
                            matched_projects = [p for p in all_projects if p.get('sampleCode') == sample_code]
                            if matched_projects:
                                return matched_projects

                        # 检查是否还有更多数据
                        if current_page_records < page_size:
                            break
                        else:
                            page_no += 1
                    else:
                        break
                else:
                    break

            return all_projects

        except Exception as e:
            return []

    def get_all_configs(self, sample_project_ids, method_standard_no, result_checkin_ids=None, sample_id=None,
                        log_func=None, method_id=None, project_names=None):
        """获取所有配置信息 - 集成动态列处理"""
        if not self.login_system.current_user:
            raise Exception("用户未登录")

        # 参数验证
        if not sample_project_ids:
            raise Exception("项目ID列表为空")

        if not method_standard_no:
            raise Exception("检测方法标准号为空")

        try:
            # 统一处理子方法切换
            actual_method_id, actual_method_name = self.ensure_sub_method_and_update(
                sample_project_ids,
                method_standard_no,
                method_id,
                project_names or [],
                log_func
            )

            configs = {}

            # 1. 获取实验配置 - 使用实际的方法ID
            experiment_config = self.get_experiment_config(
                sample_project_ids,
                actual_method_name,
                result_checkin_ids,
                sample_id,
                log_func,
                actual_method_id
            )
            configs['experiment'] = experiment_config

            # 2. 获取设备信息 - 使用实际的方法ID
            if actual_method_id:
                self.current_sample_id = sample_id
                equipment_config = self.get_detection_equipment(actual_method_id, log_func)
                configs['equipment'] = equipment_config
            else:
                configs['equipment'] = self.get_empty_equipment()

            # 3. 获取单位配置
            configs['units'] = self.get_units_config(log_func)

            # 4. 获取修约方法配置
            configs['round_methods'] = self.get_round_methods_config(log_func)

            # 5. 获取计算方法配置
            configs['calc_methods'] = self.get_calc_methods_config(log_func)

            # 6. 获取动态列配置 - 使用实际的方法ID
            dynamic_columns = self.get_method_specific_dynamic_columns(
                sample_project_ids, actual_method_name, actual_method_id, log_func
            )

            # 7. 如果API返回的动态列为空，尝试从实验配置中提取
            if not dynamic_columns and experiment_config:
                dynamic_columns = self.extract_dynamic_columns_from_config(experiment_config, log_func)

            configs['dynamic_columns'] = dynamic_columns

            # 返回实际使用的方法信息
            configs['actual_method_id'] = actual_method_id
            configs['actual_method_name'] = actual_method_name

            return configs

        except Exception as e:
            raise

    def get_detection_equipment(self, detection_method_id, log_func=None):
        """获取检测方法对应的设备信息 - 修复参数格式问题"""
        try:
            # 检查并切换子方法 - 确保返回的是字符串ID
            original_method_id = detection_method_id
            detection_method_id = self.check_and_switch_sub_method(detection_method_id, log_func)

            # 确保detection_method_id是字符串，不是元组
            if isinstance(detection_method_id, tuple):
                detection_method_id = detection_method_id[0]

            # 构建请求参数 - 确保参数是字符串
            params = {
                "detectionMethodId": str(detection_method_id),
                "pid": self.get_user_pid(),
                "pname": self.get_user_pname(),
                "loginId": self.get_user_login_id(),
                "_": str(int(time.time() * 1000))
            }

            # 构建Referer
            sample_id_param = f"&sampleId={self.current_sample_id}" if self.current_sample_id else ""
            referer_url = f"{self.login_system.base_url}/web/detectionResultCheckInCalc.html?ids=&decideProjectOrgIds=23&type=0&recordNumber=null&checkInStatus=CHECK_IN_STATUS_NO&verifyStatus=&auditStatus=&resultCheckInIds={sample_id_param}&souce=checkIn"

            headers = {
                'Accept': 'application/json, text/javascript, */*; q=0.01',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36',
                'Referer': referer_url,
                'Accept-Encoding': 'gzip, deflate',
                'Accept-Language': 'zh-CN,zh;q=0.9',
                'X-Requested-With': 'XMLHttpRequest'
            }

            response = self.login_system.session.get(
                f"{self.login_system.base_url}/detectionManager/manager/detectionMethod/selectByDetectionMethodId",
                params=params,
                headers=headers,
                verify=False,
                timeout=30
            )

            if response.status_code == 200:
                result = response.json()
                if result.get('success'):
                    equipment_list = result.get('resultData', [])
                    equipment_data = self.process_equipment_list(equipment_list, log_func)
                    return equipment_data
                else:
                    return self.get_empty_equipment()
            else:
                return self.get_empty_equipment()

        except Exception as e:
            return self.get_empty_equipment()

    def process_equipment_list(self, equipment_list, log_func=None):
        """处理设备列表，提取主检设备和称样设备 - 修复设备信息格式，避免橙色背景"""
        main_equipments = []
        weighing_equipments = []
        main_equipment_names_list = []
        main_equipment_ids_list = []
        weighing_equipment_ids_list = []

        # 遍历设备列表，按用途分类，只选择默认设备
        for equipment in equipment_list:
            used_category = equipment.get('usedCategory', '')
            name = equipment.get('name', '')
            main_equipment_names_value = equipment.get('mainEquipmentNames', '')
            is_default = equipment.get('isDefault')
            equipment_id = equipment.get('id')
            equipment_bill_id = equipment.get('equipmentBillId')

            # 优先使用 equipmentBillId，如果不存在则使用 id
            effective_id = equipment_bill_id if equipment_bill_id else equipment_id

            # 只处理默认设备 (isDefault=1)
            if is_default != 1:
                continue

            if used_category == '检测设备':
                # 对于检测设备，优先使用 mainEquipmentNames，如果没有则使用 name
                display_name = main_equipment_names_value if main_equipment_names_value else name

                equipment_info = {
                    'name': name,
                    'display_name': display_name,
                    'mainEquipmentNames': main_equipment_names_value,
                    'isDefault': is_default,
                    'equipment': equipment,
                    'id': effective_id
                }
                main_equipments.append(equipment_info)

                # 记录显示名称和ID
                main_equipment_names_list.append(display_name)
                if effective_id:
                    main_equipment_ids_list.append(str(effective_id))

            elif used_category == '称样设备':
                # 对于称样设备，优先使用 mainEquipmentNames，如果没有则使用 name
                display_name = main_equipment_names_value if main_equipment_names_value else name

                equipment_info = {
                    'name': name,
                    'display_name': display_name,
                    'isDefault': is_default,
                    'equipment': equipment,
                    'id': effective_id
                }
                weighing_equipments.append(equipment_info)

                if effective_id:
                    weighing_equipment_ids_list.append(str(effective_id))

        # 选择主检设备 - 只选择默认设备
        if main_equipments:
            # 构建主检设备显示名称列表
            main_equipment_display_names = [eq['display_name'] for eq in main_equipments]

            # 用分号连接所有默认主检设备
            main_equipment = "; ".join(main_equipment_display_names)
            main_equipment_names = "; ".join(main_equipment_names_list)
            main_equipment_ids = ",".join(main_equipment_ids_list)
        else:
            main_equipment = ""
            main_equipment_names = ""
            main_equipment_ids = ""

        # 选择称样设备 - 只选择默认设备
        weighing_equipment_base_name = ""
        weighing_equipment_display_name = ""
        weighing_equipment_id = ""

        if weighing_equipments:
            # 只选择默认称样设备
            selected_weighing = weighing_equipments[0]
            weighing_equipment_name = selected_weighing['name']
            weighing_equipment_display_name = selected_weighing['display_name']
            weighing_equipment_id = selected_weighing['id']

            # 提取设备基本名称（去掉编码和日期信息）
            if ',' in weighing_equipment_display_name:
                parts = weighing_equipment_display_name.split(',')
                if len(parts) >= 2:
                    weighing_equipment_base_name = parts[1]
                else:
                    weighing_equipment_base_name = weighing_equipment_name
            else:
                weighing_equipment_base_name = weighing_equipment_name
        else:
            weighing_equipment_display_name = ""
            weighing_equipment_base_name = ""
            weighing_equipment_id = ""

        # 修复：返回与前端操作一致的设备信息格式
        return {
            'mainEquipment': main_equipment,
            'weighingEquipment': weighing_equipment_display_name,
            'weighingEquipmentBaseName': weighing_equipment_base_name,
            'weighingEquipmentId': weighing_equipment_id,
            'titrationEquipment': '',
            'cultivationEquipment': '',
            'mainEquipmentNames': main_equipment_names,
            'mainEquipmentIds': main_equipment_ids,
            'raw_data': equipment_list
        }

    def get_empty_equipment(self):
        """返回空的设备信息"""
        return {
            'mainEquipment': '',
            'weighingEquipment': '',
            'weighingEquipmentId': '',
            'titrationEquipment': '',
            'cultivationEquipment': '',
            'mainEquipmentNames': '',
            'mainEquipmentIds': '',
            'raw_data': []
        }

    def get_current_sample_id(self):
        """获取当前样品ID - 用于构建Referer"""
        return ""

    def get_method_name_by_id(self, method_id, log_func=None):
        """根据方法ID获取方法名称 - 修复API调用"""
        try:
            # 尝试使用不同的API端点
            params = {
                "id": method_id,
                "pid": self.get_user_pid(),
                "pname": self.get_user_pname(),
                "loginId": self.get_user_login_id(),
                "_": str(int(time.time() * 1000))
            }

            # 尝试第一个API端点
            response = self.login_system.session.get(
                f"{self.login_system.base_url}/detectionManager/manager/detectionMethod/getObj",
                params=params,
                headers={
                    'Accept': 'application/json, text/javascript, */*; q=0.01',
                    'Referer': f'{self.login_system.base_url}/web/detectionResultCheckInCalc.html',
                    'X-Requested-With': 'XMLHttpRequest'
                },
                verify=False,
                timeout=30
            )

            if response.status_code == 200:
                result = response.json()
                if result.get('success'):
                    method_data = result.get('resultData', {})
                    method_name = method_data.get('name', '')

                    if method_name:
                        return method_name

            # 如果第一个API失败，尝试第二个API端点
            # 使用项目名称和方法ID查询
            response2 = self.login_system.session.get(
                f"{self.login_system.base_url}/detectionManager/manager/detectionMethod/getObjByIdAndDetProjectName",
                params=params,
                headers={
                    'Accept': 'application/json, text/javascript, */*; q=0.01',
                    'Referer': f'{self.login_system.base_url}/web/detectionResultCheckInCalc.html',
                    'X-Requested-With': 'XMLHttpRequest'
                },
                verify=False,
                timeout=30
            )

            if response2.status_code == 200:
                result = response2.json()
                if result.get('success'):
                    method_list = result.get('resultData', [])
                    for method_data in method_list:
                        if str(method_data.get('decideProjectMethodId')) == str(method_id):
                            method_name = method_data.get('decideProjectMethodName', '')
                            if method_name:
                                return method_name

            return None

        except Exception as e:
            return None

    def check_and_switch_sub_method(self, method_id, log_func=None):
        """检查并切换子方法 - 修复返回格式问题"""
        if not method_id:
            return method_id

        method_id_str = str(method_id)
        if method_id_str in self.sub_method_map:
            sub_method_info = self.sub_method_map[method_id_str]
            new_method_id = sub_method_info['sub_method_id']
            return new_method_id

        return method_id

    def update_method(self, sample_project_ids, method_id, project_names=None, log_func=None):
        """更新检测方法 - 修复参数格式"""
        try:
            # 构建项目名称字符串 - 用反引号分隔重复的项目名称
            if project_names and len(project_names) > 0:
                # 如果只有一个项目，直接使用项目名称
                if len(project_names) == 1:
                    decide_project_names = project_names[0]
                else:
                    # 多个项目时，用反引号连接相同的项目名称
                    decide_project_names = "`".join([project_names[0]] * len(project_names))
            else:
                decide_project_names = ""

            # 根据抓包数据修正请求参数格式
            request_data = {
                "ids": sample_project_ids,
                "decideProjectMethodId": method_id,
                "decideProjectMethodName": "GCMS法",
                "decideProjectNames": decide_project_names,
                "pid": self.get_user_pid(),
                "pname": self.get_user_pname(),
                "loginId": self.get_user_login_id()
            }

            headers = {
                'Content-Type': 'application/json;charset=UTF-8',
                'Accept': 'application/json, text/javascript, */*; q=0.01',
                'Referer': f'{self.login_system.base_url}/web/detectionResultCheckInCalc.html',
                'X-Requested-With': 'XMLHttpRequest'
            }

            response = self.login_system.session.post(
                f"{self.login_system.base_url}/detectionManager/manager/resultCheckIn/updateMethod",
                json=request_data,
                headers=headers,
                verify=False,
                timeout=30
            )

            if response.status_code == 200:
                result = response.json()
                if result.get('success'):
                    return True
                else:
                    return False
            else:
                return False

        except Exception as e:
            return False

    def get_solution_full_info(self, configure_order, log_func=None):
        """获取标准物质的完整信息 - 简化版本"""
        try:
            # 使用动态查询API获取标准物质信息
            dynamic_params = {
                "_search": "false",
                "nd": int(time.time() * 1000),
                "pageSize": "100",
                "pageNo": "1",
                "sidx": "",
                "sord": "asc",
                "state": "0",
                "keyword": configure_order,
                "type": "SOLUTION_TYPE_D",
                "pid": self.get_user_pid(),
                "pname": self.get_user_pname(),
                "loginId": self.get_user_login_id()
            }

            headers = {
                'Accept': 'application/json, text/javascript, */*; q=0.01',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36',
                'Referer': f"{self.login_system.base_url}/web/detectionResultCheckInCalc.html",
                'X-Requested-With': 'XMLHttpRequest'
            }

            response = self.login_system.session.get(
                f"{self.login_system.base_url}/detectionManager/manager/ocExperiment/getSolutionAdata",
                params=dynamic_params,
                headers=headers,
                verify=False,
                timeout=30
            )

            if response.status_code == 200:
                result = response.json()
                if result.get('success'):
                    vo_list = result.get('resultData', {}).get('voList', [])
                    # 查找匹配的配置序号
                    for solution in vo_list:
                        if solution.get('configureOrder') == configure_order:
                            # 提取详细信息并构建完整信息字符串
                            solution_name = solution.get('solutionName', '')
                            solution_code = solution.get('solutionCode', '')
                            concentration = solution.get('concentration', '')
                            configure_date = solution.get('configureDate', '')
                            validity_date = solution.get('validityDate', '')

                            full_info = f"{configure_order},{solution_name},{solution_code},{concentration},{configure_date},{validity_date}"
                            return full_info

            # 如果查询失败，返回配置序号
            return configure_order

        except Exception as e:
            if log_func:
                log_func(f"获取标准物质完整信息异常: {str(e)}")
            return configure_order

    def extract_dynamic_columns_from_config(self, experiment_config, log_func=None):
        """从实验配置中提取动态列 - 替代硬编码版本"""
        try:
            dynamic_columns = []

            # 从分析记录中提取动态字段
            analysis_records = experiment_config.get('ocAnalysisRecordList', [])
            if analysis_records:
                # 取第一个记录作为参考
                first_record = analysis_records[0]

                # 查找所有动态字段
                dynamic_fields = {k: v for k, v in first_record.items() if k.startswith('dynamic')}

                for field_code, field_value in dynamic_fields.items():
                    # 从字段代码中提取ID
                    field_id = field_code.replace('dynamic', '')

                    # 尝试从配置中获取字段显示名称
                    display_name = self.get_dynamic_field_display_name_from_system(field_code, field_id, log_func)

                    column = {
                        'id': field_id,
                        'columeCode': field_code,
                        'columeName': display_name,
                        'editType': 'EDIT_TYPE_TEXT',
                        'defaultVal': str(field_value) if field_value else '',
                        'columeOrder': len(dynamic_columns) + 1
                    }
                    dynamic_columns.append(column)

            return dynamic_columns

        except Exception as e:
            if log_func:
                log_func(f"从配置提取动态列异常: {str(e)}")
            return []

    def get_dynamic_field_display_name_from_system(self, field_code, field_id, log_func=None):
        """从系统获取动态字段的显示名称 - 替代硬编码版本"""
        try:
            # 方法1: 尝试通过API获取字段配置
            display_name = self.get_dynamic_field_config(field_id, log_func)
            if display_name:
                return display_name

            # 方法2: 从本地缓存或配置文件中获取
            display_name = self.get_dynamic_field_from_local_config(field_code, log_func)
            if display_name:
                return display_name

            # 方法3: 使用智能推断（基于字段代码模式）
            display_name = self.infer_display_name_from_pattern(field_code, log_func)
            if display_name:
                return display_name

            # 最终回退：返回字段代码
            return field_code

        except Exception as e:
            if log_func:
                log_func(f"获取动态字段显示名称异常: {str(e)}")
            return field_code

    def get_dynamic_field_config(self, field_id, log_func=None):
        """通过API获取动态字段配置"""
        try:
            # 这里可以调用系统的字段配置API
            # 由于系统API可能没有直接提供这个接口，暂时返回None
            return None

        except Exception as e:
            return None

    def get_dynamic_field_from_local_config(self, field_code, log_func=None):
        """从本地配置文件获取动态字段显示名称"""
        try:
            # 可以在这里读取外部配置文件（JSON、YAML等）
            # 示例配置字典 - 可以移到外部文件
            field_mapping = {
                # 这里可以维护一个字段映射表，但不再是硬编码在业务逻辑中
                # 可以定期更新这个映射表
            }

            # 如果配置了映射，返回映射值
            if field_code in field_mapping:
                return field_mapping[field_code]

            return None

        except Exception as e:
            return None

    def infer_display_name_from_pattern(self, field_code, log_func=None):
        """基于字段代码模式智能推断显示名称"""
        try:
            # 常见的字段代码模式推断
            patterns = {
                'dynamic11224': '体积V（ml）',
                'dynamic11225': '样品空白C0（μg/mL）',
                'dynamic11226': '样品浓度C1（μg/mL）',
                'dynamic11227': '稀释因子F',
                'dynamic1692': '体积V（ml）',
                'dynamic1693': '样品空白C0（μg/mL）',
                'dynamic1694': '样品浓度C1（μg/mL）',
                'dynamic1695': '稀释因子F'
            }

            # 精确匹配
            if field_code in patterns:
                return patterns[field_code]

            # 模式匹配（例如：dynamic后跟数字）
            if field_code.startswith('dynamic'):
                # 可以根据字段ID的范围或模式进行推断
                field_id = field_code.replace('dynamic', '')
                if field_id.isdigit():
                    # 这里可以添加更复杂的推断逻辑
                    # 例如：特定范围的ID对应特定类型的字段
                    pass

            return None

        except Exception as e:
            return None


class _Box:
    # ponytail: tk-StringVar/Text 替身，build_grouped_experiment_data 只用 .get()/.set()。
    # remark_text.get('1.0','end-1c') 需接受两参数，故 get 忽略位置参数。仅供无头 host 使用。
    def __init__(self, v=""):
        self._v = v

    def get(self, a=None, b=None):
        return self._v

    def set(self, v):
        self._v = v


def build_grouped_experiment_data(host, projects, experiment_code, method_name):
    """构建分组实验数据 - 包含完整标准物质信息（共享实现，阶段0从 detection_entry_main 抽出）。

    host 需提供: login_system(.current_user/.users)、experiment_config、equipment_config、
    dynamic_columns、data_fields[columeCode]、remark_text、solution_type_var、
    actual_method_name/actual_method_id(getattr)、api.get_solution_full_info、log、
    temperature_var/humidity_var/start_date_var/end_date_var。
    手动录入传 DetectionEntrySystem 实例；序列无头传 HeadlessHost。
    """
    self = host  # ponytail: 复用原方法体，self 即 host，函数体零改动
    # 获取第一个项目的信息作为基础
    first_project = projects[0]

    # 获取当前用户信息
    current_user = self.login_system.current_user
    user_info = self.login_system.users.get(current_user, {})
    pid = user_info.get('pid', '377')

    try:
        pid_int = int(pid)
    except (ValueError, TypeError):
        pid_int = 377

    # 关键修复：正确构建样品项目ID字符串
    sample_project_ids_list = []
    for project in projects:
        project_id = project.get('projectId')
        if project_id:
            sample_project_ids_list.append(str(project_id))

    sample_project_ids_str = ",".join(sample_project_ids_list)

    # 从动态获取的实验配置中获取关键信息
    oc_method_settings = self.experiment_config.get('ocMethodSettings', {})
    computing_formula = self.experiment_config.get('computingFormula')
    round_method = self.experiment_config.get('roundMethod')
    round_method_level_json = self.experiment_config.get('roundMethodLevelJson')
    calc_method = self.experiment_config.get('calcMethod')
    experiment_process = self.experiment_config.get('experimentProcess')

    # 使用保存的实际方法名称和方法ID（可能经过子方法切换）
    actual_method_name = getattr(self, 'actual_method_name', method_name)
    actual_method_id = getattr(self, 'actual_method_id', oc_method_settings.get('methodId'))

    # 获取分析记录配置
    analysis_records_config = self.experiment_config.get('ocAnalysisRecordList', [])

    # 按项目ID分组分析记录
    project_analysis_records = {}
    for record in analysis_records_config:
        project_id = record.get('projectId')
        if project_id not in project_analysis_records:
            project_analysis_records[project_id] = []
        project_analysis_records[project_id].append(record)

    # 从设备配置中获取设备信息
    main_equipment = self.equipment_config.get('mainEquipment', '')
    main_equipment_names = self.equipment_config.get('mainEquipmentNames', '')
    main_equipment_ids = self.equipment_config.get('mainEquipmentIds', '')
    weighing_equipment = self.equipment_config.get('weighingEquipment', '')
    weighing_equipment_base_name = self.equipment_config.get('weighingEquipmentBaseName', '')
    weighing_equipment_id = self.equipment_config.get('weighingEquipmentId', '')

    # 修复设备信息格式 - 与前端保持一致
    weighing_equipment_obj = {
        "name": weighing_equipment_base_name or weighing_equipment,
        "id": weighing_equipment_id
    } if weighing_equipment_id else ""

    weighing_equipment_json = weighing_equipment

    # 构建主检设备数组格式 - 与前端保持一致
    main_equipment_array = []
    if main_equipment_ids:
        equipment_ids = main_equipment_ids.split(',')
        equipment_names = main_equipment.split('; ')

        for i, equipment_id in enumerate(equipment_ids):
            equipment_name = ""
            equipment_checkout_date = ""

            if i < len(equipment_names):
                full_name = equipment_names[i]
                # 解析设备名称和日期信息
                if ',' in full_name:
                    parts = full_name.split(',')
                    if len(parts) >= 2:
                        equipment_name = parts[1].strip()
                    if len(parts) >= 3:
                        equipment_checkout_date = parts[2].strip()
                else:
                    equipment_name = full_name

                # 构建与前端完全一致的设备对象结构
                equipment_info = {
                    "id": equipment_id,
                    "name": equipment_name,
                    "checkOutDate": equipment_checkout_date,
                    "condition": None,
                    "instrumentAttachId": None
                }
                main_equipment_array.append(equipment_info)

    # 检查关键配置是否存在
    if not computing_formula:
        raise Exception("实验配置中缺少计算公式")
    if not round_method:
        raise Exception("实验配置中缺少修约方法")
    if not calc_method:
        raise Exception("实验配置中缺少计算方法")
    if not experiment_process:
        raise Exception("实验配置中缺少实验过程")

    # 修复：将 calcMethod 从对象转换为字符串
    if isinstance(calc_method, dict):
        calc_method_str = calc_method.get('key', 'CALC_METHOD_ENUM_AVG')
    else:
        calc_method_str = str(calc_method)

    # 修复：确保 ocMethodSettings 中的枚举字段是字符串而不是对象
    if oc_method_settings and isinstance(oc_method_settings, dict):
        cleaned_oc_method_settings = {}
        for key, value in oc_method_settings.items():
            if key in ['id', 'curveCalcMethod', 'curveXCalc', 'curveYCalc', 'curveCalcXyReverse',
                       'isCalculatedValue', 'sampleEvaluate', 'reportUnit', 'isOptionalTable',
                       'calculatedValueWidth', 'calculatedUnitWidth', 'reportValueWidth',
                       'reportUnitWidth', 'isRepeatabilityLimit', 'associatedRecords']:
                if key in ['curveCalcMethod', 'curveXCalc', 'curveYCalc']:
                    if isinstance(value, dict) and 'key' in value:
                        cleaned_oc_method_settings[key] = value['key']
                    else:
                        cleaned_oc_method_settings[key] = value
                else:
                    cleaned_oc_method_settings[key] = value
        oc_method_settings = cleaned_oc_method_settings

    # 构建分析记录列表
    oc_analysis_record_save_list = []
    record_id_counter = 1

    # 获取固定备注字段的内容
    remark_content = ""
    if hasattr(self, 'remark_text'):
        remark_content = self.remark_text.get('1.0', 'end-1c').strip()

    for i, project in enumerate(projects):
        sample_id = project.get('sampleId')
        project_id = project.get('projectId')
        detection_no = project.get('detectionNo')
        sample_small_no = project.get('sampleSmallNo')
        sample_name = project.get('sampleName', '未知样品')

        # 检查必要字段
        if not sample_id:
            raise Exception(f"项目 {project.get('projectName')} 缺少样品ID")
        if not project_id:
            raise Exception(f"项目 {project.get('projectName')} 缺少项目ID")
        if not detection_no:
            raise Exception(f"项目 {project.get('projectName')} 缺少检测编号")
        if not sample_small_no:
            raise Exception(f"项目 {project.get('projectName')} 缺少样品小号")

        # 获取该项目对应的分析记录
        project_records = project_analysis_records.get(project_id, [])

        if not project_records:
            project_records = [{
                "serialNumber": 1,
                "reportUnitName": "mg/kg",
                "calculatedUnit": "mg/kg",
                "detectionLimit": "0.1",
                "standardValue": "≤50",
                "detectionLimitType": "检出限"
            }]

        # 修正 detectionNoAndSmallNo 格式 - 与前端保持一致
        detection_no_and_small_no = f"{detection_no}-{sample_small_no.zfill(3)}"

        # 为每个分析记录生成记录
        for record_index, record_config in enumerate(project_records):
            # 动态构建字段数据
            dynamic_fields = {}
            selectmap_data = {}

            # 使用用户输入的值覆盖默认值
            for col in self.dynamic_columns:
                col_id = col.get('id')
                col_code = col.get('columeCode', f'dynamic{col_id}')
                col_name = col.get('columeName', '')
                edit_type = col.get('editType', 'EDIT_TYPE_TEXT')
                default_val = col.get('defaultVal', '')

                # 从用户输入获取值，正确处理空值
                user_value = ""
                if hasattr(self, 'data_fields') and col_code in self.data_fields:
                    user_value = self.data_fields[col_code].get().strip()

                # 如果用户输入了值，使用用户输入的值；否则使用配置中的默认值
                if user_value:
                    dynamic_fields[col_code] = user_value
                else:
                    # 如果用户没有输入，使用配置记录中的值或默认值
                    config_value = record_config.get(col_code)
                    if config_value is not None:
                        dynamic_fields[col_code] = str(config_value) if config_value is not None else default_val
                    else:
                        dynamic_fields[col_code] = default_val

                # 构建selectmap数据 - 与前端保持一致
                if edit_type == 'EDIT_TYPE_SELECT' and dynamic_fields[col_code]:
                    selectmap_data[col_code] = [{
                        "id": col_id,
                        "createDatetime": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "methodSettingsId": oc_method_settings.get('id', ''),
                        "methodTitleSettingId": col_id,
                        "sort": 1,
                        "isDefault": 1,
                        "selectValueName": dynamic_fields[col_code],
                        "creatorName": "系统管理员",
                        "creatorId": 1,
                        "modifierName": "系统管理员",
                        "modifyDatetime": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }]

            # 从分析记录配置中获取关键参数
            detection_limit = record_config.get('detectionLimit', '0.1')
            standard_value = record_config.get('standardValue', '≤1.0')
            standard_and_warning_val = record_config.get('standardAndWarningVal', '≤1.0')
            detection_limit_type = record_config.get('detectionLimitType', '检出限')
            accuracy = record_config.get('accuracy', 'STANDARD_DEVIATION')
            sample_remark_content = remark_content
            report_unit_name = record_config.get('reportUnitName', 'mg/kg')
            calculated_unit = record_config.get('calculatedUnit', 'mg/kg')

            # 构建分析记录 - 与前端数据结构完全一致
            analysis_record = {
                "id": record_id_counter,
                "createDatetime": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "ocExperiment": None,
                "serialNumber": record_config.get('serialNumber', record_index + 1),
                "sampleId": sample_id,
                "sampleCode": detection_no + sample_small_no,
                "sampleName": sample_name,
                "projectId": project_id,
                "projectCode": None,
                "projectName": project.get('projectName', ''),
                "originalNumber": None,
                "detectionLimit": detection_limit,
                "standardValue": standard_value,
                "reportValue": None,
                "reportUnit": None,
                "reportUnitName": report_unit_name,
                "other": None,
                "spectrumId": None,
                "description": None,
                "calculatedValue": None,
                "calculatedUnit": calculated_unit,
                "accuracy": accuracy,
                "remark": remark_content,
                "precisionCalcEnumMap": "{\"PROCESS_STATUS_NEEDLESS\":\"不需要\",\"STANDARD_DEVIATION\":\"标准偏差\",\"RELATIVE_LABEL_DEVIATION\":\"相对标准偏差(>2)\",\"RELATIVE_LABEL_DEVIATION_TWO\":\"相对标准偏差(≥2)\",\"RELATIVE_DIFFERENCE\":\"相对相差(误差)\",\"ABSOLUTE_DEVIATION\":\"绝对偏差\",\"RELATIVE_DEVIATION\":\"相对偏差\",\"ABSOLUTE_DIFFERENCE\":\"绝对差值\"}",
                "ocAnalysisExperimentList": None,
                "ocCurve": None,
                "ocCurveSave": None,
                "ocCurveValueSaveList": None,
                "sampleProjectId": project_id,
                "detectionLimitType": detection_limit_type,
                "originalNo": None,
                "standardAndWarningVal": standard_and_warning_val,
                "creatorName": None,
                "creatorId": None,
                "modifierName": None,
                "modifyDatetime": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                # 动态字段
                **dynamic_fields,
                # 添加selectmap字段
                "selectmap": json.dumps(selectmap_data, ensure_ascii=False) if selectmap_data else "{}",
                "detectionNo": detection_no,
                "detectionNoAndSmallNo": detection_no_and_small_no,
                "samplePhotos": False,
                "sampleEvaluate": None,
                "sampleRemarkContent": sample_remark_content,
                "allSelectStatus": False,
                "picColorStatus": False,
                "cid": None,
                "_X_ID": f"row_{8 + record_id_counter}"
            }
            oc_analysis_record_save_list.append(analysis_record)
            record_id_counter += 1

    # 构建检测方法对象 - 与前端保持一致
    detection_method = {
        "standardNo": actual_method_name,
        "id": actual_method_id if actual_method_id else "",
        "subMethodName": None
    }

    # 获取标准物质信息
    reference_material = ""
    if hasattr(self, 'solution_type_var') and self.solution_type_var.get().strip():
        configure_order = self.solution_type_var.get().strip()
        # 获取完整的标准物质信息
        reference_material = self.api.get_solution_full_info(configure_order, self.log)

    # 构建基础实验数据 - 与前端数据结构完全一致
    experiment_data = {
        "sampleProjectIds2": "",
        "newSampleProjectIds": "",
        "experimentCode": experiment_code,
        "fileCode": first_project.get('detectionNo', 'TR25100009'),
        # 使用完整的标准物质信息
        "referenceMaterial": reference_material,
        "standardStrain": None,
        "reagent": None,
        "cultureMedium": None,
        "temperature": self.temperature_var.get() or "22",
        "humidity": self.humidity_var.get() or "55",
        "detectionMethod": detection_method,
        "computingFormula": computing_formula,
        "experimentProcess": experiment_process,
        "methodDescription": None,
        "weighingEquipment": weighing_equipment_obj,
        "weighingEquipmentJson": weighing_equipment_json,
        "cultivationEquipmentJson": "",
        "titrationEquipmentJson": "",
        "cultivationEquipment": "",
        "titrationEquipment": "",
        "reagentId": "",
        "titrationSolution": None,
        "userId": "",
        "description": None,
        "otherDescription": None,
        "mainEquipmentNames": main_equipment_names,
        "detectionMethodEquipmentBills": [],
        "mainEquipmentIds": main_equipment_ids,
        "ocCurve": {
            "id": "",
            "equation": "",
            "correlationCoefficient": "",
            "slope": "",
            "intercept": ""
        },
        "roundMethod": round_method,
        "roundMethodLevelJson": round_method_level_json,
        "calcMethod": calc_method_str,
        "ocCurveSaveList": "[]",
        "fileIds": "",
        "ocMethodSettings": oc_method_settings,
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
        "mainEquipment": json.dumps(main_equipment_array, ensure_ascii=False) if main_equipment_array else "[]",
        "dynamicColumns": json.dumps(self.dynamic_columns, ensure_ascii=False) if self.dynamic_columns else "[]",
        "startTime": self.start_date_var.get() + " 00:00:00",
        "endTime": self.end_date_var.get() + " 00:00:00",
        "KeyTime": "",
        "ocAnalysisRecordSaveList": json.dumps(oc_analysis_record_save_list, ensure_ascii=False),
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
        # 关键修复：使用所有项目的ID
        "sampleProjectIds": sample_project_ids_str,
        "pid": pid_int,
        "pname": current_user,
        "loginId": pid_int
    }

    return experiment_data