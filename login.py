import sys
import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as ttkb  # 档1: 现代主题(sandstone-light)，ttk 控件自动套用
import requests
import hashlib
import json
import time
import os
from PIL import Image, ImageTk, ImageDraw
from io import BytesIO
import urllib3
import threading
import pystray
from pystray import MenuItem as item
import datetime
import socket
import paths

# 禁用SSL警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


class MultiUserLoginSystem:
    def __init__(self, parent_window=None):
        self.parent_window = parent_window
        self.session = requests.Session()
        self.base_url = "http://192.168.12.234:60015"
        self.session.verify = False
        self.current_user = None
        self.current_pid = None  # 登录后由「获取用户详情」接口缓存，供会话校验/保活使用
        self.current_realname = None  # 登录后由「获取用户详情」接口缓存的真实姓名，供状态栏显示
        self.current_org_id = None  # 登录后由「获取用户详情」接口缓存的所属实验室 orgId(查询机构过滤用)

        # 使用智能路径查找
        self.users_file = os.path.join(paths.data_dir(), "users_config.json")
        self.session_file = os.path.join(paths.data_dir(), "session_info.json")

        # 会话保持相关属性
        self.keep_alive_thread = None
        self.keep_alive_flag = False
        self.keep_alive_interval = 3600  # 60分钟执行一次

        # 时间控制相关属性
        self.keep_alive_enabled = True  # 控制是否启用会话保持
        self.last_check_time = None  # 上次检查时间

        # 存储当前验证码的会话
        self.captcha_session = None

        # 初始化请求头
        self.update_headers()

        # 加载用户配置
        self.users = self.load_users()

    def _get_resource_path(self, filename):
        """获取资源文件的绝对路径，支持打包环境和开发环境"""
        # 其他文件使用原来的逻辑
        # 如果是打包环境
        if getattr(sys, 'frozen', False):
            # 打包后的环境：首先在exe同级目录查找
            base_dir = os.path.dirname(sys.executable)

            # 1. 首先在exe同级目录查找
            exe_dir_path = os.path.join(base_dir, filename)
            if os.path.exists(exe_dir_path):
                return exe_dir_path

            # 2. 在 _internal 目录查找（PyInstaller打包时可能的位置）
            internal_dir = os.path.join(base_dir, '_internal')
            if os.path.exists(internal_dir):
                internal_path = os.path.join(internal_dir, filename)
                if os.path.exists(internal_path):
                    return internal_path

            # 3. 在 MEIPASS 目录查找（PyInstaller临时解压目录）
            if hasattr(sys, '_MEIPASS'):
                meipass_path = os.path.join(sys._MEIPASS, filename)
                if os.path.exists(meipass_path):
                    return meipass_path

            # 4. 如果都没找到，返回exe同级目录的路径
            return exe_dir_path
        else:
            # 开发环境：在脚本所在目录查找
            base_dir = os.path.dirname(os.path.abspath(__file__))
            return os.path.join(base_dir, filename)

    def update_user_password(self, username, new_password_hash, new_pid=None):
        """更新用户密码和相关信息"""
        if username in self.users:
            # 更新密码哈希
            self.users[username]["password_hash"] = new_password_hash

            # 更新PID（如果提供）
            if new_pid is not None:
                self.users[username]["pid"] = new_pid

            # 保存更新后的用户配置
            self.save_users()

            # 如果当前登录用户是此用户，则清除会话，强制重新登录
            if self.current_user == username:
                self.logout()

            return True
        return False

    def force_refresh_session(self, username):
        """强制刷新指定用户的会话"""
        # 清除会话文件
        if os.path.exists(self.session_file):
            os.remove(self.session_file)

        # 清除cookies
        self.session.cookies.clear()

        # 重置当前用户
        if self.current_user == username:
            self.current_user = None

        # 停止会话保持
        self.stop_keep_alive()

    def update_headers(self):
        """更新请求头"""
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Safari/537.36",
            "Accept": "application/json, text/javascript, */*; q=0.01",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Origin": self.base_url,
            "Referer": f"{self.base_url}/web/login.html",
            "Connection": "keep-alive",
            "Host": "192.168.12.234:60015"
        }
        self.session.headers.update(self.headers)

    def load_users(self):
        """加载用户配置 - 修复版本，支持打包环境"""
        try:
            # 确保配置文件存在
            if not os.path.exists(self.users_file):
                print(f"用户配置文件不存在，创建默认文件: {self.users_file}")
                self.save_users({})
                return {}

            # 读取配置文件
            with open(self.users_file, 'r', encoding='utf-8') as f:
                users = json.load(f)
                return users

        except json.JSONDecodeError as e:
            print(f"解析用户配置文件失败: {e}")
            # 创建备份文件
            if os.path.exists(self.users_file):
                backup_file = self.users_file + ".backup"
                try:
                    import shutil
                    shutil.copy2(self.users_file, backup_file)
                    print(f"已创建配置文件备份: {backup_file}")
                except:
                    pass
            # 返回空字典并重置文件
            self.save_users({})
            return {}
        except Exception as e:
            print(f"加载用户配置文件失败: {e}")
            return {}

    def save_users(self, users=None):
        """保存用户配置 - 修复版本，支持打包环境"""
        if users is None:
            users = self.users

        try:
            # 确保目录存在
            file_dir = os.path.dirname(self.users_file)
            if file_dir and not os.path.exists(file_dir):
                os.makedirs(file_dir, exist_ok=True)
                print(f"创建目录: {file_dir}")

            # 保存文件
            with open(self.users_file, 'w', encoding='utf-8') as f:
                json.dump(users, f, ensure_ascii=False, indent=2)

            print(f"用户配置已保存到: {self.users_file}")
            return True
        except Exception as e:
            print(f"保存用户配置失败: {e}")
            # 尝试在其他位置保存
            try:
                # 尝试在当前工作目录保存
                fallback_path = os.path.join(paths.data_dir(), "users_config.json")
                with open(fallback_path, 'w', encoding='utf-8') as f:
                    json.dump(users, f, ensure_ascii=False, indent=2)
                print(f"用户配置已保存到备用位置: {fallback_path}")
                self.users_file = fallback_path  # 更新文件路径
                return True
            except Exception as e2:
                print(f"备用保存也失败: {e2}")
                return False

    def load_session(self):
        """尝试加载保存的会话 - 修复版本，支持打包环境"""
        try:
            if os.path.exists(self.session_file):
                with open(self.session_file, 'r', encoding='utf-8') as f:
                    session_info = json.load(f)

                # 检查会话是否过期（超过24小时）
                login_time_str = session_info.get("login_time")
                if login_time_str:
                    try:
                        login_time = time.strptime(login_time_str, "%Y-%m-%d %H:%M:%S")
                        login_timestamp = time.mktime(login_time)
                        current_timestamp = time.time()
                        # 如果会话超过24小时，视为过期
                        if current_timestamp - login_timestamp > 24 * 60 * 60:
                            print("会话已过期")
                            os.remove(self.session_file)
                            return False
                    except:
                        # 如果解析时间失败，也视为过期
                        os.remove(self.session_file)
                        return False

                # 恢复cookies
                cookies = session_info.get("cookies", {})
                for name, value in cookies.items():
                    self.session.cookies.set(name, value)

                # 恢复headers
                self.session.headers.update(session_info.get("headers", {}))
                # 恢复当前用户信息
                self.current_user = session_info.get("username")
                self.current_pid = session_info.get("pid")
                self.current_realname = session_info.get("realname")
                return True
            else:
                print(f"会话文件不存在: {self.session_file}")
                return False

        except Exception as e:
            print(f"❌ 加载会话失败: {e}")
            # 加载失败时删除损坏的会话文件
            if os.path.exists(self.session_file):
                try:
                    os.remove(self.session_file)
                except:
                    pass
            return False

    def save_session(self):
        """保存会话信息 - 修复版本，支持打包环境"""
        try:
            # 确保目录存在
            file_dir = os.path.dirname(self.session_file)
            if file_dir and not os.path.exists(file_dir):
                os.makedirs(file_dir, exist_ok=True)

            # 处理cookies，避免重复名称
            cookies_dict = {}
            for cookie in self.session.cookies:
                cookies_dict[cookie.name] = cookie.value

            session_info = {
                "username": self.current_user,
                "login_time": time.strftime("%Y-%m-%d %H:%M:%S"),
                "cookies": cookies_dict,
                "headers": dict(self.session.headers),
                "pid": self.current_pid,
                "realname": self.current_realname
            }

            with open(self.session_file, "w", encoding="utf-8") as f:
                json.dump(session_info, f, ensure_ascii=False, indent=2)

            print(f"会话信息已保存到: {self.session_file}")
            return True
        except Exception as e:
            print(f"❌ 保存会话信息失败: {e}")
            # 尝试备用位置
            try:
                fallback_path = os.path.join(paths.data_dir(), "session_info.json")
                with open(fallback_path, "w", encoding="utf-8") as f:
                    json.dump(session_info, f, ensure_ascii=False, indent=2)
                print(f"会话信息已保存到备用位置: {fallback_path}")
                self.session_file = fallback_path  # 更新文件路径
                return True
            except Exception as e2:
                print(f"备用保存也失败: {e2}")
                return False

    def verify_session(self):
        """验证会话是否仍然有效 - 通过实际搜索来验证"""
        if not self.current_user:
            return False

        try:
            # 使用登录时缓存的 pid
            user_pid = self.current_pid or ""

            # 尝试执行一个简单的搜索来验证会话
            test_params = {
                "_search": "false",
                "nd": str(int(time.time() * 1000)),
                "pageSize": 1,
                "pageNo": 1,
                "sidx": "",
                "sord": "asc",
                "acceptStartDate": time.strftime("%Y-%m-%d"),
                "acceptEndDate": time.strftime("%Y-%m-%d"),
                "retrievalWay": "decideProjectName",
                "decideProjectOrgId": "23",
                "checkInStatus": "CHECK_IN_STATUS_NO",
                "screenTime": "",
                "sort": "0",
                "keyword": "test",  # 使用更通用的关键词
                "pid": user_pid,
                "pname": self.current_user,
                "loginId": user_pid  # 使用用户的pid作为loginId
            }

            headers = {
                "Referer": f"{self.base_url}/web/detectionResultCheckInListMgt.html?menuId=41",
                "User-Agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Mobile Safari/537.36"
            }

            url = f"{self.base_url}/detectionManager/manager/resultCheckIn/pagePCObjAndSample"
            response = self.session.get(url, params=test_params, headers=headers, timeout=5)

            if response.status_code == 200:
                result = response.json()
                if result.get("success"):
                    return True
        except Exception as e:
            print(f"会话验证过程中出错: {e}")

        return False

    def ensure_login(self, username):
        """确保用户已登录，优先使用会话，失败则重新登录"""
        # 尝试加载已有会话
        if self.load_session():
            # 验证会话是否仍然有效
            if self.verify_session():
                # 检查当前会话用户是否与目标用户一致
                if self.current_user == username:
                    return True, "会话有效，无需重新登录"
                else:
                    # 用户不一致，需要重新登录
                    self.session.cookies.clear()
                    if os.path.exists(self.session_file):
                        os.remove(self.session_file)
            else:
                # 会话无效，需要重新登录
                self.session.cookies.clear()
                if os.path.exists(self.session_file):
                    os.remove(self.session_file)

        # 需要重新登录
        return False, "需要重新登录"

    def get_captcha_image(self):
        """获取验证码图片 - 修复版本，确保每次都是新的会话"""
        try:
            # 创建一个新的临时会话来获取验证码
            temp_session = requests.Session()
            temp_session.verify = False

            # 复制主会话的headers
            temp_session.headers.update(self.headers)

            timestamp = str(int(time.time() * 1000))
            captcha_url = f"{self.base_url}/detectionManager/core/security/validatecodes?{timestamp}"

            # 添加随机参数避免缓存
            captcha_url += f"&r={timestamp}"

            response = temp_session.get(captcha_url)
            if response.status_code == 200:
                # 保存验证码会话，用于后续登录
                self.captcha_session = temp_session
                return Image.open(BytesIO(response.content))
            else:
                print(f"获取验证码失败，状态码: {response.status_code}")
                return None
        except Exception as e:
            print(f"获取验证码时出错: {e}")
            return None

    def _hash_password(self, password):
        """明文密码经 1024 次 MD5 迭代后的十六进制串（每轮对上一轮 hex 再 md5）"""
        h = password
        for _ in range(1024):
            h = hashlib.md5(h.encode("utf-8")).hexdigest()
        return h

    def _fetch_user_pid(self):
        """登录后调用「获取用户详情」接口，取出 pid 并缓存；顺带缓存真实姓名"""
        self.current_realname = None
        try:
            url = f"{self.base_url}/detectionManager/core/users/info"
            resp = self.session.get(url)
            if resp.status_code != 200:
                return ""
            data = resp.json()
            result = data.get("resultData") if isinstance(data, dict) else data
            user_info = result.get("userInfo") if isinstance(result, dict) else None
            user_ext = result.get("userExtInfo") if isinstance(result, dict) else None
            # 用户字段分布在 userInfo / userExtInfo，合并后探测 realName / pid
            merged = {}
            for part in (user_info, user_ext):
                if isinstance(part, dict):
                    merged.update(part)
            if merged:
                self.current_realname = self._extract_realname(merged)
                self.current_org_id = merged.get("orgId")  # 所属实验室(查询机构过滤用，见 DetectionAPI._current_org_id)
                return self._extract_pid(merged)
            return ""
        except Exception as e:
            print(f"获取用户详情失败: {e}")
            return ""

    def _extract_realname(self, payload):
        """从用户详情中探测真实姓名（字段名未知，按候选顺序取首个非空）"""
        for key in ("realName", "realname", "nickName", "empName", "staffName", "name", "xm", "cname", "userName"):
            val = payload.get(key)
            if val not in (None, "", 0):
                return str(val)
        return ""

    def _extract_pid(self, payload):
        """从用户详情中探测 pid（字段名未知，按候选顺序取首个非空）"""
        for key in ("pid", "id", "userId", "personId", "staffId", "empId", "employeeId"):
            val = payload.get(key)
            if val not in (None, "", 0):
                return str(val)
        return ""

    def login_with_user(self, username, captcha):
        """使用指定用户登录 - 修复版本，使用验证码会话"""
        if username not in self.users:
            return False, f"用户 {username} 未在配置中找到"

        user_info = self.users[username]
        # ponytail: 双格式兼容 — 既有 password_hash 用户(不可逆)直接用；新用户明文 password 走哈希
        if "password_hash" in user_info:
            password_hash = user_info["password_hash"]
        else:
            password_hash = self._hash_password(user_info.get("password", ""))

        # 登录数据
        login_data = {
            "account": username,
            "password": password_hash,
            "validCode": captcha
        }

        login_url = f"{self.base_url}/detectionManager/core/security/login"

        try:
            # 使用验证码会话进行登录
            if self.captcha_session:
                response = self.captcha_session.post(login_url, data=login_data)
                # 登录成功后，将验证码会话的cookies合并到主会话
                if response.status_code == 200:
                    # 将验证码会话的cookies复制到主会话
                    for cookie in self.captcha_session.cookies:
                        self.session.cookies.set(cookie.name, cookie.value)
            else:
                # 如果没有验证码会话，使用主会话
                response = self.session.post(login_url, data=login_data)

            # 检查响应状态
            if response.status_code != 200:
                return False, f"登录请求失败，状态码: {response.status_code}"

            result = response.json()

            if result.get("success"):
                user_data = result.get("resultData", {})

                # 检查是否有错误代码
                if user_data.get("errorCode"):
                    error_msg = user_data.get('errorMsg', '未知错误')
                    # 如果是验证码错误，特殊处理
                    if "验证码" in error_msg or "validCode" in error_msg:
                        return False, "验证码错误，请刷新验证码重试"
                    return False, f"登录失败: {error_msg}"
                else:
                    # 更新当前用户
                    self.current_user = username

                    # 获取并缓存用户 pid（用于后续会话校验/保活）
                    self.current_pid = self._fetch_user_pid()

                    # 保存会话
                    self.save_session()

                    # 启动会话保持（根据时间判断是否启用）
                    self.start_keep_alive()

                    # 返回用户信息（优先用「用户详情」里的真实姓名，其次登录响应的 nickName）
                    display_name = self.current_realname or user_data.get('nickName') or user_data.get('username') or username
                    # 用真实姓名更新显示名，供「已登录」状态显示（持久化以便自动登录也显示真实姓名）
                    if display_name and display_name != username:
                        self.users[username]['display_name'] = display_name
                        self.save_users()
                    user_info = f"用户: {display_name}"
                    return True, user_info
            else:
                error_msg = result.get("errorCtx", {}).get("errorMsg", "未知错误")
                # 如果是验证码错误，特殊处理
                if "验证码" in error_msg or "validCode" in error_msg:
                    return False, "验证码错误，请刷新验证码重试"
                return False, f"登录失败: {error_msg}"

        except Exception as e:
            return False, f"登录过程中出错: {e}"

    def logout(self):
        """退出登录"""
        # 停止会话保持
        self.stop_keep_alive()

        self.current_user = None
        self.session.cookies.clear()
        self.captcha_session = None
        # 清除会话文件
        if os.path.exists(self.session_file):
            os.remove(self.session_file)
        return True, "已退出登录"

    def should_keep_alive(self):
        """判断当前时间是否应该保持会话（不主动清除会话，只返回状态）"""
        now = datetime.datetime.now()
        current_hour = now.hour

        # 如果当前时间在18点之后，不保持会话
        if current_hour >= 20:
            return False

        return True

    def start_keep_alive(self):
        """启动会话保持线程（根据时间判断是否启用）"""
        # 检查当前时间是否应该保持会话
        if not self.should_keep_alive():
            print("当前时间已超过18点，暂停会话保持")
            self.keep_alive_enabled = False
            return  # 不主动清除会话，只是停止保持

        self.keep_alive_enabled = True

        if self.keep_alive_flag:
            return  # 已经在运行

        self.keep_alive_flag = True
        self.keep_alive_thread = threading.Thread(target=self._keep_alive_worker, daemon=True)
        self.keep_alive_thread.start()

    def stop_keep_alive(self):
        """停止会话保持线程"""
        self.keep_alive_flag = False
        self.keep_alive_enabled = False
        if self.keep_alive_thread:
            self.keep_alive_thread.join(timeout=1)
            self.keep_alive_thread = None

    def _keep_alive_worker(self):
        """后台工作线程，定期执行保持会话的操作"""
        while self.keep_alive_flag and self.keep_alive_enabled:
            try:
                # 每次执行前检查时间
                if not self.should_keep_alive():
                    print("当前时间已超过18点，停止会话保持")
                    self.keep_alive_flag = False
                    break

                if self.current_user and self.verify_session():
                    # 执行一个轻量级的操作来保持会话活跃
                    self.perform_keep_alive_action()
                else:
                    # 会话已失效，停止保持
                    self.keep_alive_flag = False
                    break
            except Exception as e:
                print(f"⚠️ 会话保持操作出错: {e}")

            # 等待指定间隔
            for _ in range(self.keep_alive_interval):
                if not self.keep_alive_flag or not self.keep_alive_enabled:
                    break
                time.sleep(1)

    def perform_keep_alive_action(self):
        """执行保持会话的轻量级操作"""
        try:
            # 方法1: 简单的用户信息查询（轻量级）
            user_info_url = f"{self.base_url}/detectionManager/core/security/getLoginUser"
            response = self.session.get(user_info_url)

            if response.status_code == 200:
                result = response.json()
                if result.get("success"):
                    print(f"🔄 会话保持成功: {time.strftime('%H:%M:%S')}")
                    return True

            # 方法2: 如果方法1失败，尝试一个简单的搜索
            return self.perform_light_search()

        except Exception as e:
            print(f"⚠️ 保持会话操作失败: {e}")
            return False

    def perform_light_search(self):
        """执行一个轻量级的搜索来保持会话"""
        try:
            # 使用登录时缓存的 pid
            user_pid = self.current_pid or ""

            # 使用一个非常简单的搜索条件，减少服务器负载
            test_params = {
                "_search": "false",
                "nd": str(int(time.time() * 1000)),
                "pageSize": 1,
                "pageNo": 1,
                "sidx": "",
                "sord": "asc",
                "acceptStartDate": time.strftime("%Y-%m-%d"),
                "acceptEndDate": time.strftime("%Y-%m-%d"),
                "retrievalWay": "decideProjectName",
                "decideProjectOrgId": "23",
                "checkInStatus": "CHECK_IN_STATUS_NO",
                "keyword": "test",
                "pid": user_pid,
                "pname": self.current_user
            }

            headers = {
                "Referer": f"{self.base_url}/web/detectionResultCheckInListMgt.html?menuId=41",
                "User-Agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.64 Mobile Safari/537.36"
            }

            url = f"{self.base_url}/detectionManager/manager/resultCheckIn/pagePCObjAndSample"
            response = self.session.get(url, params=test_params, headers=headers, timeout=5)

            if response.status_code == 200:
                result = response.json()
                if result.get("success"):
                    pass
                    return True

            return False

        except Exception as e:
            print(f"⚠️ 轻量级搜索失败: {e}")
            return False


class CompactLoginApp:
    def __init__(self, root):
        self.root = root
        self.root.title("登录系统")

        # 单实例检查
        if not self.check_single_instance():
            # 短暂显示提示然后退出
            self.root.withdraw()  # 隐藏窗口
            messagebox.showwarning("提示", "程序已经在运行中")
            self.root.quit()
            return

        # 图像引用列表 - 防止图像被垃圾回收
        self.image_references = []

        # 设置窗口图标
        self.set_window_icon()

        # 设置窗口大小
        window_width = 360
        window_height = 240

        # 获取屏幕尺寸
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        # 计算窗口居中的位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        # 设置窗口位置和大小
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.resizable(False, False)

        # 设置关闭窗口行为
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 创建登录系统实例
        self.login_system = MultiUserLoginSystem()

        # 系统托盘图标
        self.tray_icon = None

        # 单实例socket
        self.single_instance_socket = None

        # 创建界面
        self.create_widgets()

        # 初始化时检查时间状态
        self.initialize_session_state()

        # 等待窗口完全初始化后再刷新验证码
        def delayed_init():
            # 延迟验证码刷新，确保UI完全初始化
            self.root.after(200, self.refresh_captcha)
            # 延迟自动登录
            self.root.after(300, self.try_auto_login)

        self.root.after(100, delayed_init)

    def delayed_auto_login(self):
        """延迟自动登录，确保UI完全初始化"""
        try:
            self.try_auto_login()
        except Exception as e:
            print(f"自动登录时出错: {e}")
            import traceback
            traceback.print_exc()
            # 即使出错，也刷新验证码
            self.refresh_captcha()

    def check_single_instance(self):
        """检查是否已经有实例在运行"""
        try:
            # 尝试连接到已存在的实例
            test_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            test_socket.settimeout(1)  # 设置超时
            test_socket.connect(('localhost', 54321))
            test_socket.close()
            return False  # 连接成功，说明已有实例
        except socket.error:
            # 连接失败，创建新的服务器
            try:
                self.single_instance_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                self.single_instance_socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
                self.single_instance_socket.bind(('localhost', 54321))
                self.single_instance_socket.listen(1)

                # 在后台线程中处理连接请求
                def handle_connections():
                    while True:
                        try:
                            conn, addr = self.single_instance_socket.accept()
                            conn.close()
                            # 当有新的连接尝试时，恢复窗口
                            self.root.after(0, self.bring_to_front)
                        except:
                            break

                threading.Thread(target=handle_connections, daemon=True).start()
                return True
            except socket.error:
                return False

    def bring_to_front(self):
        """将窗口带到前台"""
        if self.tray_icon:
            self.show_window()
        else:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()

    def cleanup_single_instance(self):
        """清理单实例资源"""
        if self.single_instance_socket:
            try:
                self.single_instance_socket.close()
            except:
                pass

    def initialize_session_state(self):
        """初始化会话状态（只更新显示，不主动清除会话）"""
        # 检查当前时间是否允许保持会话
        if not self.login_system.should_keep_alive():
            print("⏰ 程序启动时检测到时间已过期")
            # 只更新显示状态，不清除实际会话
            if hasattr(self, 'status_label'):
                self.status_label.configure(text="未登录", foreground="red")

    def find_icon_file(self, icon_file):
        """查找图标文件"""
        # 首先在当前目录查找
        if os.path.exists(icon_file):
            return icon_file

        # 在打包环境中，尝试在 _internal 目录查找
        if getattr(sys, 'frozen', False):
            # 尝试在 exe 同级目录查找
            exe_dir = os.path.dirname(sys.executable)
            icon_path = os.path.join(exe_dir, icon_file)
            if os.path.exists(icon_path):
                return icon_path

            # 尝试在 _internal 目录查找
            internal_dir = os.path.join(exe_dir, '_internal')
            if os.path.exists(internal_dir):
                icon_path = os.path.join(internal_dir, icon_file)
                if os.path.exists(icon_path):
                    return icon_path

        # 尝试在脚本所在目录查找
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, icon_file)
        if os.path.exists(icon_path):
            return icon_path

        # 尝试在父目录查找
        parent_dir = os.path.dirname(script_dir)
        icon_path = os.path.join(parent_dir, icon_file)
        if os.path.exists(icon_path):
            return icon_path

        return None

    def set_window_icon(self):
        """设置窗口图标 - 修复版本"""
        try:
            # 尝试加载多个可能的图标文件
            icon_files = [
                "login.ico",  # 备用的登录图标
                "release_rate_control.ico",  # 主程序图标
                "settings.ico"  # 设置图标
            ]

            icon_loaded = False
            for icon_file in icon_files:
                icon_path = self.find_icon_file(icon_file)
                if icon_path:
                    try:
                        self.root.iconbitmap(icon_path)
                        print(f"✓ 设置窗口图标: {icon_file}")
                        icon_loaded = True
                        break
                    except Exception as e:
                        print(f"❌ 设置图标 {icon_file} 失败: {e}")
                        continue

            if not icon_loaded:
                print("⚠️ 所有图标文件都未找到或无法加载")
                # 尝试使用 PIL 创建简单图标
                try:
                    # 创建一个简单的绿色图标
                    img = Image.new('RGB', (32, 32), 'green')
                    draw = ImageDraw.Draw(img)
                    draw.ellipse([(2, 2), (30, 30)], fill='green', outline='white')

                    # 转换为 PhotoImage
                    photo = ImageTk.PhotoImage(img)
                    self.root.iconphoto(False, photo)
                    self.image_references.append(photo)  # 保存引用
                    print("✓ 使用程序生成的图标")
                except Exception as e:
                    print(f"❌ 生成图标失败: {e}")

        except Exception as e:
            print(f"❌ 设置窗口图标失败: {e}")

    def create_tray_icon(self):
        """创建系统托盘图标"""
        try:
            # 首先尝试使用login.ico文件
            icon_files = ["login.ico", "release_rate_control.ico", "settings.ico"]
            image = None

            for icon_file in icon_files:
                icon_path = self.find_icon_file(icon_file)
                if icon_path:
                    try:
                        image = Image.open(icon_path)
                        # 调整大小为系统托盘图标的标准尺寸
                        image = image.resize((64, 64), Image.Resampling.LANCZOS)
                        print(f"✓ 加载系统托盘图标: {icon_file}")
                        break
                    except Exception as e:
                        print(f"❌ 加载图标 {icon_file} 失败: {e}")
                        continue

            if image is None:
                # 如果找不到ico文件，使用程序生成的简单图标
                print("⚠️ 未找到图标文件，使用程序生成的图标")
                image = Image.new('RGB', (64, 64), 'green')
                dc = ImageDraw.Draw(image)
                dc.ellipse([(2, 2), (62, 62)], fill='green', outline='white')

            # 创建托盘菜单
            menu = (
                item('显示窗口', self.show_window),
                item('退出', self.quit_application)
            )

            self.tray_icon = pystray.Icon(
                "login_system",
                image,
                "登录系统",
                menu
            )
        except Exception as e:
            print(f"创建系统托盘图标失败: {e}")
            # 如果失败，不创建系统托盘图标
            self.tray_icon = None

    def show_window(self, icon=None, item=None):
        """从系统托盘恢复窗口"""
        if self.tray_icon:
            self.tray_icon.stop()
            self.tray_icon = None

        # 恢复窗口前检查会话状态
        self.check_session_on_restore()

        self.root.after(0, self.root.deiconify)

    def check_session_on_restore(self):
        """恢复窗口时检查会话状态（只更新显示，不主动清除会话）"""
        if self.login_system.current_user:
            # 检查时间是否应该保持会话
            if not self.login_system.should_keep_alive():
                # 时间已过期，只更新显示状态，不清除实际会话
                self.status_label.configure(text="未登录", foreground="red")
                print("⏰ 恢复窗口时检测到时间过期，已更新显示状态")
            else:
                # 时间允许，验证会话是否仍然有效
                if not self.login_system.verify_session():
                    # 会话无效，更新显示状态
                    self.status_label.configure(text="未登录", foreground="red")
                    print("🔄 恢复窗口时检测到会话无效，已更新显示状态")
                else:
                    # 会话有效，更新状态显示
                    display_name = self.login_system.users.get(self.login_system.current_user, {}).get('display_name',
                                                                                                       self.login_system.current_user)
                    self.status_label.configure(text=f"已登录: {display_name}", foreground="green")

    def quit_application(self, icon=None, item=None):
        """退出应用程序"""
        # 清理单实例资源
        self.cleanup_single_instance()

        if self.tray_icon:
            self.tray_icon.stop()

        # 停止会话保持
        self.login_system.stop_keep_alive()

        # 完全退出程序
        self.root.quit()
        self.root.destroy()

    def on_closing(self):
        """处理窗口关闭事件"""
        # 在最小化到托盘前检查会话状态，只更新显示
        if self.login_system.current_user:
            # 检查时间是否允许保持会话
            if not self.login_system.should_keep_alive():
                # 时间已过期，只更新显示状态
                self.status_label.configure(text="未登录", foreground="red")
                print("⏰ 关闭窗口时检测到时间过期，已更新显示状态")

        # 检查当前显示状态决定是否最小化到托盘
        current_status = self.status_label.cget("text")
        if "已登录" in current_status:
            # 有用户登录，最小化到系统托盘
            self.minimize_to_tray()
        else:
            # 没有用户登录，退出程序
            self.quit_application()

    def minimize_to_tray(self):
        """最小化到系统托盘"""
        # 隐藏主窗口
        self.root.withdraw()

        # 创建系统托盘图标（如果还没有）
        if not self.tray_icon:
            self.create_tray_icon()

        # 在单独的线程中运行系统托盘
        if self.tray_icon:
            threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def create_widgets(self):
        """创建界面组件"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 第一行：用户选择区域
        user_frame = ttk.Frame(main_frame)
        user_frame.pack(fill=tk.X, pady=(13, 10))

        ttk.Label(user_frame, text="用户", width=5, anchor="w").pack(side=tk.LEFT)

        self.user_var = tk.StringVar()
        self.user_combo = ttk.Combobox(user_frame, textvariable=self.user_var, state="readonly", width=15)
        self.user_combo.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)

        # 第二行：验证码区域
        captcha_frame = ttk.Frame(main_frame)
        captcha_frame.pack(fill=tk.X, pady=(5, 5))

        # 验证码标签放在最左边
        ttk.Label(captcha_frame, text="验证码", width=5, anchor="w").pack(side=tk.LEFT)

        # 验证码输入框
        self.captcha_var = tk.StringVar()
        self.captcha_entry = ttk.Entry(captcha_frame, textvariable=self.captcha_var, width=21)
        self.captcha_entry.pack(side=tk.LEFT, padx=(5, 10), fill=tk.X, expand=True)

        # 验证码图片紧贴输入框右侧
        self.captcha_label = tk.Label(captcha_frame, cursor="hand2")
        self.captcha_label.pack(side=tk.LEFT)

        # 创建一个空的初始图像
        self._captcha_photo = None

        self.captcha_label.bind("<Button-1>", lambda e: self.refresh_captcha())
        # 第三行：按钮区域 - 居中对齐
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 10))

        # 创建一个内部框架用于居中对齐按钮
        button_container = ttk.Frame(button_frame)
        button_container.pack(expand=True)

        # 登录按钮
        login_btn = ttk.Button(button_container, text="登录", command=self.login, width=10)
        login_btn.pack(side=tk.LEFT, padx=(0, 5))

        # 新用户配置按钮
        new_user_btn = ttk.Button(button_container, text="新用户", command=self.show_new_user_dialog, width=10)
        new_user_btn.pack(side=tk.LEFT, padx=5)

        # 退出登录按钮
        logout_btn = ttk.Button(button_container, text="退出登录", command=self.logout, width=10)
        logout_btn.pack(side=tk.LEFT, padx=(5, 0))

        # 第四行：当前用户状态 - 显示在最后一行
        self.status_label = ttk.Label(main_frame, text="未登录", font=("SimSun", 10))
        self.status_label.configure(foreground="red")
        self.status_label.pack(fill=tk.X)

        # 绑定回车键到登录功能
        self.root.bind('<Return>', lambda event: self.login())
        # 绑定Tab键在输入框和验证码之间切换
        self.captcha_entry.bind('<Tab>', self.focus_next_widget)

        # 填充用户列表
        self.update_user_list()

    def focus_next_widget(self, event):
        """Tab键焦点切换"""
        event.widget.tk_focusNext().focus()
        return "break"

    def update_user_list(self):
        """更新用户列表"""
        users = list(self.login_system.users.keys())
        self.user_combo['values'] = users
        if users:
            self.user_combo.current(0)

    def try_auto_login(self):
        """尝试自动登录"""
        try:
            # 先检查当前时间是否允许自动登录
            if not self.login_system.should_keep_alive():
                print("⏰ 当前时间不允许自动登录")
                self.status_label.configure(text="未登录", foreground="red")
                return

            # 如果有保存的会话，尝试自动登录
            if self.login_system.load_session() and self.login_system.verify_session():
                if self.login_system.current_user:
                    # 自动登录成功
                    display_name = self.login_system.users.get(self.login_system.current_user, {}).get('display_name',
                                                                                                       self.login_system.current_user)
                    self.status_label.configure(text=f"已登录: {display_name}", foreground="green")
                    # 设置用户选择框为当前用户
                    if self.login_system.current_user in self.login_system.users:
                        self.user_combo.set(self.login_system.current_user)

                    # 启动会话保持（根据时间判断是否启用）
                    self.login_system.start_keep_alive()
                    return

            # 自动登录失败，保持验证码显示
            self.status_label.configure(text="未登录", foreground="red")

        except Exception as e:
            print(f"自动登录过程中出错: {e}")
            import traceback
            traceback.print_exc()
            self.status_label.configure(text="未登录", foreground="red")
            # 确保验证码已刷新
            self.refresh_captcha()

    def refresh_captcha(self):
        """刷新验证码图片 - 修复版本，确保图像引用"""
        try:
            # 确保在主线程中执行GUI操作
            if not hasattr(self, '_captcha_lock'):
                self._captcha_lock = threading.Lock()

            def _refresh():
                try:
                    # 清除旧的验证码输入
                    self.captcha_var.set("")

                    # 获取新的验证码
                    captcha_image = self.login_system.get_captcha_image()
                    if captcha_image:
                        # 调整图片大小
                        captcha_image = captcha_image.resize((80, 30), Image.Resampling.LANCZOS)

                        # 创建PhotoImage对象
                        captcha_photo = ImageTk.PhotoImage(captcha_image)

                        # 直接保存引用，防止被垃圾回收
                        if not hasattr(self, '_captcha_photo'):
                            self._captcha_photo = captcha_photo
                        else:
                            # 更新引用，确保旧的图像不会被垃圾回收
                            self._captcha_photo = captcha_photo

                        # 更新标签
                        self.captcha_label.configure(image=self._captcha_photo)

                        # 确保图像引用被保留
                        self.image_references.append(self._captcha_photo)

                        # 限制保存的图像引用数量，防止内存泄漏
                        if len(self.image_references) > 10:
                            self.image_references.pop(0)

                        print("✓ 验证码已刷新")
                    else:
                        print("获取验证码失败")

                except Exception as e:
                    print(f"刷新验证码时出错: {e}")
                    import traceback
                    traceback.print_exc()

            # 在主线程中执行GUI更新
            self.root.after(0, _refresh)

        except Exception as e:
            print(f"刷新验证码时出错: {e}")
            import traceback
            traceback.print_exc()
            # 延迟重试
            self.root.after(100, self.refresh_captcha_retry)

    def refresh_captcha_retry(self):
        """重试验证码刷新"""
        try:
            print("正在重试验证码刷新...")
            self.refresh_captcha()
        except Exception as e:
            print(f"重试失败: {e}")

    def login(self):
        """登录操作"""
        selected = self.user_combo.get()
        if not selected:
            messagebox.showerror("错误", "请选择用户")
            return

        username = selected
        captcha = self.captcha_var.get().strip()

        if not captcha:
            messagebox.showerror("错误", "请输入验证码")
            return

        # 如果当前时间已过期，需要先清除会话
        if not self.login_system.should_keep_alive():
            # 时间已过期，清除会话
            self.login_system.logout()
            self.status_label.configure(text="未登录", foreground="red")

        # 首先尝试使用会话自动登录
        success, message = self.login_system.ensure_login(username)

        if success:
            # 自动登录成功
            display_name = self.login_system.users[username]['display_name']
            self.status_label.configure(text=f"已登录: {display_name}", foreground="green")
            # 启动会话保持（根据时间判断是否启用）
            self.login_system.start_keep_alive()
        else:
            # 自动登录失败，使用验证码登录
            success, message = self.login_system.login_with_user(username, captcha)

            if success:
                # 更新当前用户显示
                display_name = self.login_system.users[username]['display_name']
                self.status_label.configure(text=f"已登录: {display_name}", foreground="green")
            else:
                # 如果是验证码错误，自动刷新验证码
                if "验证码" in message or "validCode" in message:
                    self.refresh_captcha()
                    messagebox.showerror("登录失败", "验证码错误，已自动刷新，请重新输入")
                else:
                    messagebox.showerror("登录失败", message)

    def logout(self):
        """退出登录操作"""
        if self.login_system.current_user:
            success, message = self.login_system.logout()
            if success:
                self.status_label.configure(text="未登录", foreground="red")
                self.captcha_var.set("")
                self.refresh_captcha()
        else:
            # 未登录状态下点击退出登录，退出程序
            self.quit_application()

    def show_new_user_dialog(self):
        """显示新用户配置对话框"""
        dialog = ttkb.Toplevel(self.root)
        dialog.title("新用户配置")
        dialog.geometry("500x270")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        # 设置对话框图标
        try:
            icon_path = self.find_icon_file("login.ico")
            if icon_path:
                dialog.iconbitmap(icon_path)
        except:
            pass

        # 设置对话框居中显示
        self.center_window(dialog)

        # 创建对话框内容
        main_frame = ttk.Frame(dialog, padding="18")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 表单（grid 布局：标签右对齐，输入框垂直对齐）
        form_frame = ttk.Frame(main_frame)
        form_frame.pack(fill=tk.X, pady=(0, 10))
        form_frame.columnconfigure(1, weight=1)

        ttk.Label(form_frame, text="登录名:").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=7)
        account_entry = ttk.Entry(form_frame)
        account_entry.grid(row=0, column=1, sticky="ew", pady=7)

        ttk.Label(form_frame, text="密码:").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=7)
        password_entry = ttk.Entry(form_frame)
        password_entry.grid(row=1, column=1, sticky="ew", pady=7)

        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)

        # 创建一个内部框架用于居中对齐按钮
        button_container = ttk.Frame(button_frame)
        button_container.pack(expand=True)

        def save_user():
            account = account_entry.get().strip()
            password = password_entry.get()

            if not account or not password:
                messagebox.showerror("错误", "请输入登录名和密码")
                return

            # 检查用户是否已存在
            if account in self.login_system.users:
                old_password = self.login_system.users[account].get("password", "")
                if old_password == password:
                    # 密码相同，只更新显示信息
                    overwrite = messagebox.askyesno("用户已存在",
                                                    f"用户 {account} 已存在且密码相同，是否更新显示信息?")
                    if not overwrite:
                        return
                else:
                    # 密码不同，询问是否更新密码
                    overwrite = messagebox.askyesno("密码已更改",
                                                    f"检测到用户 {account} 的密码已更改，是否更新密码?\n\n"
                                                    f"注意：更新密码后需要重新登录")
                    if not overwrite:
                        return
                    else:
                        # 强制刷新会话
                        self.login_system.force_refresh_session(account)

            # 保存/更新用户信息（仅记录登录名 + 明文密码）
            self.login_system.users[account] = {
                "password": password,
                "display_name": account,
            }

            self.login_system.save_users()
            self.update_user_list()

            # 显示成功消息
            messagebox.showinfo("保存成功", f"用户 {account} 配置已保存")

            account_entry.delete(0, tk.END)
            password_entry.delete(0, tk.END)

            # 如果更新的是当前用户，更新状态显示
            if self.login_system.current_user == account:
                self.status_label.configure(text="未登录", foreground="red")

        def show_existing_users():
            """显示已有用户信息"""
            users_window = ttkb.Toplevel(dialog)
            users_window.title("已有用户信息")
            users_window.geometry("460x500")
            users_window.transient(dialog)

            # 设置用户信息窗口图标
            try:
                icon_path = self.find_icon_file("login.ico")
                if icon_path:
                    users_window.iconbitmap(icon_path)
            except:
                pass

            # 设置用户信息窗口居中显示
            self.center_window(users_window)

            # 创建主框架
            main_frame = ttk.Frame(users_window, padding="16")
            main_frame.pack(fill=tk.BOTH, expand=True)

            # 创建滚动框架
            canvas = tk.Canvas(main_frame, borderwidth=0)
            scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)

            # 配置滚动区域
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            # 显示每个用户的信息
            for username, info in self.login_system.users.items():
                user_frame = ttk.Frame(scrollable_frame, relief="solid", borderwidth=1)
                user_frame.pack(fill=tk.X, padx=8, pady=4)

                # 用户信息标签
                user_info = f"账号名: {username}\n密码: *******************************"
                user_label = ttk.Label(user_frame, text=user_info, justify=tk.LEFT)
                user_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10, pady=8)

                # 删除按钮
                delete_btn = ttkb.Button(user_frame, text="删除", width=6, bootstyle="danger",
                                         command=lambda u=username: self.delete_user(u, users_window))
                delete_btn.pack(side=tk.RIGHT, padx=10, pady=8)

            # 打包Canvas和Scrollbar
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            # 绑定鼠标滚轮滚动
            def _on_mousewheel(event):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        save_btn = ttkb.Button(button_container, text="保存", command=save_user, width=12, bootstyle="primary")
        save_btn.pack(side=tk.LEFT, padx=6)

        view_users_btn = ttkb.Button(button_container, text="查看已有用户", command=show_existing_users, width=12, bootstyle="secondary")
        view_users_btn.pack(side=tk.LEFT, padx=6)

    def center_window(self, window):
        """将窗口居中显示"""
        window.update_idletasks()
        width = window.winfo_width()
        height = window.winfo_height()
        x = (window.winfo_screenwidth() // 2) - (width // 2)
        y = (window.winfo_screenheight() // 2) - (height // 2)
        window.geometry('{}x{}+{}+{}'.format(width, height, x, y))

    def delete_user(self, username, users_window):
        """删除用户"""
        # 确认删除
        result = messagebox.askyesno("确认删除", f"确定要删除用户 {username} 吗？")
        if not result:
            return

        # 从用户列表中删除
        if username in self.login_system.users:
            del self.login_system.users[username]
            self.login_system.save_users()

            # 如果删除的是当前登录用户，则退出登录
            if self.login_system.current_user == username:
                self.login_system.current_user = None
                self.status_label.configure(text="未登录", foreground="red")

            # 更新用户列表
            self.update_user_list()

            # 关闭用户信息窗口
            users_window.destroy()
        else:
            messagebox.showerror("删除失败", f"用户 {username} 不存在")


class LoginAppWrapper:
    """登录应用包装器，提供启动接口"""

    def __init__(self):
        self.root = None
        self.app = None

    def start(self, standalone=True):
        """启动登录应用
        Args:
            standalone: 是否为独立运行，True则创建新Tk窗口
        """
        if standalone:
            # 独立运行模式
            main()
        else:
            # 被调用模式，在主线程中运行
            self.run_in_main_thread()

    def run_in_main_thread(self):
        """在主线程中运行登录应用"""
        import threading
        threading.Thread(target=main, daemon=True).start()


def run_login():
    """供外部调用的登录函数"""
    import threading
    thread = threading.Thread(target=main, daemon=True)
    thread.start()
    return thread


def main():
    """主函数"""
    root = ttkb.Window(themename="sandstone-light")
    app = CompactLoginApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()