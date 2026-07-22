import sys
import tkinter as tk
from tkinter import ttk, messagebox
import requests
import hashlib
import json
import time
import os
from PIL import Image, ImageTk, ImageDraw
from io import BytesIO
import urllib3
from urllib.parse import parse_qs
import threading
import pystray
from pystray import MenuItem as item
import datetime
import socket  # 新增导入

# 禁用SSL警告
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


class MultiUserLoginSystem:
    def __init__(self):
        self.session = requests.Session()
        self.base_url = "http://192.168.12.234:60015"
        self.session.verify = False
        self.current_user = None
        self.users_file = "users_config.json"
        self.session_file = "session_info.json"

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
        """加载用户配置"""
        # 如果存在配置文件，则加载
        if os.path.exists(self.users_file):
            try:
                with open(self.users_file, 'r', encoding='utf-8') as f:
                    users = json.load(f)
                return users
            except Exception as e:
                print(f"加载用户配置文件失败: {e}")
                # 如果加载失败，返回空字典
                return {}
        else:
            # 如果配置文件不存在，创建空的配置文件
            self.save_users({})
            return {}

    def save_users(self, users=None):
        """保存用户配置"""
        if users is None:
            users = self.users

        with open(self.users_file, 'w', encoding='utf-8') as f:
            json.dump(users, f, ensure_ascii=False, indent=2)

    def load_session(self):
        """尝试加载保存的会话"""
        if os.path.exists(self.session_file):
            try:
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
                return True
            except Exception as e:
                print(f"❌ 加载会话失败: {e}")
                # 加载失败时删除损坏的会话文件
                if os.path.exists(self.session_file):
                    os.remove(self.session_file)

        return False

    def save_session(self):
        """保存会话信息"""
        try:
            # 处理cookies，避免重复名称
            cookies_dict = {}
            for cookie in self.session.cookies:
                cookies_dict[cookie.name] = cookie.value

            session_info = {
                "username": self.current_user,
                "login_time": time.strftime("%Y-%m-%d %H:%M:%S"),
                "cookies": cookies_dict,
                "headers": dict(self.session.headers)
            }

            with open(self.session_file, "w", encoding="utf-8") as f:
                json.dump(session_info, f, ensure_ascii=False, indent=2)

            return True
        except Exception as e:
            print(f"❌ 保存会话信息失败: {e}")
            return False

    def verify_session(self):
        """验证会话是否仍然有效 - 通过实际搜索来验证"""
        if not self.current_user:
            return False

        try:
            # 获取当前用户的配置信息
            user_info = self.users.get(self.current_user, {})
            user_pid = user_info.get("pid", "")

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
            response = self.session.get(url, params=test_params, headers=headers)

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

    def parse_raw_data(self, raw_data):
        """从原始请求数据中解析用户信息"""
        try:
            # 移除可能的前缀（如"data-raw"）
            if raw_data.startswith('data-raw '):
                raw_data = raw_data.replace('data-raw ', '', 1)

            # 移除引号
            raw_data = raw_data.strip('"\'')

            # 解析URL编码的参数
            parsed_data = parse_qs(raw_data)

            # 提取参数（parse_qs返回的是列表，取第一个值）
            user_info = {
                'account': parsed_data.get('account', [''])[0],
                'password_hash': parsed_data.get('password', [''])[0],
                'pid': parsed_data.get('pid', [''])[0]
            }

            # 验证必要参数
            if not user_info['account'] or not user_info['password_hash']:
                return None, "解析失败：缺少必要参数(account或password)"

            return user_info, "成功解析用户信息"

        except Exception as e:
            return None, f"解析原始数据失败: {e}"

    def login_with_user(self, username, captcha):
        """使用指定用户登录 - 修复版本，使用验证码会话"""
        if username not in self.users:
            return False, f"用户 {username} 未在配置中找到"

        user_info = self.users[username]
        password_hash = user_info["password_hash"]

        # 登录数据
        login_data = {
            "account": username,
            "password": password_hash,
            "validCode": captcha
        }

        # 添加pid参数（如果存在）
        if user_info.get("pid"):
            login_data["pid"] = user_info["pid"]

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

                    # 保存会话
                    self.save_session()

                    # 启动会话保持（根据时间判断是否启用）
                    self.start_keep_alive()

                    # 返回用户信息
                    display_name = user_data.get('nickName', user_data.get('username', username))
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
            # 使用当前用户的pid
            user_pid = self.users.get(self.current_user, {}).get("pid", "")

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
            response = self.session.get(url, params=test_params, headers=headers)

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

        # 设置窗口图标
        self.set_window_icon()

        # 设置窗口大小
        window_width = 300
        window_height = 180

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

        # 尝试自动登录
        self.try_auto_login()

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
            self.status_label.configure(text="未登录", foreground="red")

    def set_window_icon(self):
        """设置窗口图标"""
        try:
            # 尝试从当前目录加载login.ico
            icon_path = "login.ico"
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
            else:
                # 如果当前目录没有，尝试在打包后的路径中查找
                # 对于pyinstaller打包，图标可能被包含在临时目录中
                if hasattr(sys, '_MEIPASS'):
                    icon_path = os.path.join(sys._MEIPASS, "login.ico")
                    if os.path.exists(icon_path):
                        self.root.iconbitmap(icon_path)
                    else:
                        print("⚠️ 未找到图标文件: login.ico")
                else:
                    print("⚠️ 未找到图标文件: login.ico")
        except Exception as e:
            print(f"❌ 设置窗口图标失败: {e}")

    def create_tray_icon(self):
        """创建系统托盘图标"""
        try:
            # 首先尝试使用login.ico文件
            icon_path = "login.ico"
            if os.path.exists(icon_path):
                # 使用ICO文件创建系统托盘图标
                image = Image.open(icon_path)
                # 调整大小为系统托盘图标的标准尺寸
                image = image.resize((64, 64), Image.Resampling.LANCZOS)
            else:
                # 如果找不到ico文件，使用程序生成的简单图标
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
                    print("🔄 恢复窗口时会话仍然有效")

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
        # 首先检查当前时间是否允许自动登录
        if not self.login_system.should_keep_alive():
            print("⏰ 当前时间不允许自动登录")
            self.status_label.configure(text="未登录", foreground="red")
            self.refresh_captcha()
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

        # 没有保存的会话或自动登录失败，刷新验证码
        self.status_label.configure(text="未登录", foreground="red")
        self.refresh_captcha()

    def refresh_captcha(self):
        """刷新验证码图片 - 修复版本"""
        # 清除旧的验证码输入
        self.captcha_var.set("")

        captcha_image = self.login_system.get_captcha_image()
        if captcha_image:
            # 调整图片大小
            captcha_image = captcha_image.resize((80, 30), Image.Resampling.LANCZOS)
            self.captcha_photo = ImageTk.PhotoImage(captcha_image)
            self.captcha_label.configure(image=self.captcha_photo)
        else:
            messagebox.showerror("错误", "获取验证码失败，请检查网络连接")

    def login(self):
        """登录操作 - 修复版本"""
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
        dialog = tk.Toplevel(self.root)
        dialog.title("新用户配置")
        dialog.geometry("300x200")
        dialog.resizable(False, False)
        dialog.transient(self.root)

        # 设置对话框图标
        try:
            icon_path = "login.ico"
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except:
            pass

        # 设置对话框居中显示
        self.center_window(dialog)

        # 创建对话框内容
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="粘贴配置信息:").pack(anchor=tk.W, pady=(0, 5))

        raw_data_text = tk.Text(main_frame, height=6, width=50)
        raw_data_text.pack(fill=tk.X, pady=(0, 10))

        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)

        # 创建一个内部框架用于居中对齐按钮
        button_container = ttk.Frame(button_frame)
        button_container.pack(expand=True)

        def parse_and_save():
            raw_data = raw_data_text.get(1.0, tk.END).strip()

            if not raw_data:
                messagebox.showerror("错误", "请输入原始请求数据")
                return

            parsed_data, parse_message = self.login_system.parse_raw_data(raw_data)

            if not parsed_data:
                messagebox.showerror("解析失败", parse_message)
                return

            username = parsed_data['account']
            new_password_hash = parsed_data['password_hash']
            new_pid = parsed_data.get('pid', '')

            # 检查用户是否已存在
            if username in self.login_system.users:
                # 检查密码是否相同
                old_password_hash = self.login_system.users[username]["password_hash"]
                if old_password_hash == new_password_hash:
                    # 密码相同，只更新显示信息
                    overwrite = messagebox.askyesno("用户已存在",
                                                    f"用户 {username} 已存在且密码相同，是否更新显示信息?")
                    if not overwrite:
                        return
                else:
                    # 密码不同，询问是否更新密码
                    overwrite = messagebox.askyesno("密码已更改",
                                                    f"检测到用户 {username} 的密码已更改，是否更新密码?\n\n"
                                                    f"注意：更新密码后需要重新登录")
                    if not overwrite:
                        return
                    else:
                        # 强制刷新会话
                        self.login_system.force_refresh_session(username)

            # 保存/更新用户信息
            self.login_system.users[username] = {
                "password_hash": new_password_hash,
                "display_name": username,
                "pid": new_pid
            }

            self.login_system.save_users()
            self.update_user_list()

            # 显示成功消息
            success_message = f"用户 {username} 配置已保存"
            if username in self.login_system.users and old_password_hash != new_password_hash:
                success_message += "\n密码已更新，请重新登录"

            messagebox.showinfo("保存成功", success_message)

            raw_data_text.delete(1.0, tk.END)

            # 如果更新的是当前用户，更新状态显示
            if self.login_system.current_user == username:
                self.status_label.configure(text="未登录", foreground="red")

        def show_existing_users():
            """显示已有用户信息"""
            users_window = tk.Toplevel(dialog)
            users_window.title("已有用户信息")
            users_window.geometry("300x400")
            users_window.transient(dialog)

            # 设置用户信息窗口图标
            try:
                icon_path = "login.ico"
                if os.path.exists(icon_path):
                    users_window.iconbitmap(icon_path)
            except:
                pass

            # 设置用户信息窗口居中显示
            self.center_window(users_window)

            # 创建主框架
            main_frame = ttk.Frame(users_window, padding="10")
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
                user_frame.pack(fill=tk.X, padx=5, pady=2)

                # 用户信息标签
                user_info = f"用户名: {username}\n显示名: {info.get('display_name', '')}\n密码: *******************************\nPID: {info.get('pid', '')}"
                user_label = ttk.Label(user_frame, text=user_info, justify=tk.LEFT)
                user_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5, pady=5)

                # 删除按钮
                delete_btn = ttk.Button(user_frame, text="删除", width=6,
                                        command=lambda u=username: self.delete_user(u, users_window))
                delete_btn.pack(side=tk.RIGHT, padx=5, pady=5)

            # 打包Canvas和Scrollbar
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

            # 绑定鼠标滚轮滚动
            def _on_mousewheel(event):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        parse_btn = ttk.Button(button_container, text="解析并保存", command=parse_and_save, width=12)
        parse_btn.pack(side=tk.LEFT, padx=5)

        view_users_btn = ttk.Button(button_container, text="查看已有用户", command=show_existing_users, width=12)
        view_users_btn.pack(side=tk.LEFT, padx=5)

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


def main():
    root = tk.Tk()
    app = CompactLoginApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()