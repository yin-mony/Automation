import os
import subprocess
import time
import winreg
import csv
import io
from datetime import datetime
from pathlib import Path

import psutil
from pywinauto import Desktop


class QiYeVxLogin:
    def __init__(self, exe_path=None, extra_search_paths=None):
        """初始化企业微信登录管理器，支持自定义可执行文件和额外搜索路径。"""
        self.extra_search_paths = extra_search_paths or []
        # 延迟解析可执行路径，避免在“企业微信已运行”时因路径未找到而初始化失败。
        self.exe_path = exe_path
        self.is_login = False
        self.login_time = None
        # 登录结果状态：
        # unknown / direct_login_success / wait_scan / scan_login_success /
        # scan_login_timeout / launch_failed
        self.login_state = "unknown"
        self.login_state_text = "未开始登录检测"
        self.last_detected_texts = []

    @staticmethod
    def _get_wxwork_pids():
        """获取企业微信相关进程 PID 集合。"""
        target_names = {"wxwork.exe", "wecom.exe", "wxworkapp.exe"}
        pids = set()

        # 优先用 tasklist，减少 psutil 在部分环境的异常影响。
        try:
            output = subprocess.check_output(
                ["tasklist", "/FO", "CSV", "/NH"],
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            text = output.decode("gbk", errors="ignore")
            reader = csv.reader(io.StringIO(text))
            for row in reader:
                if len(row) < 2:
                    continue
                image_name = row[0].strip().lower()
                if image_name in target_names:
                    try:
                        pids.add(int(row[1]))
                    except ValueError:
                        continue
            if pids:
                return pids
        except Exception:
            pass

        # 回退到 psutil
        for proc in psutil.process_iter():
            try:
                if (proc.name() or "").lower() in target_names:
                    pids.add(proc.pid)
            except Exception:
                continue
        return pids

    def resolve_wxwork(self):
        """多路径定位本机企业微信可执行文件，找不到则抛出异常。"""
        candidates = []

        # 1) 优先从正在运行的进程中获取可执行路径（最准确）
        for proc in psutil.process_iter(["name", "exe"]):
            name = (proc.info.get("name") or "").lower()
            exe = proc.info.get("exe")
            if name in {"wxwork.exe", "wecom.exe", "wxworkapp.exe"} and exe:
                candidates.append(Path(exe))

        # 2) 从注册表读取安装信息（DisplayIcon / InstallLocation）
        candidates.extend(self._get_exe_candidates_from_registry())

        # 3) 常见安装路径兜底
        candidates.extend([
            Path(r"C:\Program Files\WXWork\WXWork.exe"),
            Path(r"C:\Program Files (x86)\WXWork\WXWork.exe"),
            Path(r"D:\WXWork\WXWork.exe"),
            Path(os.environ.get("LOCALAPPDATA", "")) / "Tencent" / "WeCom" / "WXWork.exe",
            Path(os.environ.get("LOCALAPPDATA", "")) / "WXWork" / "WXWork.exe",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "WXWork" / "WXWork.exe",
            Path(os.environ.get("ProgramW6432", "")) / "WXWork" / "WXWork.exe",
            Path(os.environ.get("ProgramFiles", "")) / "WXWork" / "WXWork.exe",
            Path(os.environ.get("ProgramFiles(x86)", "")) / "WXWork" / "WXWork.exe",
        ])
        for custom_path in self.extra_search_paths:
            custom = Path(custom_path)
            if custom.suffix.lower() == ".exe":
                candidates.append(custom)
            else:
                candidates.append(custom / "WXWork.exe")
                candidates.append(custom / "WeCom.exe")

        for exe in candidates:
            if exe and exe.exists() and exe.is_file():
                return str(exe)

        search_roots = [
            Path(r"C:\Program Files"),
            Path(r"C:\Program Files (x86)"),
            Path(os.environ.get("LOCALAPPDATA", "")),
            Path(os.environ.get("ProgramW6432", "")),
            Path(os.environ.get("ProgramFiles", "")),
            Path(os.environ.get("ProgramFiles(x86)", "")),
            Path(r"D:\Program Files"),
            Path(r"D:\Program Files (x86)"),
        ]
        for custom_path in self.extra_search_paths:
            custom = Path(custom_path)
            if custom.exists():
                search_roots.append(custom if custom.is_dir() else custom.parent)

        for root in search_roots:
            if root.exists():
                try:
                    for exe in root.rglob("WXWork.exe"):
                        return str(exe)
                    for exe in root.rglob("WeCom.exe"):
                        return str(exe)
                except (PermissionError, OSError):
                    continue

        raise FileNotFoundError(
            "未找到企业微信可执行文件（WXWork.exe/WeCom.exe），可通过 extra_search_paths 传入安装目录。"
        )

    def _get_exe_candidates_from_registry(self):
        """从注册表卸载项中提取企业微信安装路径候选。"""
        candidates = []
        reg_roots = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
            (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Uninstall"),
        ]

        for hive, root_path in reg_roots:
            try:
                with winreg.OpenKey(hive, root_path) as root:
                    subkey_count = winreg.QueryInfoKey(root)[0]
                    for i in range(subkey_count):
                        sub_name = winreg.EnumKey(root, i)
                        try:
                            with winreg.OpenKey(root, sub_name) as sub:
                                display_name = self._read_reg_value(sub, "DisplayName")
                                if not display_name:
                                    continue
                                name = str(display_name).lower()
                                if ("企业微信" not in str(display_name)) and ("wecom" not in name) and ("wxwork" not in name):
                                    continue

                                display_icon = self._read_reg_value(sub, "DisplayIcon")
                                install_location = self._read_reg_value(sub, "InstallLocation")

                                if display_icon:
                                    icon_path = str(display_icon).strip().strip('"')
                                    # 兼容 "D:\WXWork\WXWork.exe,0" 这类格式
                                    if "," in icon_path:
                                        icon_path = icon_path.split(",", 1)[0].strip()
                                    candidates.append(Path(icon_path))

                                if install_location:
                                    install_dir = Path(str(install_location).strip().strip('"'))
                                    candidates.append(install_dir / "WXWork.exe")
                                    candidates.append(install_dir / "WeCom.exe")
                        except OSError:
                            continue
            except OSError:
                continue

        return candidates

    @staticmethod
    def _read_reg_value(key, value_name):
        """安全读取注册表值，读取失败返回 None。"""
        try:
            return winreg.QueryValueEx(key, value_name)[0]
        except OSError:
            return None

    def _is_process_running(self):
        """检查企业微信进程是否已在运行。"""
        return bool(self._get_wxwork_pids())

    def _show_client_window(self):
        """将企业微信窗口还原并置顶到前台，返回是否成功显示。"""
        window = self._get_wxwork_window(visible_only=False)
        if not window:
            return False
        try:
            try:
                window.restore()
            except Exception:
                pass
            window.set_focus()
            return True
        except Exception:
            return False

    def open_client(self, show_window=False):
        """启动企业微信；show_window=True 时确保弹出客户端窗口。"""
        try:
            if not self.exe_path:
                self.exe_path = self.resolve_wxwork()
        except FileNotFoundError as exc:
            print(f"启动企业微信失败: {exc}")
            return False

        if self._is_process_running():
            if show_window:
                if self._show_client_window():
                    return True
                # 进程已在但窗口未弹出时，再次执行程序触发前台窗口。
                try:
                    os.startfile(self.exe_path)
                    time.sleep(2)
                    return self._show_client_window()
                except Exception as exc:
                    print(f"唤起企业微信窗口异常: {exc}")
                    return False
            return True

        try:
            subprocess.Popen([self.exe_path])
            time.sleep(3)
            if show_window:
                return self._show_client_window()
            return True
        except Exception as exc:
            print(f"启动企业微信异常: {exc}")
            return False

    def _get_wxwork_window(self, visible_only=True):
        """获取当前可见的企业微信主窗口对象，未找到则返回 None。"""
        desktop = Desktop(backend="uia")
        wxwork_pids = self._get_wxwork_pids()
        if not wxwork_pids:
            return None

        windows = desktop.windows()
        # 仅匹配企业微信进程 PID 的窗口，避免误识别浏览器页面标题。
        for window in windows:
            try:
                if visible_only and (not window.is_visible()):
                    continue
                try:
                    pid = window.process_id()
                    if pid in wxwork_pids:
                        return window
                except Exception:
                    pass
            except Exception:
                continue
        return None

    def _detect_auth_phase(self):
        """返回认证阶段：logged_in / unauth / unknown。"""
        window = self._get_wxwork_window()
        if not window:
            return "unknown"

        title = window.window_text() or ""
        # 避免使用过宽泛的“登录”关键词导致误判（如“退出登录”）。
        unauth_markers = (
            "扫码登录",
            "切换账号",
            "手机端确认",
            "刷新二维码",
            "请使用企业微信扫码登录",
            "使用微信扫码登录",
            "登录企业微信",
            "二维码已过期",
        )
        texts = set()
        try:
            for control in window.descendants(control_type="Text"):
                text = (control.window_text() or "").strip()
                if text:
                    texts.add(text)
            # 部分版本侧边栏是 Button，不是 Text，补充采集提升识别率。
            for control in window.descendants(control_type="Button"):
                text = (control.window_text() or "").strip()
                if text:
                    texts.add(text)
        except Exception:
            pass

        if title:
            texts.add(title.strip())
        self.last_detected_texts = sorted(texts)

        if any(marker in title for marker in unauth_markers):
            return "unauth"

        if any(marker in text for marker in unauth_markers for text in texts):
            return "unauth"

        main_markers = ("消息", "通讯录", "工作台", "文档", "日程", "邮件", "客户联系")
        if any(marker in texts for marker in main_markers):
            return "logged_in"

        # 兜底：窗口存在且未出现登录/扫码关键词时，按已登录处理，避免扫码后识别不到。
        if texts:
            return "logged_in"

        return "unknown"

    def check_login_status(self):
        """检测是否已登录企业微信，并更新 is_login/login_time 状态。"""
        phase = self._detect_auth_phase()
        if phase == "logged_in":
            self.is_login = True
            if not self.login_time:
                self.login_time = datetime.now()
            return True

        # 未命中主界面特征时，继续按未登录处理。
        self.is_login = False
        return False

    def wait_scan_login(self, timeout=300, interval=2):
        """在超时时间内循环检测登录状态，等待用户扫码完成登录。"""
        start_time = time.time()
        self.login_state = "wait_scan"
        self.login_state_text = "客户端已拉起，等待扫码登录"
        unauth_seen = False
        unknown_after_scan_streak = 0
        while time.time() - start_time <= timeout:
            # 等待期间若客户端被关闭，则自动重新拉起。
            if not self._is_process_running():
                self.open_client(show_window=True)
            elif not self._get_wxwork_window():
                self._show_client_window()

            phase = self._detect_auth_phase()
            if phase == "unauth":
                unauth_seen = True
                unknown_after_scan_streak = 0
                self.is_login = False
            elif phase == "logged_in":
                self.is_login = True
                if not self.login_time:
                    self.login_time = datetime.now()
                self.login_state = "scan_login_success"
                self.login_state_text = "扫码登录成功"
                print("检测到已登录，结束等待。", flush=True)
                return True
            else:
                # 某些版本扫码后窗口短暂不可读；若此前已确认在登录页，连续未知则视为已登录。
                if unauth_seen and self._is_process_running():
                    unknown_after_scan_streak += 1
                    if unknown_after_scan_streak >= 3:
                        self.is_login = True
                        if not self.login_time:
                            self.login_time = datetime.now()
                        self.login_state = "scan_login_success"
                        self.login_state_text = "扫码后状态切换成功（未知态兜底判定）"
                        print("检测到扫码后状态已切换，结束等待。", flush=True)
                        return True

            elapsed = int(time.time() - start_time)
            left = max(0, int(timeout - elapsed))
            print(
                f"等待扫码中... 已等待{elapsed}s，剩余{left}s，阶段={phase}，文本数={len(self.last_detected_texts)}",
                flush=True,
            )
            time.sleep(interval)
        self.is_login = False
        self.login_state = "scan_login_timeout"
        self.login_state_text = "等待扫码登录超时"
        print("等待扫码登录超时，请确认已在手机端完成扫码。")
        return False

    def ensure_login(self, timeout=300, interval=2):
        """统一登录入口：启动客户端并确保登录，必要时进入扫码等待。"""
        if self.check_login_status():
            self.login_state = "direct_login_success"
            self.login_state_text = "检测到已登录（无需扫码）"
            return True
        # 未登录时，强制弹出企业微信客户端窗口，方便扫码。
        if not self.open_client(show_window=True):
            self.is_login = False
            self.login_state = "launch_failed"
            self.login_state_text = "企业微信客户端拉起失败"
            return False
        print("检测到未登录，请在企业微信客户端扫码登录...")
        return self.wait_scan_login(timeout=timeout, interval=interval)

    def run_login_test(self, timeout=120, interval=2, print_text_limit=20):
        """执行企业微信登录测试并输出完整状态信息。"""
        success = False
        print("开始执行企业微信登录检测...", flush=True)
        try:
            success = self.ensure_login(timeout=timeout, interval=interval)
        except KeyboardInterrupt:
            print("检测被手动中断，输出当前状态：", flush=True)
        finally:
            print(f"企业微信登录状态: {success}", flush=True)
            print(f"登录流程状态码: {self.login_state}", flush=True)
            print(f"登录流程说明: {self.login_state_text}", flush=True)
            print(f"实例状态 is_login: {self.is_login}", flush=True)
            print(f"登录时间 login_time: {self.login_time}", flush=True)
            print(
                f"最近检测到文本(前{print_text_limit}条): {self.last_detected_texts[:print_text_limit]}",
                flush=True,
            )
        return {
            "success": success,
            "login_state": self.login_state,
            "login_state_text": self.login_state_text,
            "is_login": self.is_login,
            "login_time": self.login_time,
            "last_detected_texts": self.last_detected_texts[:print_text_limit],
        }
