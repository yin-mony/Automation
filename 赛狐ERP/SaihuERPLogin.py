import base64
import os
import subprocess
import sys
import time


class SaihuERPLogin:
    DEFAULT_USERNAME = "zidonghua"
    DEFAULT_PASSWORD = "Aiworkds123."
    LOGIN_URL = "https://www.sellfox.com/amzup-web-main/login.html"
    ENTRY_URL = "https://www.sellfox.com/amzup-web-main/web/dashboard.html"
    LOGIN_CHECK_TIMEOUT = 1.2
    MANUAL_WAIT_TIMEOUT = 120
    _ocr_engine = None
    _ocr_checked = False
    _ocr_error = ""
    _ocr_warned = False
    _ocr_repair_attempted = False

    def __init__(self, page, username=None, password=None, img_dir=None):
        self.page = page
        self.username = username or self.DEFAULT_USERNAME
        self.password = password or self.DEFAULT_PASSWORD
        self.img_dir = img_dir or os.getcwd()

    def login(self):
        print("开始执行赛狐 ERP 登录流程...", flush=True)

        # 1) 每次优先尝试进入业务页；进入成功则直接使用当前会话。
        if self._enter_entry_once_and_check():
            print("检测到账号未退出，直接进入赛狐业务页面。", flush=True)
            return True
        print("ENTRY_URL 未能直接进入业务页，回到登录页执行登录。", flush=True)

        # 2) 回到登录页，按 Qt 传入账号密码执行登录。
        self.page.get(self.LOGIN_URL)
        username_input = self.page.ele('x://input[@id="username"]', timeout=6)
        password_input = self.page.ele('x://input[@id="password"]', timeout=6)
        if not username_input or not password_input:
            raise RuntimeError(
                f"未检测到登录输入框，当前页面: {self.page.url}，请确认是否有弹窗或页面结构变化。"
            )

        username_input.input(self.username, clear=True)
        password_input.input(self.password, clear=True)

        for index in range(7):
            print(f"登录尝试第 {index + 1} 次...", flush=True)
            captcha_img = self.page.ele('x://div[@class="login_vcode"]/a/img', timeout=5)
            if not captcha_img:
                if self._is_logged_in():
                    print("检测到已进入系统，跳过验证码环节。", flush=True)
                    return True
                raise RuntimeError(f"未找到验证码图片，当前页面: {self.page.url}")

            captcha_input = self.page.ele('x://*[@placeholder="请输入图形验证码"]', timeout=5)
            if captcha_input:
                captcha_input.click()
            captcha_code = self._solve_captcha_with_retry(captcha_img)
            if captcha_code and captcha_input:
                captcha_input.input(captcha_code, clear=True)
                print(f"验证码自动识别并填入: {captcha_code}", flush=True)
            checkbox_label = self.page.ele("@class=el-checkbox center_align")

            if checkbox_label:
                # 检查是否已选中（通过类名判断）
                if "is-checked" not in checkbox_label.attr("class"):
                    print("复选框未选中，准备点击", flush=True)
                    self.page.ele('x://span[contains(text(), "阅读并接受")]/preceding-sibling::*').click()

            if captcha_code:
                self.page.ele('x://button/*[text()="登录"]').click()
                if self._wait_for_login_success(timeout=6):
                    print("赛狐 ERP 登录成功。", flush=True)
                    return True
                print("自动识别验证码登录未成功，切换手动验证码流程。", flush=True)

            print(
                f"请在页面手动输入验证码并点击“登录”，当前等待 {self.MANUAL_WAIT_TIMEOUT} 秒...",
                flush=True,
            )
            if self._wait_manual_login_result(timeout=self.MANUAL_WAIT_TIMEOUT):
                print("赛狐 ERP 登录成功。", flush=True)
                return True

            print("本次等待超时或登录未成功，刷新验证码后重试。", flush=True)
            captcha_img.click(by_js=True)
            time.sleep(0.4)

        raise RuntimeError("赛狐 ERP 登录失败：连续尝试后仍未进入系统。")

    def _is_logged_in(self):
        return bool(self.page.ele('x://span[text()="商品"]', timeout=self.LOGIN_CHECK_TIMEOUT))

    def _enter_entry_once_and_check(self):
        self.page.get(self.ENTRY_URL)
        return self._is_logged_in()

    def _wait_manual_login_result(self, timeout):
        end_time = time.time() + max(timeout, 1)
        while time.time() < end_time:
            if self._is_logged_in():
                return True
            time.sleep(1)
        return False

    def _wait_for_login_success(self, timeout):
        end_time = time.time() + max(timeout, 1)
        while time.time() < end_time:
            if self._is_logged_in():
                return True
            time.sleep(0.6)
        return False

    def _solve_captcha_with_retry(self, captcha_img):
        ocr_engine = self._get_ocr_engine()
        if ocr_engine is None:
            if not self.__class__._ocr_warned:
                self.__class__._ocr_warned = True
                print(
                    "验证码自动识别组件不可用，已自动切换为手动输入验证码。"
                    f"原因: {self.__class__._ocr_error}",
                    flush=True,
                )
            return ""

        for _ in range(2):
            src = captcha_img.attr("src") or ""
            if "," not in src:
                captcha_img.click(by_js=True)
                time.sleep(0.4)
                continue

            img_bs64 = src.split(",", 1)[1]
            try:
                img_bytes = base64.b64decode(img_bs64)
                result = ocr_engine.classification(img_bytes)
                result = "".join(ch for ch in str(result).strip() if ch.isalnum())
                if result:
                    return result[:6]
            except Exception:
                # 本轮识别失败，刷新验证码重试
                pass

            captcha_img.click(by_js=True)
            time.sleep(0.4)
        return ""

    def _get_ocr_engine(self):
        if self.__class__._ocr_checked:
            return self.__class__._ocr_engine

        self.__class__._ocr_checked = True
        engine = self._try_load_ocr_engine()
        if engine is not None:
            return engine

        # 自动修复一次依赖环境后再重试，尽量保证自动识别可用。
        if not self.__class__._ocr_repair_attempted:
            self.__class__._ocr_repair_attempted = True
            print("检测到验证码识别组件异常，尝试自动修复依赖后重试...", flush=True)
            self._try_auto_repair_ocr_env()
            engine = self._try_load_ocr_engine()
            if engine is not None:
                print("验证码识别组件自动修复成功。", flush=True)
                return engine

        self.__class__._ocr_engine = None
        return None

    def _try_load_ocr_engine(self):
        try:
            import ddddocr

            self.__class__._ocr_engine = ddddocr.DdddOcr()
            self.__class__._ocr_error = ""
            return self.__class__._ocr_engine
        except Exception as exc:
            self.__class__._ocr_error = str(exc)
            return None

    def _try_auto_repair_ocr_env(self):
        try:
            subprocess.run(
                [
                    sys.executable,
                    "-m",
                    "pip",
                    "install",
                    "--upgrade",
                    "--force-reinstall",
                    "onnxruntime",
                    "ddddocr",
                ],
                check=False,
                capture_output=True,
                text=True,
            )
        except Exception:
            # 修复失败时保持静默，后续走手动验证码兜底。
            pass