import base64
import os
import subprocess
import sys
import time

try:
    import ddddocr

    _DDDDOCR_IMPORT_ERROR = ""
except Exception as exc:
    ddddocr = None
    _DDDDOCR_IMPORT_ERROR = str(exc)

_PRELOADED_OCR_ENGINE = None
_PRELOADED_OCR_ERROR = ""
if ddddocr is None:
    _PRELOADED_OCR_ERROR = _DDDDOCR_IMPORT_ERROR or "ddddocr 导入失败"
else:
    try:
        # 模块加载时预初始化，避免每次登录流程里重复初始化。
        _PRELOADED_OCR_ENGINE = ddddocr.DdddOcr(show_ad=False, beta=True)
    except Exception as exc:
        _PRELOADED_OCR_ENGINE = None
        _PRELOADED_OCR_ERROR = str(exc)


class SaihuERPLogin:
    DEFAULT_USERNAME = "zidonghua"
    DEFAULT_PASSWORD = "Aiworkds123."
    PORTAL_URL = "https://www.sellfox.com/"
    LOGIN_URL = "https://www.sellfox.com/amzup-web-main/login.html"
    ENTRY_URL = "https://www.sellfox.com/amzup-web-main/web/dashboard.html"
    LOGIN_CHECK_TIMEOUT = 1.2
    ENTRY_WAIT_TIMEOUT = 8
    MANUAL_WAIT_TIMEOUT = 120
    _ocr_engine = _PRELOADED_OCR_ENGINE
    _ocr_checked = _PRELOADED_OCR_ENGINE is not None
    _ocr_error = _PRELOADED_OCR_ERROR
    _ocr_warned = False
    _ocr_repair_attempted = False

    def __init__(self, page, username=None, password=None, img_dir=None):
        self.page = page
        self.username = username or self.DEFAULT_USERNAME
        self.password = password or self.DEFAULT_PASSWORD
        self.img_dir = img_dir or os.getcwd()

    def login(self, prefer_entry_check=True):
        print("开始执行赛狐 ERP 登录流程...", flush=True)

        # 1) 可选：优先尝试进入业务页；进入成功则直接使用当前会话。
        if prefer_entry_check:
            if self._enter_entry_once_and_check():
                print("检测到账号未退出，直接进入赛狐业务页面。", flush=True)
                time.sleep(3)
                return True
            print("ENTRY_URL 未能直接进入业务页，回到登录页执行登录。", flush=True)
        else:
            print("按新实例流程：先进入官网首页，再进入登录页。", flush=True)

        # 2) 回到登录页，按 Qt 传入账号密码执行登录。
        self._open_login_page()
        username_input = self.page.ele('x://input[@id="username"]', timeout=6)
        password_input = self.page.ele('x://input[@id="password"]', timeout=6)
        if not username_input or not password_input:
            # 若登录页被自动会话重定向到业务页，直接视为成功。
            if self._is_logged_in(timeout=2):
                print("检测到已自动进入系统，跳过登录输入。", flush=True)
                return True
        if not username_input or not password_input:
            raise RuntimeError(
                f"未检测到登录输入框，当前页面: {self.page.url}，请确认是否有弹窗或页面结构变化。"
            )

        # 用户名输入
        username_input.clear()
        for char in self.username:
            username_input.input(char)
            time.sleep(0.2)
        time.sleep(0.5)

        # 密码输入
        password_input.clear()
        for char in self.password:
            password_input.input(char)
            time.sleep(0.2)
        time.sleep(0.5)

        self._ensure_auto_login_checked()

        # ====== 修改：持续自动识别验证码循环 ======
        auto_attempts = 0
        max_auto_attempts = 20  # 最多自动尝试20次

        while auto_attempts < max_auto_attempts:
            auto_attempts += 1
            print(f"自动识别验证码尝试第 {auto_attempts} 次...", flush=True)

            # 获取验证码图片
            captcha_img = self.page.ele('x://div[@class="login_vcode"]/a/img', timeout=5)
            if not captcha_img:
                if self._is_logged_in():
                    print("检测到已进入系统，跳过验证码环节。", flush=True)
                    return True
                # 如果找不到验证码图片，刷新页面重试
                print("未找到验证码图片，刷新页面...", flush=True)
                self.page.refresh()
                time.sleep(2)
                continue

            # 获取验证码输入框
            captcha_input = self.page.ele('x://*[@placeholder="请输入图形验证码"]', timeout=5)
            if captcha_input:
                captcha_input.click()

            # 识别验证码
            captcha_code = self._solve_captcha_with_retry(captcha_img)
            if not captcha_code:
                print("验证码识别失败，刷新验证码重试...", flush=True)
                captcha_img.click(by_js=True)
                time.sleep(0.5)
                continue

            print(f"识别到验证码: {captcha_code}", flush=True)

            # 填入验证码
            if captcha_input and captcha_code:
                fill_success = self._fill_captcha_ultimate(captcha_input, captcha_code)
                if fill_success:
                    print("验证码填入成功，准备点击登录...", flush=True)
                else:
                    print("验证码填入失败，刷新验证码重试...", flush=True)
                    captcha_img.click(by_js=True)
                    time.sleep(0.5)
                    continue

            # 点击登录
            time.sleep(1)  # 等待页面处理
            if self._click_login_ultimate():
                print("已点击登录按钮，等待登录结果...", flush=True)
                # 等待登录成功
                if self._wait_for_login_success(timeout=5):
                    print("赛狐 ERP 登录成功！", flush=True)
                    return True
                else:
                    print("登录失败，可能是验证码错误，刷新验证码重试...", flush=True)
                    # 刷新验证码
                    captcha_img.click(by_js=True)
                    time.sleep(0.5)
            else:
                print("点击登录按钮失败，刷新验证码重试...", flush=True)
                captcha_img.click(by_js=True)
                time.sleep(0.5)

            # 如果连续多次失败，重新加载登录页
            if auto_attempts % 5 == 0:
                print("连续多次失败，重新加载登录页...", flush=True)
                self.page.get(self.LOGIN_URL)
                time.sleep(2)
                # 重新输入用户名密码
                username_input = self.page.ele('x://input[@id="username"]', timeout=6)
                password_input = self.page.ele('x://input[@id="password"]', timeout=6)
                if username_input and password_input:
                    username_input.clear()
                    for char in self.username:
                        username_input.input(char)
                        time.sleep(0.2)
                    time.sleep(0.5)
                    password_input.clear()
                    for char in self.password:
                        password_input.input(char)
                        time.sleep(0.2)
                    time.sleep(0.5)
                    self._ensure_auto_login_checked()

        # ====== 如果自动尝试全部失败，进入手动模式 ======
        print(f"自动识别验证码尝试 {max_auto_attempts} 次均失败，切换手动输入...", flush=True)

        while True:
            print(
                f"请在页面手动输入验证码并点击“登录”，当前等待 {self.MANUAL_WAIT_TIMEOUT} 秒...",
                flush=True,
            )
            if self._wait_manual_login_result(timeout=self.MANUAL_WAIT_TIMEOUT):
                print("赛狐 ERP 登录成功。", flush=True)
                return True
            print("手动输入超时，刷新验证码后继续...", flush=True)
            captcha_img = self.page.ele('x://div[@class="login_vcode"]/a/img', timeout=3)
            if captcha_img:
                captcha_img.click(by_js=True)
                time.sleep(0.5)

        raise RuntimeError("赛狐 ERP 登录失败：连续尝试后仍未进入系统。")

    # ====== 终极版验证码填入 ======
    def _fill_captcha_ultimate(self, captcha_input, captcha_code):
        """
        终极版验证码填入方法
        """
        target = str(captcha_code or "").strip()
        if len(target) < 4:
            return False
        target = target[:4]

        try:
            # 1. 先清空
            captcha_input.clear()
            time.sleep(0.3)

            # 2. 逐字符输入，间隔长一点
            for char in target:
                captcha_input.input(char)
                time.sleep(0.5)
            time.sleep(0.5)

            # 3. JS强制写入
            self.page.run_js(f'document.querySelector(\'input[placeholder*="验证码"]\').value = "{target}";')
            time.sleep(0.3)
            self.page.run_js(f'document.querySelector(\'input[placeholder*="验证码"]\').dispatchEvent(new Event("input", {{bubbles: true}}));')
            time.sleep(0.3)
            self.page.run_js(f'document.querySelector(\'input[placeholder*="验证码"]\').dispatchEvent(new Event("change", {{bubbles: true}}));')

            # 4. 验证是否填入成功
            actual = self._get_captcha_input_value_enhanced()
            if actual == target:
                return True
            else:
                # 再试一次强制设置
                self.page.run_js(f'document.querySelector(\'input[placeholder*="验证码"]\').value = "{target}";')
                self.page.run_js(f'document.querySelector(\'input[placeholder*="验证码"]\').dispatchEvent(new Event("input", {{bubbles: true}}));')
                self.page.run_js(f'document.querySelector(\'input[placeholder*="验证码"]\').dispatchEvent(new Event("change", {{bubbles: true}}));')
                time.sleep(0.5)
                actual = self._get_captcha_input_value_enhanced()
                return actual == target

        except Exception as e:
            print(f"   终极版填入失败: {e}")
            return False

    # ====== 终极版登录按钮点击 ======
    def _click_login_ultimate(self):
        """
        终极版登录按钮点击
        """
        print("  正在点击登录按钮...", flush=True)

        # 方法1：直接点击按钮
        try:
            # 点击登录按钮前，先点击页面空白处，激活页面
            self.page.run_js('document.body.click();')
            time.sleep(0.3)

            # 多种方式定位按钮
            login_btn = self.page.ele('x://button[contains(., "登录")]', timeout=2)
            if login_btn:
                # 先模拟鼠标悬停
                login_btn.hover()
                time.sleep(0.3)
                login_btn.click()
                print("      ✅ 点击登录按钮成功", flush=True)
                return True
        except Exception:
            pass

        # 方法2：JS点击
        try:
            self.page.run_js('document.querySelector("button[type=\'submit\']")?.click();')
            time.sleep(0.3)
            self.page.run_js('document.querySelector("button.login-btn")?.click();')
            time.sleep(0.3)
            self.page.run_js('document.querySelector("button[class*=\'login\']")?.click();')
            print("      ✅ JS点击登录按钮成功", flush=True)
            return True
        except Exception:
            pass

        # 方法3：通过坐标点击
        try:
            login_btn = self.page.ele('x://button[contains(., "登录")]', timeout=2)
            if login_btn:
                x, y = login_btn.rect.center
                self.page.actions.move_to(x, y).click()
                print("      ✅ 坐标点击登录按钮成功", flush=True)
                return True
        except Exception:
            pass

        # 方法4：模拟回车键（可能触发登录）
        try:
            password_input = self.page.ele('x://input[@id="password"]', timeout=2)
            if password_input:
                password_input.click()
                time.sleep(0.3)
                self.page.actions.type('\n')  # 按回车
                print("      ✅ 模拟回车触发登录", flush=True)
                return True
        except Exception:
            pass

        print("      ❌ 终极版点击全部失败", flush=True)
        return False

    def _get_captcha_input_value_enhanced(self):
        """
        增强版获取输入框值
        """
        script = """
return (function(){
  try {
    var selectors = [
      'input[placeholder*="验证码"]',
      'input[placeholder*="图形验证码"]',
      'input[placeholder*="请输入验证码"]',
      'input[type="text"][autocomplete*="off"]'
    ];

    var el = null;
    for (var i = 0; i < selectors.length; i++) {
      el = document.querySelector(selectors[i]);
      if (el) break;
    }

    if (!el) return '';
    return String(el.value || '').trim();
  } catch(e) {
    return '';
  }
})();
"""
        try:
            return str(self.page.run_js(script) or "").strip()
        except Exception:
            return ""

    def _is_logged_in(self, timeout=None):
        check_timeout = self.LOGIN_CHECK_TIMEOUT if timeout is None else timeout
        return bool(self.page.ele('x://span[text()="商品"]', timeout=check_timeout))

    def _enter_entry_once_and_check(self):
        self.page.get(self.ENTRY_URL)
        state = self._wait_login_or_home(timeout=self.ENTRY_WAIT_TIMEOUT)
        return state == "home"

    def _open_login_page(self):
        self.page.get(self.PORTAL_URL)
        time.sleep(0.4)
        self.page.get(self.LOGIN_URL)
        if self.page.ele('x://input[@id="username"]', timeout=3):
            return
        # 少数情况下会被重定向到欢迎页/空白页，再次显式进入登录页。
        self.page.get(self.LOGIN_URL)

    def _wait_login_or_home(self, timeout):
        end_time = time.time() + max(timeout, 1)
        while time.time() < end_time:
            if self.page.ele('x://span[text()="商品"]', timeout=0.8):
                return "home"
            if self.page.ele('x://input[@id="username"]', timeout=0.8):
                return "login"
        return "unknown"

    def _ensure_auto_login_checked(self):
        """
        勾选“5天内自动登录”，保证会话未退出时下次可直接进入业务页。
        """
        label = self.page.ele('x://span[contains(text(), "5天内自动登录")]', timeout=1.2)
        if not label:
            return

        try:
            container = self.page.ele(
                'x://span[contains(text(), "5天内自动登录")]/ancestor::*[contains(@class, "el-checkbox")][1]',
                timeout=1.2,
            )
        except Exception:
            container = None

        try:
            container_class = (container.attr("class") if container else "") or ""
        except Exception:
            container_class = ""

        if "is-checked" in container_class:
            return

        # 优先点文案，再尝试点复选框方块。
        try:
            label.click(by_js=True)
            time.sleep(0.15)
        except Exception:
            pass

        try:
            checkbox_box = self.page.ele(
                'x://span[contains(text(), "5天内自动登录")]/preceding-sibling::*[contains(@class, "el-checkbox__input")]',
                timeout=1.2,
            )
            if checkbox_box:
                checkbox_box.click(by_js=True)
        except Exception:
            pass

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

        for _ in range(4):
            src = captcha_img.attr("src") or ""
            if "," not in src:
                captcha_img.click(by_js=True)
                time.sleep(0.4)
                continue

            img_bs64 = src.split(",", 1)[1]
            try:
                img_bytes = base64.b64decode(img_bs64)
                result = self._ocr_classify_best_effort(ocr_engine, img_bytes)
                if result:
                    return result
            except Exception:
                # 本轮识别失败，刷新验证码重试
                pass

            captcha_img.click(by_js=True)
            time.sleep(0.4)
        return ""

    def _ocr_classify_best_effort(self, ocr_engine, img_bytes):
        """
        使用 ddddocr 多策略识别同一张验证码：
        - 先尝试 png_fix=True（对透明背景/边缘锯齿更稳）
        - 再尝试默认模式
        最终统一做字符清洗，仅保留 4 位字母数字。
        """
        candidates = []
        for use_png_fix in (True, False):
            try:
                if use_png_fix:
                    raw = ocr_engine.classification(img_bytes, png_fix=True)
                else:
                    raw = ocr_engine.classification(img_bytes)
            except Exception:
                continue

            normalized = self._normalize_captcha_text(raw)
            if normalized:
                candidates.append(normalized)

        # 优先选择长度正好 4 位的结果
        for code in candidates:
            if len(code) == 4:
                return code
        return candidates[0] if candidates else ""

    def _normalize_captcha_text(self, text):
        code = "".join(ch for ch in str(text or "").strip() if ch.isalnum())
        code = code.upper()
        if len(code) < 4:
            return ""
        return code[:4]

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
            if ddddocr is None:
                self.__class__._ocr_error = _DDDDOCR_IMPORT_ERROR or "ddddocr 导入失败"
                return None

            # 关闭广告输出并使用 beta 模型，提升复杂验证码识别稳定性。
            self.__class__._ocr_engine = ddddocr.DdddOcr(show_ad=False, beta=True)
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


# ====== 测试代码 ======
if __name__ == "__main__":
    from DrissionPage import ChromiumPage

    # 创建浏览器页面
    page = ChromiumPage(9000)

    # 创建登录对象
    login = SaihuERPLogin(page)

    # 执行登录
    try:
        login.login(prefer_entry_check=True)
        print("✅ 登录成功！")
    except Exception as e:
        print(f"❌ 登录失败: {e}")