import base64
import os
import time
from pathlib import Path


class SaihuERPLogin:
    DEFAULT_USERNAME = "zidonghua"
    DEFAULT_PASSWORD = "Aiworkds123."
    LOGIN_URL = "https://www.sellfox.com/amzup-web-main/login.html"
    ENTRY_URL = "https://www.sellfox.com/amzup-web-main/web/purchase/purchaseManage/index.html"
    LOGIN_CHECK_TIMEOUT = 1.2
    _ocr_engine = None

    def __init__(self, page, username=None, password=None, img_dir=None):
        self.page = page
        self.username = username or self.DEFAULT_USERNAME
        self.password = password or self.DEFAULT_PASSWORD
        self.img_dir = img_dir or os.getcwd()

    def login(self, force_relogin=False):
        print("开始执行赛狐 ERP 登录流程...", flush=True)
        file_path = self.img_dir
        # 固定先进入业务页检查登录状态：已登录则直接进入，不走强制登录。
        self.page.get(self.ENTRY_URL)
        is_logged_in = self.page.ele('x://span[text()="商品"]', timeout=self.LOGIN_CHECK_TIMEOUT)
        if is_logged_in and not force_relogin:
            print("检测到账号未退出，直接进入赛狐业务页面。", flush=True)
            return True
        if force_relogin:
            print("已启用强制重新登录：将执行登录流程。", flush=True)
        else:
            print("检测到账号已退出，进入登录页执行登录。", flush=True)

        # 未登录则强制进入登录页执行登录
        self.page.get(self.LOGIN_URL)
        username_input = self.page.ele('x://input[@id="username"]', timeout=6)
        password_input = self.page.ele('x://input[@id="password"]', timeout=6)

        # 复用已打开浏览器时，登录页可能被会话重定向到 dashboard，导致输入框不存在。
        if not username_input or not password_input:
            if self.page.ele('x://span[text()="商品"]', timeout=self.LOGIN_CHECK_TIMEOUT):
                print("当前会话已登录，未检测到登录输入框，跳过账号密码输入。", flush=True)
                return True
            raise RuntimeError(
                f"未检测到登录输入框，当前页面: {self.page.url}，请确认是否有弹窗或页面结构变化。"
            )

        username_input.input(self.username, clear=True)
        password_input.input(self.password, clear=True)
        for index in range(7):
            print(f"登录尝试第 {index + 1} 次...", flush=True)
            captcha_img = self.page.ele('x://div[@class="login_vcode"]/a/img', timeout=5)
            if not captcha_img:
                if self.page.ele('x://span[text()="商品"]', timeout=2):
                    print("检测到已进入系统，跳过验证码环节。", flush=True)
                    return True
                raise RuntimeError(f"未找到验证码图片，当前页面: {self.page.url}")

            img_bs4 = captcha_img.attr("src").split(",")[1]
            img_url = self.img(file_path, img_bs4)
            self.page.ele('x://*[@placeholder="请输入图形验证码"]').input(img_url, clear=True)
            checkbox_label = self.page.ele("@class=el-checkbox center_align")

            if checkbox_label:
                # 检查是否已选中（通过类名判断）
                if "is-checked" not in checkbox_label.attr("class"):
                    print("复选框未选中，准备点击", flush=True)
                    self.page.ele('x://span[contains(text(), "阅读并接受")]/preceding-sibling::*').click()

            self.page.ele('x://button/*[text()="登录"]').click()
            time.sleep(1.2)
            self.page.get(self.ENTRY_URL)
            if self.page.ele('x://span[text()="商品"]', timeout=5):
                print("赛狐 ERP 登录成功。", flush=True)
                return True

        raise RuntimeError("赛狐 ERP 登录失败：连续尝试后仍未进入系统。")

    def img(self, img_path, img_data):
        img_url = img_path + "\\" + "output_image.png"
        with open(img_url, "wb") as image_file:
            image_file.write(base64.b64decode(img_data))
        image_path = Path(img_url)
        img_bytes = image_path.read_bytes()

        if self.__class__._ocr_engine is None:
            try:
                import ddddocr
            except Exception as exc:
                os.remove(img_url)
                raise RuntimeError(
                    "验证码识别组件加载失败（ddddocr/onnxruntime）。"
                    "建议改用 Python 3.10-3.12 环境并重装相关依赖。"
                ) from exc
            self.__class__._ocr_engine = ddddocr.DdddOcr()

        result = self.__class__._ocr_engine.classification(img_bytes)  # 返回识别出的字母/数字
        os.remove(img_url)
        return result