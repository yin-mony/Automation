import os
import time
import base64
from pathlib import Path

import ddddocr
from DrissionPage import ChromiumPage


# 赛狐通用登录类
class SaiHuERPLogin:
    """赛狐 ERP 登录、验证码识别与公告关闭"""

    def __init__(self, config):
        self.config = config
        self.page = config["page"]
        self.username = config["username"]
        self.password = config["password"]
        self.img_path = Path(config.get("img_path") or config.get("base_dir") or Path(__file__).resolve().parent)

    def login(self):
        """进入赛狐页面，未登录时自动输入账号密码并处理验证码"""
        self.page.get('https://www.sellfox.com/amzup-web-main/web/purchase/purchaseManage/index.html')
        login = self.page.ele('x://div[text()="免费使用"]')
        if login:
            self.page.get('https://www.sellfox.com/amzup-web-main/login.html')
            self.page.ele('x://input[@id="username"]').input(f'{self.username}', clear=True)
            self.page.ele('x://input[@id="password"]').input(f'{self.password}', clear=True)
            for index in range(7):
                img_bs4 = self.page.ele('x://div[@class="login_vcode"]/a/img').attr('src').split(",")[1]
                img_url = self.img_code(img_bs4)
                self.page.ele('x://*[@placeholder="请输入图形验证码"]').input(img_url, clear=True)
                checkbox_label = self.page.ele('@class=el-checkbox center_align')

                if checkbox_label:
                    if 'is-checked' not in checkbox_label.attr('class'):
                        print("复选框未选中，准备点击")
                        self.page.ele('x://span[contains(text(), "阅读并接受")]/preceding-sibling::*').click()

                self.page.ele('x://button[contains(., "登录")]').click()
                time.sleep(5)
                login = self.page.ele('x://span[text()="商品"]', timeout=5)
                if login:
                    inform = self.page.ele('x://div[@class="title_container"]')
                    if inform:
                        next_button = self.page.ele('x://div[@class="announcement_dialog_footer"]/button/span')
                        if next_button and next_button.text == '下一条':
                            print("按钮文本是'下一条'")
                            for _ in range(20):
                                if not next_button:
                                    break
                                if next_button.text != '下一条':
                                    break
                                next_button.click()
                                time.sleep(0.5)
                                next_button = self.page.ele('x://div[@class="announcement_dialog_footer"]/button/span')
                            close_button = self.page.ele('x://div[@class="announcement_dialog_footer"]/button/span[text()="关闭"]')
                            if close_button:
                                close_button.click()
                        else:
                            close_button = self.page.ele('x://div[@class="announcement_dialog_footer"]/button/span[text()="关闭"]')
                            if close_button:
                                close_button.click()
                    return True
            return False
        return True

    def img_code(self, img_data):
        """将 base64 验证码图片保存为临时文件并用 OCR 识别"""
        self.img_path.mkdir(parents=True, exist_ok=True)
        img_url = self.img_path / "output_image.png"
        with open(img_url, "wb") as image_file:
            image_file.write(base64.b64decode(img_data))
        img_path = Path(img_url)
        img_bytes = img_path.read_bytes()

        ocr = ddddocr.DdddOcr()
        result = ocr.classification(img_bytes)
        if img_url.exists():
            os.remove(img_url)
        return result


if __name__ == '__main__':
    config = {
        "page": ChromiumPage(),
        "username": os.getenv("SAIHU_USERNAME", ""),
        "password": os.getenv("SAIHU_PASSWORD", ""),
        "img_path": Path(__file__).resolve().parent,
    }
    login_config = SaiHuERPLogin(config)
    success = login_config.login()
    print(f"登录结果: {'成功' if success else '失败'}")
