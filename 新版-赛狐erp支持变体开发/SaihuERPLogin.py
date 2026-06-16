import os
import time
import base64
from pathlib import Path

import ddddocr
from DrissionPage import ChromiumPage



# 赛狐通用登录类
class SaiHuERPLogin:
    def __init__(self,page,username,password,img_path):
        self.page = page
        self.username = username
        self.password = password
        self.img_path = Path(img_path)

    def login(self) :
        self.page.get('https://www.sellfox.com/amzup-web-main/web/purchase/purchaseManage/index.html')
        login = self.page.ele('x://div[text()="免费使用"]')
        if login:
            self.page.get('https://www.sellfox.com/amzup-web-main/login.html')
            self.page.ele('x://input[@id="username"]').input(f'{self.username}',clear=True)
            self.page.ele('x://input[@id="password"]').input(f'{self.password}',clear=True)
            for index in range(7):
                img_bs4 = self.page.ele('x://div[@class="login_vcode"]/a/img').attr('src').split(",")[1]
                img_url = self.img_code(img_bs4)
                self.page.ele('x://*[@placeholder="请输入图形验证码"]').input(img_url, clear=True)
                checkbox_label = self.page.ele('@class=el-checkbox center_align')

                if checkbox_label:
                    # 检查是否已选中（通过类名判断）
                    if 'is-checked' not in checkbox_label.attr('class'):
                        print("复选框未选中，准备点击")
                        self.page.ele('x://span[contains(text(), "阅读并接受")]/preceding-sibling::*').click()
                self.page.ele('x://button[contains(., "登录")]').click()
                time.sleep(5)
                login = self.page.ele('x://span[text()="商品"]', timeout=5)
                if login:
                    inform = self.page.ele('x://div[@class="title_container"]')
                    # 检查是否弹出广告通知
                    if inform:
                        next_button = self.page.ele('x://div[@class="announcement_dialog_footer"]/button/span')
                        # 检查该广告是否只存在一个关闭按钮
                        if next_button and next_button.text == '下一条':
                            # 执行对应操作
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
                            next_button = self.page.ele('x://div[@class="announcement_dialog_footer"]/button/span[text()="关闭"]')
                            next_button.click()
                    return True
            return False
        return True

    def img_code(self, img_data):
        img_url = self.img_path / "output_image.png"
        with open(img_url, "wb") as image_file:
            image_file.write(base64.b64decode(img_data))
        img_path = Path(img_url)
        img_bytes = img_path.read_bytes()

        ocr = ddddocr.DdddOcr()
        result = ocr.classification(img_bytes)  # 返回识别出的字母/数字
        os.remove(img_url)
        return  result


if __name__ == '__main__':
    page = ChromiumPage()
    login_config = SaiHuERPLogin(
        page=page,
        username="zidonghua",
        password="",
        img_path=r'C:\Users\admin\Desktop'
    )

    success = login_config.login()
    print(f"登录结果: {'成功' if success else '失败'}")