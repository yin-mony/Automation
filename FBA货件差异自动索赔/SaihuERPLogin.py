import os
import time
import base64
import ctypes
from pathlib import Path

import ddddocr
from DrissionPage import ChromiumPage



# 赛狐通用登录类
class SaiHuERPLogin:
    """赛狐 ERP 登录、验证码识别与公告关闭"""

    def __init__(self,page,username,password,img_path):
        # 外部传入页面实例，便于赛狐主流程复用同一个浏览器上下文
        self.page = page
        self.username = username
        self.password = password
        self.img_path = Path(img_path)

    def login(self) :
        """进入赛狐页面，未登录时自动输入账号密码并处理验证码"""
        # 先进入业务页判断登录态，出现免费使用按钮时说明需要重新登录
        self.page.get('https://www.sellfox.com/amzup-web-main/web/purchase/purchaseManage/index.html')
        login = self.page.ele('x://div[text()="免费使用"]')
        if login:
            self.page.get('https://www.sellfox.com/amzup-web-main/login.html')
            self.page.ele('x://input[@id="username"]').input(f'{self.username}',clear=True)
            self.page.ele('x://input[@id="password"]').input(f'{self.password}',clear=True)
            for index in range(7):
                # 每次失败后重新读取验证码图片并 OCR 识别
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
                    self.closeNotice()
                    return True
            return False
        self.closeNotice()
        return True

    def hasNotice(self):
        """判断当前页面是否存在可见的赛狐公告弹窗"""
        script = """
        const keywords = ['重要通知', '公告', '操作指引', '下一条', '最新活动'];
        const nodes = Array.from(document.querySelectorAll('.el-dialog__wrapper, .el-dialog, [class*="announcement"]'));
        return nodes.some(node => {
            const style = window.getComputedStyle(node);
            const rect = node.getBoundingClientRect();
            const visible = style.display !== 'none'
                && style.visibility !== 'hidden'
                && style.opacity !== '0'
                && rect.width > 0
                && rect.height > 0;
            return visible && keywords.some(keyword => (node.innerText || '').includes(keyword));
        });
        """
        try:
            return bool(self.page.run_js(script))
        except Exception:
            return False

    def clickNotice(self):
        """尝试点击赛狐公告弹窗中的下一条或关闭按钮"""
        script = """
        const keywords = ['重要通知', '公告', '操作指引', '下一条', '最新活动'];
        const labels = ['下一条', '关闭', '我知道了', '知道了', '确定', '完成'];
        const dialogs = Array.from(document.querySelectorAll('.el-dialog__wrapper, .el-dialog, [class*="announcement"]'))
            .filter(node => {
                const style = window.getComputedStyle(node);
                const rect = node.getBoundingClientRect();
                const visible = style.display !== 'none'
                    && style.visibility !== 'hidden'
                    && style.opacity !== '0'
                    && rect.width > 0
                    && rect.height > 0;
                return visible && keywords.some(keyword => (node.innerText || '').includes(keyword));
            });
        for (const dialog of dialogs) {
            const buttons = Array.from(dialog.querySelectorAll('button, span, i, [class*="close"]'));
            for (const label of labels) {
                const button = buttons.find(item => (item.innerText || '').trim().includes(label));
                if (button) {
                    button.click();
                    return label;
                }
            }
            const closeIcon = buttons.find(item => (item.className || '').toString().includes('close'));
            if (closeIcon) {
                closeIcon.click();
                return 'closeIcon';
            }
        }
        return '';
        """
        try:
            return self.page.run_js(script) or ""
        except Exception:
            return ""

    def waitNotice(self):
        """等待用户手动关闭赛狐公告弹窗"""
        print("检测到赛狐公告弹窗，自动关闭失败，等待用户手动关闭。", flush=True)
        ctypes.windll.user32.MessageBoxW(
            0,
            "检测到赛狐公告弹窗，但未找到可自动点击的关闭方式。\\n请手动关闭公告弹窗，关闭完成后点击“确定”，流程会继续等待页面恢复。",
            "赛狐公告处理",
            0x40 | 0x40000,
        )
        while self.hasNotice():
            time.sleep(1)
        print("赛狐公告弹窗已关闭，继续后续流程。", flush=True)

    def closeNotice(self):
        """登录完成后关闭赛狐公告弹窗，必要时等待人工关闭"""
        for _ in range(20):
            if not self.hasNotice():
                return True
            clicked = self.clickNotice()
            if clicked:
                print(f"已处理赛狐公告按钮: {clicked}", flush=True)
                time.sleep(1)
                continue
            self.waitNotice()
            return True
        if self.hasNotice():
            self.waitNotice()
        return True

    def img_code(self, img_data):
        """将 base64 验证码图片保存为临时文件并用 OCR 识别"""
        # ddddocr 读取本地图片字节后返回验证码文本
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
        username="sales25",
        password="Sales123...",
        img_path=r'C:\Users\admin\Desktop'
    )

    success = login_config.login()
    print(f"登录结果: {'成功' if success else '失败'}")
