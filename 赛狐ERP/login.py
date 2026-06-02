import os
import time
import base64
from pathlib import Path

import ddddocr
from DrissionPage import ChromiumPage


class SaihuERPLogin:
    def __init__(self, page: ChromiumPage, username: str, password: str, img_dir: str):
        self.page = page
        self.username = username
        self.password = password
        self.img_dir = Path(img_dir)

    def login(self, max_retry: int = 7) -> bool:
        # 先访问业务页，判断是否已登录
        self.page.get("https://www.sellfox.com/amzup-web-main/web/purchase/purchaseManage/index.html")
        need_login = self.page.ele('x://div[text()="免费使用"]', timeout=3)

        if not need_login:
            # 已登录态（或页面无需登录）
            return True

        # 进入登录页
        self.page.get("https://www.sellfox.com/amzup-web-main/login.html")
        self.page.ele('x://input[@id="username"]', timeout=8).input(self.username, clear=True)
        self.page.ele('x://input[@id="password"]', timeout=8).input(self.password, clear=True)

        for _ in range(max_retry):
            # 读取验证码 base64 数据
            src = self.page.ele('x://div[@class="login_vcode"]/a/img', timeout=8).attr("src")
            if not src or "," not in src:
                raise RuntimeError("未获取到验证码图片数据")

            img_bs4 = src.split(",", 1)[1]
            code = self.img(self.img_dir, img_bs4)

            # 输入验证码
            self.page.ele('x://*[@placeholder="请输入图形验证码"]', timeout=8).input(code, clear=True)

            # 勾选协议（若未勾选）
            checkbox_label = self.page.ele('@class=el-checkbox center_align', timeout=3)
            if checkbox_label and "is-checked" not in (checkbox_label.attr("class") or ""):
                self.page.ele('x://span[contains(text(), "阅读并接受")]/preceding-sibling::*', timeout=5).click()

            # 点击登录
            self.page.ele('x://button/*[text()="登录"]', timeout=8).click()
            time.sleep(5)

            # 校验是否登录成功
            ok = self.page.ele('x://span[text()="商品"]', timeout=5)
            if ok:
                return True

        return False

    def img(self, img_path: Path, img_data: str) -> str:
        img_path.mkdir(parents=True, exist_ok=True)
        img_file = img_path / "output_image.png"

        with open(img_file, "wb") as f:
            f.write(base64.b64decode(img_data))

        img_bytes = img_file.read_bytes()
        ocr = ddddocr.DdddOcr()
        result = ocr.classification(img_bytes).strip()

        if img_file.exists():
            os.remove(img_file)

        return result


if __name__ == "__main__":
    # 示例用法
    page = ChromiumPage()
    login_client = SaihuERPLogin(
        page=page,
        username="CW-3",
        password="Bns123456",
        img_dir=str(Path(__file__).resolve().parent),
    )
    success = login_client.login(max_retry=7)
    print(f"登录结果: {'成功' if success else '失败'}")