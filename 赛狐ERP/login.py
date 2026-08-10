import os
from pathlib import Path

from DrissionPage import ChromiumPage

from SaihuERPLogin import SaiHuERPLogin


# 遗留登录入口，实际登录逻辑统一委托给 SaihuERPLogin.py
class SaihuERPLogin:
    def __init__(self, config):
        self.config = config
        self.login_config = SaiHuERPLogin(config)

    def login(self):
        return self.login_config.login()


if __name__ == "__main__":
    config = {
        "page": ChromiumPage(),
        "username": os.getenv("SAIHU_USERNAME", ""),
        "password": os.getenv("SAIHU_PASSWORD", ""),
        "img_path": Path(__file__).resolve().parent,
    }
    login_client = SaihuERPLogin(config)
    success = login_client.login()
    print(f"登录结果: {'成功' if success else '失败'}")
