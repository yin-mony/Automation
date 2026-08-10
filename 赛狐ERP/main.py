import argparse
import os
from pathlib import Path

from DrissionPage import ChromiumPage

from NewSet import NewSetPage
from LowPrice import LowPricePage


MODE_ONE = "mode_one"
MODE_TWO = "mode_two"


# 赛狐 ERP 统一脚本入口
class SaihuERP:
    def __init__(self, config):
        self.config = config
        self.mode = config.get("mode") or MODE_ONE
        self.page = config.get("page") or ChromiumPage()
        self.username = config["username"]
        self.password = config["password"]
        self.excel_path = config.get("excel_path") or ""
        self.base_dir = Path(config.get("base_dir") or Path(__file__).resolve().parent)

    def main(self):
        config = {
            "page": self.page,
            "username": self.username,
            "password": self.password,
            "excel_path": self.excel_path,
            "base_dir": self.base_dir,
        }

        if self.mode in (MODE_ONE, "mode1", "new_set"):
            run = NewSetPage(config)
            run.main()
            return

        if self.mode in (MODE_TWO, "mode2", "low_price"):
            run = LowPricePage(config)
            run.main()
            return

        raise ValueError(f"未知运行模式: {self.mode}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="赛狐ERP主流程入口")
    parser.add_argument(
        "--mode",
        choices=[MODE_ONE, MODE_TWO, "mode1", "mode2", "new_set", "low_price"],
        required=True,
        help="运行模式：mode_one/mode1=纯新品，mode_two/mode2=低价商城",
    )
    parser.add_argument("--username", default=os.getenv("SAIHU_USERNAME", ""), help="赛狐账号")
    parser.add_argument("--password", default=os.getenv("SAIHU_PASSWORD", ""), help="赛狐密码")
    parser.add_argument("--path", default="", help="Excel 文件路径")
    args = parser.parse_args()

    config = {
        "page": ChromiumPage(),
        "mode": args.mode,
        "username": args.username,
        "password": args.password,
        "excel_path": args.path,
        "base_dir": Path(__file__).resolve().parent,
    }
    run = SaihuERP(config)
    run.main()
