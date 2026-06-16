import argparse
from pathlib import Path

from DrissionPage import ChromiumPage

from NewSet import NewSetPage
from Variant import VariantPage

DEFAULT_CONFIG = {
    "mode": "mode1",
    "username": "zidonghua",
    "password": "",
    "excel_path": r"C:\Users\admin\Desktop\新品sku配对+横向变体配对自动提醒.xlsx",
}


def run_mode(config: dict) -> None:
    page = ChromiumPage()
    mode = config["mode"]
    username = config["username"]
    password = config["password"]
    excel_path = config.get("excel_path") or None

    if mode == "mode1":
        runner = NewSetPage(page, username, password, excel_path)
    else:
        runner = VariantPage(page, username, password, excel_path)
    runner.main()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="赛狐ERP主流程入口")
    parser.add_argument(
        "--mode",
        choices=["mode1", "mode2"],
        default=DEFAULT_CONFIG["mode"],
        help="运行模式：mode1=纯新品，mode2=横向变体",
    )
    parser.add_argument(
        "--username",
        default=DEFAULT_CONFIG["username"],
        help="赛狐账号",
    )
    parser.add_argument(
        "--password",
        default=DEFAULT_CONFIG["password"],
        help="赛狐密码",
    )
    parser.add_argument(
        "--path",
        default=DEFAULT_CONFIG["excel_path"],
        help="Excel 文件路径",
    )
    args = parser.parse_args()

    config = {
        "mode": args.mode,
        "username": args.username,
        "password": args.password,
        "excel_path": args.path or DEFAULT_CONFIG["excel_path"],
    }
    run_mode(config)
