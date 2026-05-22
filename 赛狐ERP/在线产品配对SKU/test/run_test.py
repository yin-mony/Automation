from pathlib import Path
import sys

CURRENT_DIR = Path(__file__).resolve().parent
ROOT_DIR = CURRENT_DIR.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.append(str(ROOT_DIR))

from EdgeRun import EdgeBrowserRunner
from SaihuERPLogin import SaihuERPLogin


def main():
    print("步骤1：启动并连接本地 Edge 浏览器...", flush=True)
    page = EdgeBrowserRunner.start_edge_and_connect(
        debug_port=EdgeBrowserRunner.DEFAULT_DEBUG_PORT,
        start_url=SaihuERPLogin.ENTRY_URL,
        fresh_profile=False,
        wait_seconds=3,
    )
    print(f"当前页面: {page.url}", flush=True)

    print("步骤2：执行赛狐 ERP 登录流程...", flush=True)
    login_client = SaihuERPLogin(page)
    login_client.login(force_relogin=False)
    print("步骤完成：已完成 Edge 拉起并执行赛狐 ERP 登录。", flush=True)


if __name__ == "__main__":
    main()
