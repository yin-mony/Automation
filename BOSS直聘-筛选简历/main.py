# 暂且使用 GUI 扫码登录入口；Web 控制台入口见下方注释
from run import RunGui

if __name__ == "__main__":
    app = RunGui()
    app.run()

# --- 原 Web 控制台入口（暂不使用）---
# from boss_web.server import main
# if __name__ == "__main__":
#     main()
