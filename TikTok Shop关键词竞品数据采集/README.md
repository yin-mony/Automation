# TikTok Shop 关键词竞品数据采集

基于 DrissionPage 自动化访问 TikTok Shop，按 Excel「关键词表格」中的关键词抓取竞品信息，并按关键词分工作表写回 Excel。

## 代码结构

```
TikTok Shop关键词竞品数据采集/
├── main.py              # TikTok 主流程：读关键词、抓取、写入 Excel
├── TiTokshopLogin.py    # TikTok Shop 登录与滑块验证码处理
├── DataExtract.py       # 从 Excel 提取关键词列表
├── test.py              # 本地调试脚本
├── 1.py                 # 商品页/API 请求分析辅助脚本
├── img_code/            # 验证码背景图与滑块图样例
├── request_analysis/    # 页面/JS 抓包分析产物（辅助开发）
└── README.md
```

| 文件 | 职责 |
| --- | --- |
| `main.py` | `TikTok` 类：读取关键词、循环搜索抓取、按关键词分 sheet 写入 |
| `TiTokshopLogin.py` | `TikTokPage`：打开 shop.tiktok.com、识别滑块验证码 |
| `DataExtract.py` | 从工作表「关键词表格」读取「关键词」列 |
| `1.py` | 分析 PDP 页面与接口，用于逆向请求参数 |

## 环境要求

- Windows 10/11（需 Chrome/Chromium，由 DrissionPage 接管）
- Python 3.10+
- 依赖：`DrissionPage`、`openpyxl`、`opencv-python`、`Pillow`、`requests` 等（见各脚本 import）

## 运行

1. 准备 Excel，包含工作表 **「关键词表格」**，表头含 **「关键词」** 列。
2. 修改 `main.py` 末尾 `config["files"]` 为 Excel 路径。
3. 执行：

```bash
python main.py
```

登录与验证码逻辑见 `TiTokshopLogin.py`；首次运行会在 `img_code/` 下保存验证码图片。

## 注意事项

- 脚本内路径默认为本机桌面路径，部署前请改为实际 Excel 位置。
- `request_analysis/` 为开发阶段抓包/分析文件，非运行时必需。
- 请勿将含业务数据的 Excel 或账号信息提交到 Git。
