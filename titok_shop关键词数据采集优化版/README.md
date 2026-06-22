# TikTok Shop 关键词数据采集（优化版）

基于 DrissionPage 访问 TikTok Shop，从 Excel 读取关键词，按关键词搜索商品、采集详情并分关键词导出链接 txt、JSON 与 Excel。

## 代码结构

```
titok_shop关键词数据采集优化版/
├── main.py              # 主流程：获取链接 → 采集 JSON → 导出 Excel
├── exceldf.py           # 从 Excel「关键词表格」读取关键词列
├── TiTokshopLogin.py    # 登录与滑块验证码处理（默认人工滑动）
├── test.py              # 本地调试脚本
└── README.md
```

| 文件 | 职责 |
| --- | --- |
| `main.py` | `TikTok` 类：搜索、监听采集、按关键词导出 txt/JSON/Excel |
| `exceldf.py` | `ExcelDF`：读取配置 Excel 中「关键词表格」的「关键词」列 |
| `TiTokshopLogin.py` | `TikTokPage`：验证码检测、关闭弹窗、人工/自动滑动 |
| `test.py` | 简化调试：搜索、打印链接、人工验证码等待 |

## 环境要求

- Windows 10/11（需 Chrome/Chromium，由 DrissionPage 接管）
- Python 3.10+
- 依赖：`DrissionPage`、`openpyxl`、`pandas` 等（见各脚本 import）

## 运行

1. 准备 Excel，包含工作表 **「关键词表格」**，表头含 **「关键词」** 列。
2. 修改 `main.py` 末尾配置：

```python
config = {
    'file_path': r'...\tiktok-竞品信息抓取.xlsx',  # 关键词来源（与导出目录独立）
    'file': r'D:\Desktop',                        # 链接 txt / JSON / Excel 导出目录
}
```

3. 执行完整流程：

```bash
python main.py
```

流程分三步：获取商品链接 → 采集 JSON 数据 → 导出 Excel（每个关键词一个 sheet）。

## 导出文件说明

按关键词分别生成（文件名前缀为关键词安全化后的 `safe_name`）：

- `{safe_name}_商品链接.txt`
- `{safe_name}_商品数据.json`
- `{safe_name}_商品摘要.json`
- `TikTok_Shop_竞品数据.xlsx`（多关键词分 sheet）

## 注意事项

- `file_path`（读关键词）与 `file`（写导出）路径可不同，请按实际环境修改。
- 出现滑动验证码时需人工完成；`img_code/` 为运行时验证码图片缓存，勿提交 Git。
- 请勿将含业务数据的 Excel 或账号信息提交到 Git。
