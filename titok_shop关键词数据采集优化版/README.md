# TikTok Shop 关键词数据采集（优化版）

基于 DrissionPage 访问 TikTok Shop，按关键词搜索商品、加载全部列表并遍历打印商品链接；支持人工滑动验证码等待与搜索关键词不一致时自动刷新重试。

## 代码结构

```
titok_shop关键词数据采集优化版/
├── test.py              # TikTok 主流程：搜索、加载更多、打印链接
├── TiTokshopLogin.py    # 登录与滑块验证码自动识别（备用模块）
└── README.md
```

| 文件 | 职责 |
| --- | --- |
| `test.py` | `TikTok` 类：人工验证码监测、关键词搜索、商品链接采集与打印 |
| `TiTokshopLogin.py` | `TikTokPage`：打开 shop.tiktok.com、自动/人工滑块验证码处理 |

## 环境要求

- Windows 10/11（需 Chrome/Chromium，由 DrissionPage 接管）
- Python 3.10+
- 依赖：`DrissionPage`、`openpyxl` 等（见各脚本 import）

## 运行

1. 修改 `test.py` 末尾 `config["files"]` 为 Excel 路径（写入功能启用时使用）。
2. 修改 `main()` 内 `codes` 为待搜索关键词。
3. 执行：

```bash
python test.py
```

出现滑动验证码时需人工完成；搜索框关键词与目标不一致时会刷新页面并重试。

## 当前状态

- 已实现：关键词搜索、加载更多、商品链接过滤与打印、人工验证码等待。
- 商品详情字段采集与 Excel 分 sheet 写入逻辑在 `test.py` 中预留，可按需恢复。

## 注意事项

- 脚本内 Excel 路径默认为本机桌面路径，部署前请改为实际位置。
- 请勿将含业务数据的 Excel 或账号信息提交到 Git。
