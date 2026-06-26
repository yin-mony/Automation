# 1688 商品链接数据采集

从 Excel 读取 1688 商品链接，用 DrissionPage 打开详情页并监听接口响应，保存 JSON 并汇总为规格 Excel。

## 代码结构

| 文件 | 职责 |
| --- | --- |
| `excel.py` | 读取 Excel 各工作表的「描述」「链接」列 |
| `main.py` | `Ali1688` 核心采集：监听 XHR/JSONP、按工作表落盘、汇总导出 |
| `mt.py` | 多标签批采辅助 |
| `run.py` | PyQt5 GUI 入口，配置路径并后台执行采集 |
| `test.py` | 浏览器采集调试脚本 |
| `api_1688.py` | 1688 开放平台官方 API 示例（OAuth + 商品查询，需环境变量配置密钥） |
| `run_config.json` | GUI 上次使用的本地路径（可随环境修改） |

## 依赖

- Python 3.10+
- DrissionPage、pandas、openpyxl
- GUI：`PyQt5`
- 官方 API 示例：`requests`

## 运行

### 浏览器采集（GUI）

```bash
python run.py
```

在界面填写 Excel 路径与 JSON 保存目录，点击开始运行。

### 官方 Open API 示例

需先在 1688 开放平台配置应用，并设置环境变量：

```bash
set ALI1688_APP_KEY=你的AppKey
set ALI1688_APP_SECRET=你的AppSecret
python api_1688.py
```

## 说明

- 采集 JSON 默认保存在脚本目录下按工作表命名的子文件夹中
- `.1688_token.json`、`chrome_profiles/` 等本地运行时文件已在 `.gitignore` 中排除，勿提交 Git
