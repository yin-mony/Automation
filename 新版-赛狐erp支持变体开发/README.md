# 新版 - 赛狐 ERP 支持变体开发

赛狐 ERP 浏览器自动化：支持 **纯新品（mode1）** 与 **横向变体（mode2）** 两种流程，读取 Excel 提醒表驱动 SKU 配对与变体操作。提供 Tkinter GUI 与 PyInstaller 打包脚本。

## 代码结构

```
新版-赛狐erp支持变体开发/
├── main.py              # CLI 入口，--mode mode1|mode2
├── run.py               # Tkinter GUI 入口（推荐）
├── NewSet.py            # 模式一：纯新品流程（NewSetPage）
├── Variant.py           # 模式二：横向变体流程（VariantPage）
├── SaihuERPLogin.py     # 赛狐 ERP 登录
├── requirements.txt
├── build.bat / build.spec
├── 新品sku配对+横向变体配对自动提醒.xlsx   # 业务提醒表模板
└── README.md
```

| 模式 | 类 | 说明 |
| --- | --- | --- |
| `mode1` | `NewSetPage` | 纯新品 SKU 配对与相关操作 |
| `mode2` | `VariantPage` | 横向变体配对流程 |

## 依赖

```bash
pip install -r requirements.txt
```

主要依赖：`DrissionPage`、`pandas`、`openpyxl`、`ddddocr`（验证码）。

## 运行

**图形界面（推荐）：**

```bash
python run.py
```

**命令行：**

```bash
python main.py --mode mode1 --username 账号 --password 密码 --path "Excel路径.xlsx"
python main.py --mode mode2 --username 账号 --password 密码 --path "Excel路径.xlsx"
```

## 打包 exe

```bash
build.bat
```

产物见 `dist/`（需与 `build.spec` 中配置一致）。

## 注意事项

- 默认 Excel 路径与账号在 `main.py` 的 `DEFAULT_CONFIG` 中，GUI 可覆盖；**勿将真实密码推送到公开仓库**。
- 原 `LowPrice.py` 已合并/替换为 `Variant.py` 变体流程，低价表逻辑仍保留在 [`赛狐ERP/`](../赛狐ERP/) 子项目中。
