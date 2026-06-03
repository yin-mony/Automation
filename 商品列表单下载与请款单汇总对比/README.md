# 商品列表单下载与请款单汇总对比

赛狐 ERP：自动下载采购单与商品列表单，按 SKU 合并对比采购单价，输出高亮差异的汇总 Excel。

## 代码结构

```
商品列表单下载与请款单汇总对比/
├── run.py                              # PyQt5 GUI 入口
├── main.py                             # 浏览器下载 + Excel 对比汇总
├── test.py                             # SaihuERPLogin 登录模块
├── 商品列表单下载与请款单汇总对比.spec
└── README.md
```

| 文件 | 职责 |
| --- | --- |
| `run.py` | `RunWindow`：选择下载目录；`Worker` 后台调用 `main.run(download_dir)` |
| `main.py` | `run()`：登录 → 导出采购单（模板「自动化检查采购单价」）→ 导出商品列表；`excel_file()` merge 比价并高亮 |
| `test.py` | `SaihuERPLogin`：赛狐登录与验证码处理（`main` 中 `from test import SaihuERPLogin`） |

### 配置说明

| 项 | 含义 |
| --- | --- |
| `download_dir` | GUI 传入或 `main.run()` 参数；存放 `采购单下载.xlsx`、`商品列表单下载.xlsx` 及输出对比表 |
| 输出文件 | `采购单与商品单汇总对比.xlsx` |
| 高亮规则 | 采购单价 > 商品列表单采购单价（或商品侧为空）整行标黄 |

## 依赖

- Python 3.10+
- PyQt5、DrissionPage、pandas、openpyxl、ddddocr 等（见本地环境）

## 运行

- **GUI（推荐）**：`python run.py`
- **命令行**：`python main.py`（使用 `main` 内默认路径，或通过 `run()` 传参）

## 打包

```bat
pyinstaller 商品列表单下载与请款单汇总对比.spec
```

说明：`build/`、`dist/` 已在仓库根 `.gitignore` 中忽略。
