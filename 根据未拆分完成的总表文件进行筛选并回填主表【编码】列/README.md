# 根据未拆分完成的总表文件进行筛选并回填主表【编码】列

## 用途

在总表与副表需要按订单号对齐的场景下，执行以下自动化流程：

- 读取用户选择的主表与副表 Excel 文件；
- 用主表 `描述` 与副表 `myp_order_id` 做完全匹配；
- 将匹配成功的 `asin` 聚合后回填到主表 `编码（必填）`（或兼容 `编码(必填)`）列；
- 支持多个 asin 用英文逗号拼接写入。

## 当前实现状态

- 已完成核心主流程封装：`Filter_add.py`
- 已完成双入口启动：`run.py`
  - CLI 模式（交互输入文件路径）
  - GUI 模式（Qt 选择文件并执行）
- 已支持执行成功提示弹窗与常见路径/写入异常提示

## 环境要求

- Windows 10/11（亦可在其他系统运行源码，打包脚本以 Windows 为主）。
- Python 3.10+（若从源码运行）。

## 安装依赖（源码运行 / 打包前）

在**本目录**打开终端，执行：

```bash
pip install -r requirements.txt
```

主要依赖：`pandas`、`openpyxl`、`PySide6`；打包 exe 还需 `pyinstaller`（已写在 `requirements.txt` 中）。

## 文件说明

| 文件 | 说明 |
|------|------|
| `Filter_add.py` | 主逻辑模块（读取、校验、匹配、回填、写回） |
| `run.py` | 应用入口（`--mode cli` / `--mode gui`） |
| `requirements.txt` | Python 依赖列表 |
| `run_gui.bat` | 源码运行 GUI：先安装依赖再启动（适合新环境双击） |
| `build_exe.bat` | 一键调用 PyInstaller 打包 |
| `匹配回填工具.spec` | PyInstaller 规格文件，便于复现打包参数 |
| `test.py` | 测试草稿，仅用于本地验证，不作为正式入口 |
| `README.md` | 本说明 |

## 使用方式

### 1) GUI 模式（默认）

**方式 A（推荐，新环境）：** 双击 **`run_gui.bat`**（会自动按 `requirements.txt` 安装依赖后启动界面）。

**方式 B：** 已装好依赖时：

```bash
python run.py --mode gui
```

操作步骤：

1. 选择主表文件；
2. 选择副表文件；
3. 点击“开始执行”；
4. 程序完成匹配和回填后会弹窗提示成功，并将结果写回主表文件。

### 2) CLI 模式

```bash
python run.py --mode cli
```

按提示输入：

- 主表路径（必填）
- 副表路径（必填）
- 输出路径（可留空，留空则覆盖主表文件）

## 匹配与回填规则

- 匹配键：
  - 主表列：`描述`
  - 副表列：`myp_order_id`
- 匹配方式：去除前后空格后，完全匹配
- 回填列：
  - 优先写入 `编码（必填）`
  - 若主表仅存在 `编码(必填)`，自动兼容写入该列
- 同一 `myp_order_id` 多个 `asin`：按出现顺序逗号拼接（`,`）

## 打包为 exe（Windows）

双击或在命令行运行本目录下的 **`build_exe.bat`**。完成后在 `dist` 目录可得到 **`匹配回填工具.exe`**（单文件、无控制台窗口）。

也可使用已生成的 `匹配回填工具.spec`：

```bash
pyinstaller 匹配回填工具.spec
```

说明：`dist/`、`build/` 体积较大，已通过仓库根目录 `.gitignore` 忽略，不要提交到 Git。

---

更上层仓库索引请参见根目录 [`README.md`](../README.md)。
