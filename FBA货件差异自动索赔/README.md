# FBA 货件差异自动索赔

本项目用于自动处理 Amazon FBA 货件差异索赔，当前拆分为两条独立流程：

- 赛狐流程：登录赛狐 ERP，筛选存在申收差异的 FBA 货件，读取货件详情，并按模板生成 POP PDF 文件。
- 易得客流程：通过易得客店铺环境进入 Amazon Seller Central，按 POP 文件中的货件编号上传材料、预览校验、提交 CASE 并读取问题编号。

两个流程共用同一个 GUI，但业务上是分开的。没有 POP 文件时，需要先运行赛狐流程生成 POP，再运行易得客流程提交索赔。

## 运行环境

- Windows 系统。
- Python 3.12 及以上版本。
- 本机需要可用的 Chrome 或 Edge，自动化流程通过 DrissionPage 控制浏览器。
- 赛狐流程如使用 DOCX 模板转 PDF，本机需要安装并可调用 Microsoft Word。
- 易得客流程需要本机已安装易得客客户端，并能正常打开店铺浏览器。

## 安装依赖

```bash
pip install -r requirements.txt
```

`requirements.txt` 仅保留当前子项目运行和打包需要的依赖，主要覆盖浏览器自动化、OCR、Word/PDF 处理、Windows 窗口控制、GUI 日期控件和 PyInstaller 打包。

## 启动方式

```bash
python run.py
```

GUI 入口为 `run.py`，窗口中包含：

- 公共配置区：运行环境、邮件通知、企业微信通知。
- 赛狐流程标签页：生成 POP PDF 文件。
- 易得客流程标签页：按货件来源上传发票、POP/POD，提交 CASE 并读取结果。

GUI 配置会缓存到 `run_config.json`。该文件可以删除，程序下次启动时会使用 `main.py` 中的默认配置回填，并在窗口关闭时重新保存。

## 当前架构

| 文件 | 职责 |
| --- | --- |
| `run.py` | 统一 GUI 入口，只负责界面、表单校验、日志输出和启动后台线程。 |
| `main.py` | 流程配置与调用入口，集中维护 GUI 默认值、站点列表、店铺列表、映射关系，并调用 `Saihu` / `Auto`。 |
| `saihu.py` | 赛狐流程业务逻辑，负责登录赛狐、筛选货件、获取详情、生成 POP、保存货件编号 JSON、按需发送邮件。 |
| `auto.py` | 易得客流程业务逻辑，负责连接店铺环境、进入 Amazon 后台、切换语言和站点、上传材料、预览校验、提交 CASE、读取问题编号。 |
| `export.py` | POP 文件导出逻辑，负责按 Word/PDF 模板生成最终 PDF，并处理单 SKU、多 SKU 表格样式。 |
| `email_util.py` | 邮件发送逻辑，负责 POP 附件邮件和 CASE 结果邮件，包含中文标题和正文编码处理。 |
| `wechat.py` | 企业微信机器人通知逻辑，负责汇总 CASE 结果并通过手机号 @ 指定人员。 |
| `SaihuERPLogin.py` | 赛狐 ERP 登录逻辑，包含验证码 OCR、登录状态处理、赛狐公告关闭。 |
| `YidekeLogin.py` | 易得客客户端和店铺浏览器连接逻辑，负责启动店铺环境并取得调试端口。 |
| `test.py` | 调试测试文件，保留用于局部验证；未确认同步前不作为正式流程入口。 |
| `tk_runtime_hook.py` | PyInstaller 运行时钩子，用于打包后设置 Tcl/Tk 资源路径，不需要由 `main.py` 手动调用。 |
| `FBA货件差异自动索赔.spec` | PyInstaller 打包配置，声明入口文件、隐藏导入、运行时钩子和内置资源。 |
| `build_exe.bat` | Windows 打包脚本，负责安装依赖、清理旧产物、调用 PyInstaller、复制 `run_config.json`。 |
| `docs/NAMING.md` | 当前子项目代码结构、命名和注释规范。 |
| `使用说明.md` | 面向实际操作人员的详细使用说明。 |

## 内置资源

| 文件 | 用途 |
| --- | --- |
| `服务商模板.docx` | 默认 POP Word 模板。 |
| `db53060fa183_发票模板.docx` | 历史发票模板资源，保留用于兼容。 |
| `AWD_POD.pdf` | 亚马逊分销货件使用的发票文件。 |
| `FBA_POD.pdf` | Send to Amazon 货件使用的库存所有权证明文件。 |

## 赛狐流程概览

1. 登录赛狐 ERP。
2. 进入 FBA 货件页面并切到产品维度。
3. 关闭公告并重置筛选条件。
4. 按 GUI 选择的站点、店铺、开始时间、结束时间筛选。
5. 筛选 `CLOSED(已完成)` 状态，并切换时间字段为更新时间。
6. 读取货件列表接口，按申收差异筛选需要生成 POP 的货件编号。
7. 打开货件详情接口，提取店铺、仓库、地址、SKU、FNSKU、数量等信息。
8. 按 GUI 选择的 Word/PDF 模板生成 POP PDF。
9. 保存 `shipment_ids.json`，记录本轮成功、失败和错误详情。
10. 如开启邮件通知，则发送本轮生成的 POP 附件。

## 易得客流程概览

1. 启动或连接易得客店铺浏览器。
2. 按 GUI 配置的店铺站点进入对应店铺。
3. 登录 Amazon Seller Central。
4. 切换页面语言为中文简体。
5. 按 GUI 配置的 Amazon 后台站点切换站点和账户。
6. 从 POP 目录读取 `shipment_ids.json`，没有 JSON 时兜底从 `*_POP.pdf` 文件名提取货件编号。
7. 逐个货件编号进入 Amazon 货件详情页。
8. 差值为 0 的货件跳过；已成功提交过的问题不重复提交。
9. 亚马逊分销货件只上传 `AWD_POD.pdf` 到发票入口。
10. Send to Amazon 货件上传当前货件对应的 POP PDF 到交货证明，并上传 `FBA_POD.pdf` 到库存所有权证明。
11. 点击上传文档并等待上传完成。
12. 点击预览，校验弹窗内 MSKU、差值和需要操作内容。
13. 校验通过后提交，并读取 CASE 问题编号。
14. 全部完成后汇总发送邮件和企业微信通知。

## 打包 exe

```bash
build_exe.bat
```

打包产物位于：

```text
dist/FBA货件差异自动索赔.exe
```

建议发布时发送整个 `dist` 目录，而不是只发送 exe。`dist/run_config.json` 会保存发布默认配置，用户启动后可直接看到默认账号、站点、目录和通知配置。

`tk_runtime_hook.py` 是 PyInstaller 的运行时钩子，只在打包后的 exe 启动阶段自动执行，用于修复 `tkinter` / `tkcalendar` 的 Tcl/Tk 资源路径；源码运行时不需要手动调用。

## 注意事项

- 运行赛狐流程前，确认赛狐账号、站点、店铺、时间范围和模板文件正确。
- 运行易得客流程前，确认 POP 目录中存在本轮生成的 POP PDF 和 `shipment_ids.json`。
- 不要在 Word 中打开模板或输出 DOCX/PDF 的同时运行导出流程，避免文件锁导致转换失败。
- 易得客流程按货件编号匹配 POP 文件，POP 文件名必须保留 FBA 货件编号。
- 邮件和企业微信通知均为可选项，建议先用少量货件验证结果，再批量运行。
