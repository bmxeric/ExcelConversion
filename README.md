# ExcelConversion
对 execl 表格进行自动汇总

## Python 环境

项目使用 Python 3.13，虚拟环境位于 `.venv`，VS Code 已配置默认解释器。

首次搭建或重建环境（macOS）：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

依赖用途：

- `pandas`：后续用于表格数据整理、合并及差异比较。
- `openpyxl`：读取和写入 `.xlsx` 文件，后续用于差异单元格标识。
- `PySide6`：桌面窗口与文件选择，无需安装 Tk。

在 VS Code 中安装 Microsoft Python 扩展（`ms-python.python`）。如果解释器没有自动切换，运行命令 `Python: Select Interpreter`，选择项目中的 `.venv/bin/python`。新建终端后可运行 `python -c "import sys; print(sys.executable)"` 确认解释器路径。

## 数据目录

原始 Excel 文件位于 `execl/`。以 `~$` 开头的文件是 Excel 临时锁文件，不应参与后续处理。

## 运行程序

```bash
.venv/bin/python main.py
```

1. 点击“1. 选择 vitek.xlsx”，选择文件并指定新汇总文件的保存位置。程序自动创建 A–K 列表头，完成第一步汇总并保存。
2. 点击“2. 选择 lis.xlsx”，程序完成第二步匹配、解释转换及差异标记，更新同一个结果文件。
3. 点击“打开结果”查看完整文件。窗口预览前 200 行，输出文件保存全部结果。

输入文件名必须分别为 `vitek.xlsx`、`lis.xlsx`，仅处理各自第一个工作表。无需上传原来的 `summary.xlsx`；程序自动生成新汇总表，不覆盖原始三个文件。请选择未在 Excel 中打开的结果文件，以免文件占用或外部编辑造成冲突。

## 汇总规则

- VITEK 的 B–F 列写入汇总 A–E 列，按这五列完全相同的记录合并两种报告；实验室报告写 F/G 列，微生物图表报告写 H/I 列。相同键的同类报告若重复则报错，不静默覆盖。
- LIS 的 H 列在内存中转换：耐药 → R、敏感 → S、中间/中介 → I、非敏感 → NS、空值 → 字符串 NULL、SDD 保持不变。其他解释保留原值，原始 LIS 文件不修改。
- **匹配口径说明**：LIS 的 I 列实际是完成时间，不能与汇总 A 列菌株编号匹配。本版按 LIS C 列菌株编号＋F 列抗生素匹配汇总 A 列＋E 列；LIS I 列仅用于日期核对。药物名、菌株编号只去除首尾空白，不猜测别名或修正编号。
- 同键 LIS 多条记录全部展开保留，不自动选最新记录、不互相覆盖。没有对应 LIS 的 VITEK 记录仍保留。
- LIS C 列菌株编号如果未出现在 VITEK 汇总的 A 列，将该行追加到汇总末尾：C → A、E → C、F → E、G → J、转换后的 H → K，其余列留空，整行 A–K 为淡红色背景。以追加前的菌株编号集合为准，同一新菌株的每条 LIS 记录均按源文件顺序保留，不因已追加同编号而遗漏后续行。
- 菌株编号已存在但药物未匹配的 LIS 记录仍只列在“匹配明细”，不触发新菌株追加规则。追加行不参与原有三方差异或日期标色，避免覆盖整行淡红底色；LIS 原始日期保留在“匹配明细”。
- 比较 F/H/J，存在差异时三个单元格均为黄色背景、红字；比较 G/I/K，存在差异时三个单元格均为黄色背景、黑字。缺失报告也参与比较，空值与 NULL 视为相同。MIC 比较忽略空白，并将 ≤/≥ 等同于 <=/>=，但原始显示值不变。
- LIS I 列与汇总 D 列均严格截取原文前 10 个字符比较，不自动修正日期格式；例如未补零日期可能被标记。不同则汇总 D 列用浅棕色标记，“匹配明细”中的两项原始时间也用浅棕色标记。
- “匹配明细”包含汇总行号、LIS 源行号、两侧日期、原始/转换后解释及未匹配情况，便于追溯。

## 测试

```bash
.venv/bin/python -m unittest test_excel_merge -v
```

## 打包可执行程序

使用 PyInstaller，打包配置为 `ExcelConversion.spec`。不包含 `execl/` 下的原始数据或任何结果表格。使用者无需安装 Python，通过窗口选择外部文件即可；打包版默认从用户文档目录选择和保存文件。

必须在目标操作系统上分别构建，不能在 macOS 上直接生成 Windows `.exe`。建议使用 Python 3.13；构建机的系统版本及 CPU 架构会影响运行兼容性，需在实际目标电脑验证。

### macOS

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements-build.txt
.venv/bin/python -m unittest test_excel_merge -v
.venv/bin/python -m PyInstaller --clean --noconfirm ExcelConversion.spec
open dist/ExcelConversion.app
```

产物为 `dist/ExcelConversion.app`，将整个 `.app` 分发给同架构的 Mac。Apple Silicon 和 Intel 建议在对应架构下分别构建。本地构建未配置 Developer ID 签名或公证，其他电脑可能提示来源未验证；公开分发前需完成 Apple 代码签名、公证并测试。

### Windows（PowerShell）

Windows 版为单文件 `.exe`，包含 Python 运行时、PySide6、openpyxl 及运行依赖。使用者无需安装 Python、依赖包或 Microsoft Excel，即可生成结果文件；查看结果则使用自己的表格软件。

在 Windows 构建电脑上安装 Python 3.13（包含 `py` 启动器），将项目源码放到该电脑后，双击 `build_windows.bat` 即可。脚本自动创建独立的 `.venv-windows`、安装依赖、运行测试并打包；构建失败时会停止并显示错误。首次构建需要联网下载依赖。

也可以在项目目录运行以下命令：

```powershell
py -3.13 -m venv .venv-windows
.\.venv-windows\Scripts\python.exe -m pip install -r requirements-build.txt
.\.venv-windows\Scripts\python.exe -m unittest test_excel_merge -v
.\.venv-windows\Scripts\python.exe -m PyInstaller --clean --noconfirm ExcelConversion.spec
.\dist\ExcelConversion.exe
```

最终只需分发 `dist\ExcelConversion.exe`，无需附带 `_internal` 文件夹、Python 安装目录或项目源码。输入的 Excel 文件仍由用户在窗口中选择，不打包进程序；启动时会自动将运行库解压到系统临时目录，因此启动速度可能慢于文件夹版，并且需要可写的临时目录。

Windows 版未配置代码签名，可能出现 SmartScreen 提示。当前 macOS 环境不能本地生成或运行 Windows `.exe`，但可以通过下面的 GitHub Actions 使用云端 Windows 构建。最终仍需在未安装 Python 的 Windows 电脑上验证两步汇总流程。

### GitHub Actions（Mac 提交，云端生成 EXE）

已配置 `.github/workflows/build.yml`，每次推送都会在 `windows-latest` 上安装 x64 Python 3.13、安装构建依赖、运行测试，再使用同一个 `ExcelConversion.spec` 生成包含 Python 运行环境的单文件 EXE。也支持手动触发。

1. 将项目源码推送到 GitHub 仓库，确保包含工作流、打包配置、依赖文件、Python 源码及测试文件。不要上传 `.venv`、原始 Excel 或汇总结果。
2. 打开仓库的 **Actions → Build Windows EXE**，等待本次构建成功。首次使用时需允许仓库运行 GitHub Actions；手动运行可在工作流进入默认分支后点击 **Run workflow**。
3. 打开对应运行记录，在 **Artifacts** 中下载 **ExcelConversion-Windows-x64** 并解压，获得可分发的单个 EXE。下载产物通常需要登录 GitHub，产物保留 14 天。
4. 将 EXE 放到 Windows 电脑运行，无需安装 Python；输入表格通过窗口选择。

这里不是在 macOS 本地交叉编译，实际执行打包的是 GitHub 的 Windows Runner。配置已在本地做结构检查；云端构建是否成功以 Actions 运行结果为准。流水线不会自动提交代码、创建 Release 或上传 Excel 数据。

`.gitignore` 已忽略 `execl/*.xlsx` 和 `summary_result*.xlsx`，但忽略规则不会移除已被 Git 跟踪的文件。推送前请检查暂存区，确保未包含实际数据；测试使用临时生成的表格，不需要提交源数据。
