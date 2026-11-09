# TXT → Excel 样式重建工具（公开演示版）

把仪器导出的制表符 TXT 数据，按既有「范例工作簿」的样式重新生成 Excel。范例工作簿是样式源：
字体、字号、颜色、填充、边框、对齐、数字格式、列宽、行高基准、合并、批注、冻结窗格、页面设置均从范例表复制。

- Python 3.9（已在 3.9.11 上验证）
- 必需依赖：`openpyxl`；桌面界面使用 `tkinter` + `Pillow`
- 默认后端只依赖 openpyxl，不需要安装 Excel

> **这是公开演示副本。** 仓库内的 `samples/` 与 `assets/templates/` 全部是**完全虚构**的演示数据，
> 批号统一使用 `DEMO-` 前缀，不对应任何真实客户、项目或批次。仓库**不包含** exe、真实业务数据或实测截图。

## 许可（License）

本项目采用 **PFAC License 1.0**（Personal Freedom & Anti-Commercial Copyleft License），完整原文见 [LICENSE](LICENSE)。

- **个人非商业使用**：自然人可用于个人学习、研究、实验和娱乐等非商业目的。
- **企业与商业使用受限**：不得用于获取商业收益、履行企业工作职责、代表商业主体使用，或为商业主体提供产品/服务/技术支持/竞争优势；个人身份不构成商业使用的豁免。
- **修改与衍生须同许可**：实质性修改后必须继续采用 PFAC 1.0，并在 30 日内公开修改后的完整对应源代码（即使不分发也适用）；传播时须保留版权与许可声明、附带许可全文并标注修改。
- **不得改许可规避条件**：不得删除或弱化非商业限制，不得改为允许商业使用或闭源的许可，不得通过改名、拆分模块等方式规避。
- **第三方组件不被重新授权**：openpyxl、Pillow、Python、Tcl/Tk 等第三方组件仍受其各自许可证约束，不因本项目采用 PFAC 而改为 PFAC；详见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。
- **无维护与支持义务**：作者不承诺维护、更新、修复、答疑或兼容性支持，作品按“现状”提供。
- **完整条件以 LICENSE 为准**：以上仅为摘要，不构成许可文本；如摘要与 LICENSE 有差异，以 LICENSE 原文为准。

> 注意：`LICENSE` 按作者提供的许可证原文收录，仅将版权行中的 `[YEAR]`、`[AUTHOR]` 替换为 `2026`、`KazuiWANG`，其余正文（含原有章节编号）一字未改。本许可证不是 OSI 认可的开源许可证，请勿将其描述为“开源”或“可自由商用”。

## 怎么启动

### 源码启动

```powershell
cd <项目根目录>

# 首次使用安装依赖（国内镜像）
python -m pip install -i https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt

# 启动桌面程序
python app.py
```

若机器上有多个 Python，请把示例里的 `python` 换成 Python 3.9 解释器的实际路径。

### 命令行批量转换

```powershell
# 默认把 samples/ 下的演示 TXT 转到 output/
python tools\convert_cli.py

# 指定文件与输出目录
python tools\convert_cli.py "samples\DEMO-BATCH-0001.txt" -o "output"

# 强制使用某个模板（模板 id 为 demo_batch / demo_cycle）
python tools\convert_cli.py "samples\DEMO-BATCH-0001.txt" --template demo_batch
```

- 不传文件时只扫描 `samples/*.txt`；传文件时按路径处理。
- 返回码：全部成功为 `0`，有失败为 `1`。

## 演示数据与预期输出

| 输入 TXT | 输出工作簿 | sheet | 数据行 |
|---|---|---|---|
| `samples/DEMO-BATCH-0001.txt` | `output/演示分容数据.xlsx` | `演示分容` | 7 |
| `samples/DEMO-CYCLE-0002.txt` | `output/演示循环数据.xlsx` | `演示循环` | 5 |

演示数据特意包含：

- 组内留空的分组字段（批号、日期、班组、设备等），用于验证整组补齐；
- TXT 缺值的公式列（如 `容量mAh`、`内阻mΩ`），用于验证“公式只算数值、不写公式”；
- `#VALUE!`、`#DIV/0!` 异常标记，按原值保留并写入日志；
- 百分比写法（`97.50%`），按 Excel 语义写成 `0.975`。

重新生成演示数据（脚本可重复执行，覆盖演示文件）：

```powershell
python tools\make_demo_data.py
```

脚本会重建 2 个 TXT、2 个范例工作簿（含表头合并、填充、边框、数字格式、列宽、批注、冻结、页面设置）与 2 个模板 JSON。

## 输出与实现要点

- 数据区只写数值：模板公式仅用于计算补位值，输出工作簿中不允许出现 `data_type='f'` 或以 `=` 开头的文本。
- TXT 异常标记（`#VALUE!`、`#DIV/0!`、`/` 等）不丢弃，按原值保留。
- 编码按 `utf-8-sig → utf-8 → gb18030 → gbk` 自动尝试。
- 批注按“同批号 + 同出现次序”回填，TXT 中不存在的批号不会回填。
- 输出只保留与本次 TXT 对应的 sheet，不复制范例中的其他 sheet。

## 测试

```powershell
python -m unittest discover -s tests -v
```

测试只使用 `tests/fixtures/` 与 `samples/` 中的合成数据，不依赖任何真实业务文件。当前共 18 项（解析、数值转换、公式求值、模板匹配、端到端转换、自动学习）。

## 目录结构

```text
txt重建excel/
├─ app.py                      # GUI 入口
├─ build_exe.ps1               # PyInstaller 打包脚本（可 -Python 指定解释器）
├─ txt2xlsx.spec               # PyInstaller 配置（目录版 onedir）
├─ requirements.txt            # 运行依赖 + 打包依赖
├─ LICENSE                     # PFAC 1.0 原文
├─ THIRD_PARTY_NOTICES.md      # 第三方组件与许可
├─ assets/
│  ├─ templates/               # 演示范例工作簿（样式源，虚构数据）
│  └─ ui/                      # background.png、icon.png、icon.ico（作者确认可用）
├─ samples/                    # 虚构演示 TXT
├─ templates/                  # 演示模板 JSON
├─ src/txt2xlsx/               # 解析、转换、模板学习、GUI
├─ tools/
│  ├─ convert_cli.py           # 命令行转换
│  ├─ verify.py                # 格式还原验证
│  └─ make_demo_data.py        # 演示数据生成器
├─ tests/                      # unittest（合成数据）
└─ docs/验证报告.md            # 合成数据的验证结论
```

## 打包 exe（可选）

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
# 或指定解释器
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1 -Python "C:\Path\To\python.exe"
# 产物：dist\TXT重建Excel\TXT重建Excel.exe
```

采用 onedir（目录版）而不是单文件：资源路径稳定、启动快，`templates/` 与 `output/` 可落在 exe 同级可写目录。
分发时需连同 `_internal\` 整个文件夹一起拷贝，并补齐 Python、Tcl/Tk、openpyxl、et_xmlfile、Pillow 的许可文本。
本仓库**不提交** exe 与构建产物。

## 素材说明

`assets/ui/background.png`、`icon.png`、`icon.ico` 由项目作者确认拥有使用权；仓库未附外部授权凭证，第三方再使用需自行确认。
