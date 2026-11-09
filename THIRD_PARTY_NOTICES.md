# 第三方组件与许可声明（THIRD PARTY NOTICES）

本文件列出本项目实际使用、且经本地安装包元数据核实过的第三方组件。

- 核实时点：2026-11-05
- 核实方式：读取本机 Python 3.9 `site-packages` 中各包的 `*.dist-info/METADATA`、`LICENSE`/`COPYING` 文件，以及 Python 与 Tcl/Tk 安装目录内的许可文件；未凭组件名推测。
- 本文件是事实清单，**不是法律意见**；标注「待人工确认」的项目需要版权持有人或法务自行判断。
- 本项目自身采用 PFAC License 1.0（见 [LICENSE](LICENSE)）。**第三方组件不因被打包或分发而改为 PFAC 授权**，它们仍受下表各自的许可证约束。

## 一、随程序分发（运行时）

| 组件 | 实际版本 | 许可证（本地元数据） | 官方来源 | 分发注意 |
|---|---|---|---|---|
| Python | 3.9.11 | PSF License（Python 3.9 安装目录 `LICENSE.txt`，历史条款 A–G） | https://www.python.org | exe 由 PyInstaller 内嵌解释器；分发时建议随附 PSF 许可文本 |
| Tcl/Tk | 8.6（随 Python 3.9 安装） | Tcl/Tk License（BSD 风格；`tcl/tk8.6/license.terms`） | https://www.tcl.tk | 打包目录 `_internal/_tk_data/license.terms` 已存在；要求保留版权声明并原样附带本声明 |
| openpyxl | 3.1.5 | MIT（`openpyxl-3.1.5.dist-info/LICENCE.rst`） | https://openpyxl.readthedocs.io | 保留版权与许可声明 |
| et_xmlfile | 2.0.0 | MIT（`et_xmlfile-2.0.0.dist-info`） | https://foss.heptapod.net/openpyxl/et_xmlfile | openpyxl 的依赖；保留版权与许可声明 |
| Pillow | 11.3.0 | MIT-CMU（`pillow-11.3.0.dist-info/licenses/LICENSE`） | https://python-pillow.org | 该 LICENSE 文件同时包含随 wheel 分发的原生库（如 FreeType 等）的许可声明；分发时应保留完整 LICENSE 文本 |

## 二、仅构建期使用（不随 exe 运行）

| 组件 | 实际版本 | 许可证（本地元数据） | 说明 |
|---|---|---|---|
| PyInstaller | 6.22.3 | GPL-2.0-or-later **WITH Bootloader-exception**（`COPYING.txt`） | 该例外明确允许链接/嵌入编译后的 bootloader，并可用自选条款分发由其构建的程序；因此不要求本项目整体改为 GPL |
| pyinstaller-hooks-contrib | 2026.9 | Apache-2.0 与 GPL-2.0 混合（`licenses/LICENSE` 分节列明） | 仅构建期加载的 hook，不随最终程序运行 |
| altgraph | 0.17.5 | MIT | PyInstaller 依赖 |
| pefile | 2024.8.26 | MIT | PyInstaller 依赖 |
| pywin32-ctypes | 0.2.3 | BSD-3-Clause | PyInstaller 依赖 |
| packaging | 26.3 | Apache-2.0 OR BSD-2-Clause（`License-Expression`） | 仅构建期；未发现独立 NOTICE 文件 |
| importlib-metadata | 8.7.1 | Apache-2.0 | 仅构建期 |
| zipp | 3.23.1 | MIT | importlib-metadata 依赖 |
| setuptools | 58.1.0 | MIT（METADATA 的 `License` 字段为 UNKNOWN，但含 MIT 分类器与 LICENSE 文件） | 仅构建期 |

> 说明：PyInstaller 会把 Python 解释器、Tcl/Tk 与运行期纯 Python 包打进 `_internal/` 或归档，因此“构建期组件”和“随程序分发组件”需要分开看。PyInstaller 自身的 GPL 例外已核实允许分发非自由/商业程序，但这不改变本项目采用 PFAC 的选择，也不构成对第三方的再授权。

## 三、素材与数据（不是软件依赖，单独列出）

| 项目 | 位置 | 状态 |
|---|---|---|
| 界面背景图（含动漫风格人物立绘） | `assets/ui/background.png` | 项目作者确认拥有使用权；未附外部授权凭证 |
| 窗口/任务栏图标 | `assets/ui/icon.png`、`assets/ui/icon.ico` | 项目作者确认拥有使用权；未附外部授权凭证 |
| 界面字体 | 无内置字体文件 | 运行时调用系统字体（`Microsoft YaHei UI` / `Microsoft YaHei` / `SimHei`），未随程序分发字体文件，因此不涉及字体再分发 |
| 范例工作簿与样例 TXT | `assets/templates/*.xlsx`、`samples/*.txt` | 由 `tools/make_demo_data.py` 生成的**完全虚构演示数据**，不属于第三方软件或第三方素材；随项目一并提供 |
| 转换输出 | `output/*` | 由演示数据运行时生成，已被 `.gitignore` 排除，不进入仓库 |

## 四、需要随分发保留的许可文本

公开仓库或分发 exe 时，建议至少保留：

1. 本文件（`THIRD_PARTY_NOTICES.md`）；
2. 项目 `LICENSE`（PFAC 1.0 原文）；
3. Python 的 `LICENSE.txt`（PSF）与 Tcl/Tk 的 `license.terms`；
4. openpyxl 的 `LICENCE.rst`、et_xmlfile 的 `LICENCE.rst`、Pillow 的 `LICENSE`。

本仓库不包含 exe 构建产物。若自行打包并对外分发，需在分发物中补齐上述第三方许可文本；当前源码仓库已通过本文件披露组件清单。

## 五、无法确认 / 未核实事项

- `assets/ui/*` 三类图片：由项目作者确认拥有使用权，但未附外部授权凭证；第三方若再使用需自行确认。
- Pillow wheel 内各原生库的完整许可清单：仅核对了 `LICENSE` 文件中存在相关章节，未逐库核验版本与条款。
- Apache-2.0 组件是否需要 NOTICE：以各包实际文件为准，本次未发现 `packaging` 附带 NOTICE；若后续升级依赖，需重新核对。
- 本项目不包含从网络复制的第三方代码片段；此项基于对 `src/`、`tools/`、`tests/` 的阅读与依赖声明，无法做全量代码相似度比对。
