---
name: excel-tool
description: 在 Windows Microsoft Excel 中用 LaTeX 创建、批量排版和更新原生可编辑数学公式，支持分式、上下标、分段函数和多行公式，并验证保存后的 OfficeMath。用于 Excel 策划表、数学说明表及已有公式编辑；普通单元格计算公式不触发此能力。
---

# Excel Tool

使用本技能随附的 `scripts/Excel-Latex.ps1` 将数学 LaTeX 粘贴到 Excel 原生公式编辑区。它控制本机 Microsoft Excel；不是 Office.js 或已连接 Excel 会话工具。用户要求在 Excel 中创建/编辑数学排版公式时可自动使用，不需要再次询问已经授权的文件修改。

## 前提与范围

- Windows 交互式桌面、Windows PowerShell 5.1、Microsoft 365 Excel Build 20131 或更新版本。已在 2609 / 16.0.20430.20032 验证。
- 文件为 `.xlsx` 或 `.xlsm`，明确指定完整路径和工作表；默认只操作已打开的同名完整路径工作簿。用户要求打开现有文件时用 `-Open`；创建新文件时用 `-Create`。不关闭其他工作簿或终止 Excel。
- 编辑前工作簿应已保存、可写、工作表未保护，AutoSave 关闭。脚本拒绝覆盖未保存的修改和无工具标记的既有形状。若 Excel 存在多个实例，找不到指定工作簿时说明边界，不猜测或重开成只读副本。
- 数学公式是浮动形状中的 OfficeMath，不是单元格计算式。需要保留 LaTeX 源码时，在用户指定位置另存 JSON 计划；不要向策划表塞入实现说明。

## 调用

从当前技能所在目录解析脚本绝对路径，用 `powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File <完整脚本路径>`。脚本使用 UTF-8 BOM，可直接 `-File`，不需要此前临时脚本的 `ReadAllText` workaround。

先 `Inspect` 确认工作簿路径、工作表、保存状态和已有形状：

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "<技能目录>\scripts\Excel-Latex.ps1" -Operation Inspect -Workbook "D:\Docs\设计.xlsx"
```

将用户公式写成 JSON 计划，`latex` 提供数学内容，不包含外层 `$...$` / `\[...\]`、文档前导或包加载指令。每个公式有稳定的 `name`、`sheet`、`anchor`、`latex`。完整字段和示例见 [references/plan.md](references/plan.md)。

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "<技能目录>\scripts\Excel-Latex.ps1" -Operation Apply -Workbook "D:\Docs\设计.xlsx" -PlanPath "D:\Docs\equations.json" -Capture
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "<技能目录>\scripts\Excel-Latex.ps1" -Operation Verify -Workbook "D:\Docs\设计.xlsx" -PlanPath "D:\Docs\equations.json"
```

对已有工具创建的公式，复用相同工作表及 `name` 进行更新。脚本用 `AlternativeText=excel-latex/v1` 区分可更新对象。不能仅因为形状名称相同就覆盖用户的其他对象。

## 工作流程与验证

1. 依据公式高度设置 `row_height` 或 `height`，依据范围宽度设置 `anchor` 或 `width`。默认保留单元格内容；只在用户授权替换该范围原文本时指定 `clear_text: true`。
2. 脚本在同一连续操作中验证窗口焦点、插入公式、执行真实剪贴板粘贴、转换 Professional。不能把 LaTeX 赋给 `TextRange.Text` 当作转换，也不要跨工具调用拆开插入和粘贴。
3. 现有文件编辑前备份。所有新公式先保存在 staged 副本并验证原生 XML，然后替换工具创建的旧公式、按计划清理文本并保存目标。输出 JSON 包含备份、验证和截图位置。
4. 检查 `status=saved_and_verified`，通过 `Verify` 对照公式名称确认 `m:oMath`；有分式时确认 `m:f`，分段函数确认矩阵结构。打开 `-Capture` 产生的截图检查裁切和布局。截图只涵盖每个工作表第一个公式附近，长表需要滚动补充检查。成功退出或形状数量本身不证明 LaTeX 正确。
5. 确认无误后报告准确成果路径。失败由 Agent 自行诊断，不要求用户确认已知失败。用户明确要求查看效果/弹窗时，仅在 Agent 的结构和显示检查通过后展示结果并请求一次反馈。

异常时检查输出目录 `failure.json` 的阶段和备份。stage 阶段会尽量移除本次临时公式并还原行高；commit 阶段可能存在部分修改，不能宣称原子回滚或自动覆盖原文件恢复。先检查状态和备份，再按当前授权修复。详细机制和限制见 [references/native-math.md](references/native-math.md)。

## 安装与自动发现

`agents/openai.yaml` 保持 `allow_implicit_invocation: true`。用仓库提供的 `Install-ExcelTool.ps1` 注册到 Codex 的 skills 目录；有 Skillshare 时安装到其 source，再将 Codex 联接到该安装版本，先检查 sync dry-run 再同步。仓库是维护源码，重新运行安装脚本升级已安装版本；若安装版本有本地修改，脚本会保留并停止。无 Skillshare 时 Codex 直接联接仓库中的技能。新对话才能可靠加载新加入的技能；本对话仍可直接调用脚本。
