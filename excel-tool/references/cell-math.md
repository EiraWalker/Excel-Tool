# 单元格数学富文本

用于状态、属性、参数和“含义与作用”等列中的数学字母。输出属于真实单元格，可以随行列移动、自动换行和编辑文字。数学显示不会参与 Excel 计算。

默认中文 MiSans 12 pt、字母 Cambria 正体 14 pt。上下标采用 Excel 字符级 `Font.Subscript` / `Font.Superscript`，保持基底与上下标紧邻。不要用空格垫出上下标，不要把 Cambria Math 当作普通单元格默认字体。脚本始终关闭斜体，选定字体未安装时会停止。

## 计划格式

JSON 中 `cells` 为非空数组，每项指定 `sheet`、单个 A1 地址 `cell` 和有序 `segments`。每段只能包含 `text` 或 `latex` 之一。LaTeX 不含外层 `$...$` / `\[...\]`；JSON 中反斜杠写成 `\\`。文本段可用 `\n` 换行。

```json
{
  "text_font": "MiSans",
  "math_font": "Cambria",
  "text_size": 12,
  "math_size": 14,
  "cells": [{
    "sheet": "状态与属性",
    "cell": "F4",
    "replace_text": true,
    "row_height": 48,
    "segments": [
      {"text": "有效点燃温度 "},
      {"latex": "T_{\\mathrm{crit}}"},
      {"text": "：用于判断是否点燃。"}
    ]
  }]
}
```

`text_font`、`math_font`、`text_size`、`math_size`、`text_color`、`math_color` 可以在计划级设置，也可逐单元格覆盖。默认颜色分别为 `#334155`、`#147568`，颜色使用 `#RRGGBB`，字号范围 8–72 pt。`row_height` 为可选逐单元格字段，范围 16–409 pt，会改变该单元格所在整行的高度。列宽保留，按表格布局另行设置足够的显示宽度。

`replace_text` 默认 false，只有用户授权替换既有文本时设为 true。无论此字段如何，已有 Excel 计算公式均拒绝覆盖。合并区域必须指定左上角地址；一次计划不能重复目标地址。内容覆盖与字体修改只作用于指定单元格，开启换行、顶端对齐；保留原数字格式和既有图形。

## 支持范围

- 字母、数字、括号和线性运算符；`-` 转为数学减号 `−`。
- `T_{crit}`、`T_{\mathrm{ign}}`、`\alpha_O`、`H^p`、`(1-O)^r`：单层上下标，支持花括号和正体文字组。
- 希腊字母：`\alpha`、`\beta`、`\gamma`、`\delta`、`\epsilon`、`\theta`、`\lambda`、`\mu`、`\pi`、`\rho`、`\sigma`、`\tau`、`\phi`、`\omega`，以及 `\Delta`、`\Theta`、`\Lambda`、`\Pi`、`\Sigma`、`\Omega`。
- 运算符：`\times`、`\cdot`、`\le` / `\leq`、`\ge` / `\geq`、`\ne` / `\neq`、`\in`、`\land`、`\lor`、`\pm`、`\infty`、`\prime`。
- `\mathrm{...}`、`\text{...}` 保持数学段字体正体；中文说明优先放在 `text` 段。`\quad`、`\,` 为显式空格。

此解析器是有限的行内表示法，不支持分式、根式、矩阵、分段函数、嵌套上下标，以及同一基底同时带上标和下标（普通文字无法把二者垂直堆叠）。这些结构会在连接或修改 Excel 前拒绝。可以在用户认可线性写法后使用 `/`，不要默默改写。只接受 BMP 字符；暂不支持 emoji 等代理对字符，以免字符索引错位。

## 调用与验证

从当前技能目录解析完整脚本路径，使用 Windows PowerShell 5.1：

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "<技能目录>\scripts\Excel-CellMath.ps1" -Operation Compile -Workbook "D:\Docs\设计.xlsx" -PlanPath "D:\Docs\cell-math.json"
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "<技能目录>\scripts\Excel-CellMath.ps1" -Operation Inspect -Workbook "D:\Docs\设计.xlsx"
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "<技能目录>\scripts\Excel-CellMath.ps1" -Operation Apply -Workbook "D:\Docs\设计.xlsx" -PlanPath "D:\Docs\cell-math.json"
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "<技能目录>\scripts\Excel-CellMath.ps1" -Operation Verify -Workbook "D:\Docs\设计.xlsx" -PlanPath "D:\Docs\cell-math.json"
```

`Compile` 不连接 Excel，输出准确字符序列及样式范围，先检查基底和上下标之间没有新增空格。`Inspect` 不修改内容。默认 Apply 只使用当前连接实例内完整路径匹配的工作簿；用户授权打开该文件时加 `-Open`。`Apply -Create` 只创建尚不存在的 `.xlsx`，父目录需已存在；不会关闭其他工作簿。`Verify` 直接读取保存文件，无需打开 Excel。

Apply 在写入前为已有文件生成完整备份，写入后先保存 staged 副本并校验，再保存目标。结果必须为 `status=saved_and_verified`。保存验证检查单元格字符串、逐字符正体、上下标以及字母/数字字体，不通过会停止；Excel 自动替换部分标点、运算符字体时记录 `fallback_symbols`，不将其误报为所有符号均保留指定字体。

备份和 staged 文件默认在工作簿旁的 `.excel-latex-artifacts`；可用 `-ArtifactsDirectory` 指定目录。失败时查看 `failure.json` 的阶段和备份；已有工作簿可能留下部分未保存修改，不宣称原子回滚，不自动覆盖恢复。结构验证通过后再检查显示宽度、换行与裁切；用户要求看效果或弹窗时，仅对 Agent 已确认正确的结果请求反馈。
