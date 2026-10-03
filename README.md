# Excel Tool

Windows Microsoft Excel 数学排版工具与 Codex Skill。支持真实单元格中的行内字母、上下标和中文混排，也支持浮动形状中的原生可编辑 OfficeMath。两种模式均验证保存后的结构，用于数学显示。

要求 Windows PowerShell 5.1、交互式桌面、本机 Excel 和所选字体。原生 OfficeMath 模式另需 Microsoft 365 Excel Build 20131 或更高，已在 Excel 2609 / 16.0.20430.20032 验证。无需 Python、LaTeX 安装或 Excel add-in。

## 单元格字母与上下标

默认中文 **MiSans**、数学字母 **Cambria 正体**，用真实字符上下标保持紧凑间距，不创建文本框。支持 `T_{crit}`、`\alpha_O`、`(1-O)^r` 等行内 LaTeX 和中文说明混排。

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File '.\excel-tool\scripts\Excel-CellMath.ps1' -Operation Compile -Workbook 'D:\Docs\单元格测试.xlsx' -PlanPath '.\examples\cell-math.json'
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File '.\excel-tool\scripts\Excel-CellMath.ps1' -Operation Apply -Workbook 'D:\Docs\单元格测试.xlsx' -PlanPath '.\examples\cell-math.json' -Create
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File '.\excel-tool\scripts\Excel-CellMath.ps1' -Operation Verify -Workbook 'D:\Docs\单元格测试.xlsx' -PlanPath '.\examples\cell-math.json'
```

详细字段与边界见 [单元格富文本](excel-tool/references/cell-math.md)。既有文本需显式 `replace_text: true`，不覆盖计算公式；分式、分段函数等二维结构会在修改前拒绝，使用下方原生模式处理。

## 原生 OfficeMath Latex粘贴方法

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File '.\excel-tool\scripts\Excel-Latex.ps1' -Operation Inspect -Workbook 'D:\Docs\设计.xlsx'
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File '.\excel-tool\scripts\Excel-Latex.ps1' -Operation Apply -Workbook 'D:\Docs\设计.xlsx' -PlanPath '.\examples\equations.json' -Open -Capture
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File '.\excel-tool\scripts\Excel-Latex.ps1' -Operation Verify -Workbook 'D:\Docs\设计.xlsx' -PlanPath '.\examples\equations.json'
```

输入字段见 [计划格式](excel-tool/references/plan.md)。新增文件使用 `Apply -Create`；已有未保存修改、保护或 AutoSave 开启时会停止。输出包括备份、staged 校验副本、原生结构信息和可选截图。生成文件不进入本仓库。

## Codex 安装

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File '.\Install-ExcelTool.ps1'
```

安装脚本不覆盖其他技能。有 Skillshare 时安装到其 source，并把 Codex 联接到安装版本；无 Skillshare 时直接联接仓库内技能。更新仓库后重新执行安装脚本，已安装版本的本地修改会触发停止，避免覆盖。新开 Codex 对话后可自动发现 **Excel Tool**，也可通过 `$excel-tool` 显式调用。自动发现取决于请求与 skill 描述匹配，不是保证每条 Excel 请求都会选用。

## 验证

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File '.\tests\Integration.ps1'
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File '.\tests\CellMath.ps1'
```

测试只使用 artifacts 下的独立工作簿。原生测试覆盖创建、多类公式和更新；单元格测试覆盖真实上下标、正体字体、无新增浮动对象、保存后重新打开、其他文本/计算/图形保留和拒绝条件，不修改用户策划文档。

## License

本项目采用 [MIT License](LICENSE)。
