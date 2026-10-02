# Excel Tool

Windows Microsoft Excel 原生 LaTeX 数学公式编辑工具与 Codex Skill。可新增和批量更新原生可编辑公式，按范围排版，并验证保存文件中的 OfficeMath；不是单元格计算公式编辑器。

要求 Windows PowerShell 5.1、交互式桌面和 Microsoft 365 Excel Build 20131 或更高。已在 Excel 2609 / 16.0.20430.20032 验证。无需 Python、LaTeX 安装或 Excel add-in。

## 使用

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
```

测试只使用 artifacts 下的独立工作簿，覆盖创建、多类公式、同名更新、保存后重新打开、原生结构和前置拒绝条件，不修改用户策划文档。
