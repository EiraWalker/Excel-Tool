# 已验证的原生转换方法

本能力由 2026-10-03 在 Excel 2609 / Build 16.0.20430.20032 上完成的策划表编辑提取。成功样例覆盖分式、下标、撇号、分段判断、中文条件、幂和两行有效点燃温度。

操作链是 `EquationInsertNew` → Unicode 剪贴板 `\[数学内容\]` → `Paste` → `EquationProfessional`，全部使用真实 Excel 命令。插入前核对实际窗口焦点；粘贴前再次检查目标工作簿和原生公式编辑上下文。`Paste` 是指向 Excel 的 COM 命令，不是向当前前台窗口发送 Ctrl+V。单个公式仅对焦点/编辑上下文丢失等错误重试一次，重试前删除本次临时公式；其他错误直接报告。

前台焦点需要把自动化线程分别连接到前台窗口线程和 Excel 窗口线程，再调用 ShowWindow、BringWindowToTop、SetForegroundWindow、SetFocus，结束后解除连接。只使用 AppActivate 的成功返回值，或只连接前台线程，在本机不足以确保焦点。不得向前台 Codex 或其他窗口发送全局 Ctrl+V。

保存后的公式在 `xl/drawings/*.xml` 的 `xdr:sp` / 文本区内包含原生 OfficeMath。Office 也保存兼容 fallback 形状；验证应只统计带 `m:oMath` 的原生分支，避免把一个公式计成两个。对齐公式中的 `&` 可能成为 OfficeMath 对齐标记，不应只因原始 XML 含 `&` 就判断渲染失败，应检查实际显示。

## 边界

- 本工具不实现完整 TeX，不安装 TeX 或额外 Excel add-in。支持范围由当前 Excel 的 LaTeX 转换器决定。
- Excel 原生方程在形状中，不等同于 `=SUM(...)` 计算公式。普通单元格不是导入方程的目标。
- 不提供 Excel 的 Word-only `OMaths.Add`，不依赖 `TextRange.Text` 赋值解析数学内容。
- staged XML 检查验证结构存在，不验证用户公式的数学含义，也不证明任何分辨率下不裁切。Agent 要检查结果截图。
- 出错后 stage 清理只还原本工具临时公式和行高，不保存回滚状态。commit 期间有旧形状删除和原文本清理，因此保留 before 备份及阶段错误报告，不能承诺事务级原子回滚。
- 脚本只附加到 COM 当前返回的 Excel 实例，不扫描所有实例；不关闭原工作簿或全局重启 Excel。
- 尝试恢复剪贴板，输出 `clipboard_restored`。若恢复失败，准确报告；不要读取或输出用户剪贴板内容。

## 官方依据

- Excel LaTeX 转换与最低版本：https://learn.microsoft.com/en-us/office/math/latex
- 剪贴板界定符与支持：https://learn.microsoft.com/en-us/office/math/latex#clipboard-support
- ExecuteMso：https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.executemso
- GetEnabledMso：https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.getenabledmso
- SaveCopyAs：https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas
- 官方 Excel 命令标识：https://github.com/OfficeDev/office-fluent-ui-command-identifiers/blob/main/Microsoft%20365/Current%20Channel/excelcontrols.xlsx
- AttachThreadInput：https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-attachthreadinput
- SetForegroundWindow：https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setforegroundwindow
