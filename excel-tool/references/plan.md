# JSON 公式计划

计划是 UTF-8 JSON，根对象有 `equations` 数组。不支持对 JSON 中的命令求值。

| 字段 | 要求与作用 |
| --- | --- |
| `name` | 必填。字母开头的 ASCII 字母、数字、下划线，最多 80 字符；整个计划中唯一，用于更新 |
| `sheet` | 必填。工作表名称；编辑已有文件时须存在；`-Create` 会创建计划需要的工作表 |
| `anchor` | 必填。单个 A1 单元格或矩形范围，允许 `$`。单个合并单元格自动使用 MergeArea |
| `latex` | 必填。不带外层界定符的数学 LaTeX，最多 20000 字符。JSON 中反斜杠写成 `\\` |
| `font_size` | 可选，默认 17，范围 8..72，单位 pt，使用 Cambria Math |
| `row_height` | 可选，范围 16..409 pt，对 anchor 涉及的每行设置；不自动扩大周围行 |
| `width` / `height` | 可选，至少 16 pt。默认锚定范围宽高减去边距；较大公式需明确预留空间 |
| `clear_text` | 可选，默认 false。true 清空 anchor 范围内容，保留格式；只在授权替换文本时使用 |

脚本自动在表达式外增加 `\[...\]`，然后从 Unicode 剪贴板实际粘贴到新建的原生公式编辑区。

```json
{
  "equations": [
    {
      "name": "HeatExchange",
      "sheet": "Sheet1",
      "anchor": "A3:F3",
      "row_height": 60,
      "latex": "D=\\frac{T_A-T_B}{2}\\times k_A\\times k_B"
    },
    {
      "name": "Destruction",
      "sheet": "Sheet1",
      "anchor": "A5:F5",
      "row_height": 72,
      "latex": "D(F,R)=\\begin{cases}1,&F\\ge R\\\\0,&F<R\\end{cases}"
    }
  ]
}
```

用 `Apply -Create` 创建新的 `.xlsx`；文件父目录需已存在，文件本身必须不存在。创建新文件不覆盖已有文件。对已有文件使用 `-Open` 明确授权脚本打开指定文件，已打开文件无需此参数。

输出目录默认为工作簿目录下 `.excel-latex-artifacts/<时间>-<随机ID>/`，可用 `-ArtifactsDirectory` 指定其他目录。现有文件产生 before 备份、staged 原生检查副本、result.json；失败时产生 failure.json。备份和截图可能含文档内容，不提交到工具仓库。

