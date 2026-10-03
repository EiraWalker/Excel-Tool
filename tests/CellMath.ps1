#requires -Version 5.1
[CmdletBinding()]
param([string]$ArtifactsDirectory)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$script=Join-Path $root 'excel-tool\scripts\Excel-CellMath.ps1'
if(-not $ArtifactsDirectory){$ArtifactsDirectory=Join-Path $root 'artifacts'}
$run=Join-Path $ArtifactsDirectory ('cell-math-test-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($run)
$path=Join-Path $run 'cell-math.xlsx'
$planPath=Join-Path $root 'examples\cell-math.json'
$checks=New-Object Collections.Generic.List[string]
function Assert($condition,[string]$message){if(-not $condition){throw $message};$checks.Add($message)}
function Fail([scriptblock]$action,[string]$pattern){
 $caught=$null;try{& $action|Out-Null}catch{$caught=$_}
 Assert ($null -ne $caught -and $caught.Exception.Message -match $pattern) ('Expected refusal: '+$pattern)
}
$book=$null;$excel=$null
try{
 $compiled=(& $script -Operation Compile -Workbook $path -PlanPath $planPath|Out-String|ConvertFrom-Json)
 Assert ($compiled.cells[0].text.Contains('Tcrit') -and $compiled.cells[1].text.Contains('αO(1−O)r')) 'No spaces inserted between base and scripts'
 $badPath=Join-Path $run 'bad.json'
 @{cells=@(@{sheet='CellMath';cell='A1';segments=@(@{latex='\frac{1}{2}'})})}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $badPath -Encoding UTF8
 Fail {& $script -Operation Apply -Workbook (Join-Path $run 'must-not-exist.xlsx') -PlanPath $badPath -Create} 'Unsupported cell-math command'
 Assert (-not(Test-Path -LiteralPath (Join-Path $run 'must-not-exist.xlsx'))) 'Unsupported formula rejected before Excel creation'
 Fail {[CellMathParser]::Parse('H^{p')} 'Unclosed group'
 Fail {[CellMathParser]::Parse('T_{a^b}')} 'Nested scripts'
 Fail {[CellMathParser]::Parse('T_a^b')} 'Combined scripts'
 $r=(& $script -Operation Apply -Workbook $path -PlanPath $planPath -Create -ArtifactsDirectory $run|Out-String|ConvertFrom-Json)
 Assert ($r.status -eq 'saved_and_verified' -and $r.cell_count -eq 3) 'Three native text cells saved and XML verified'
 $book=[Runtime.InteropServices.Marshal]::BindToMoniker($path);$excel=$book.Application;$sheet=$book.Worksheets.Item('CellMath')
 Assert ($sheet.Shapes.Count -eq 0) 'No floating shapes created'
 $sheet.Columns.Item('B').ColumnWidth=105
 $sheet.Range('A1').Value2='保留内容';$sheet.Range('C1').Formula='=2+3'
 $shape=$sheet.Shapes.AddTextbox(1,600,10,100,30);$shape.TextFrame2.TextRange.Text='preserve shape'
 $book.Save()
 $update=@{cells=@(@{sheet='CellMath';cell='B4';replace_text=$true;row_height=46;segments=@(@{text='缺氧惩罚 '},@{latex='\alpha_O(1-O)^r'})})}
 $updatePath=Join-Path $run 'update.json';$update|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $updatePath -Encoding UTF8
 $r=(& $script -Operation Apply -Workbook $path -PlanPath $updatePath -Open -ArtifactsDirectory $run|Out-String|ConvertFrom-Json)
 Assert (Test-Path -LiteralPath $r.backup) 'Backup saved before updating cells'
 Assert ($sheet.Range('A1').Value2 -eq '保留内容' -and $sheet.Range('C1').Formula -eq '=2+3' -and $sheet.Shapes.Count -eq 1 -and $shape.TextFrame2.TextRange.Text -eq 'preserve shape') 'Unrelated text, calculation, and drawing preserved'
 $text=[string]$sheet.Range('B4').Value2
 Assert ($sheet.Range('B4').Characters($text.IndexOf('O')+1,1).Font.Subscript) 'O is a real character-level subscript'
 Assert ($sheet.Range('B4').Characters($text.LastIndexOf('r')+1,1).Font.Superscript) 'r is a real character-level superscript'
 Assert ($sheet.Range('B4').Characters($text.IndexOf('α')+1,1).Font.Name -eq 'Cambria' -and -not $sheet.Range('B4').Characters($text.IndexOf('α')+1,1).Font.Italic) 'Math glyphs use upright Cambria'
 $literalPath=Join-Path $run 'literal.json'
 @{cells=@(@{sheet='CellMath';cell='B8';segments=@(@{latex='123'})},@{sheet='CellMath';cell='B10';segments=@(@{latex='=T_{crit}'})})}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $literalPath -Encoding UTF8
 [void](& $script -Operation Apply -Workbook $path -PlanPath $literalPath -Open -ArtifactsDirectory $run)
 Assert ($sheet.Range('B8').Value2 -is [string] -and -not $sheet.Range('B10').HasFormula -and $sheet.Range('B10').Value2 -eq '=Tcrit') 'Numeric and equals-prefixed display remain literal text'
 $update.cells[0].replace_text=$false;$update|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $badPath -Encoding UTF8
 Fail {& $script -Operation Apply -Workbook $path -PlanPath $badPath -Open} 'Nonempty cell'
 $update.cells[0].replace_text=$true;$update.cells[0].cell='C1';$update|ConvertTo-Json -Depth 8|Set-Content -LiteralPath $badPath -Encoding UTF8
 Fail {& $script -Operation Apply -Workbook $path -PlanPath $badPath -Open} 'calculation formula'
 $sheet.Range('A1').Value2='unsaved sentinel'
 Fail {& $script -Operation Apply -Workbook $path -PlanPath $updatePath -Open} 'no unsaved changes'
 Assert ($sheet.Range('A1').Value2 -eq 'unsaved sentinel') 'Unsaved edits preserved'
 $sheet.Range('A1').Value2='保留内容';$book.Save()
 $sheet.Protect('test-password')
 $book.Save()
 Fail {& $script -Operation Apply -Workbook $path -PlanPath $updatePath -Open} 'protected'
 $sheet.Unprotect('test-password');$book.Save()
 $book.Close($true);$book=$null
 $book=$excel.Workbooks.Open($path)
 $v=(& $script -Operation Verify -Workbook $path -PlanPath $updatePath|Out-String|ConvertFrom-Json)
 Assert ($v.verified -and $v.cells[0].text -eq '缺氧惩罚 αO(1−O)r') 'Exact text, fonts and scripts survive reopening'
 $report=[pscustomobject]@{status='passed';checks=$checks.ToArray();workbook=$path;artifacts=$run}
 $report|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $run 'result.json') -Encoding UTF8
 $report|ConvertTo-Json -Depth 6
}finally{
 if($null -ne $book){$book.Close($true)}
 if($null -ne $excel -and $excel.Workbooks.Count -eq 0){$excel.Quit()}
}
