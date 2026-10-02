#requires -Version 5.1
[CmdletBinding()]
param([string]$ArtifactsDirectory)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$script=Join-Path $root 'excel-tool\scripts\Excel-Latex.ps1'
if(-not $ArtifactsDirectory){$ArtifactsDirectory=Join-Path $root 'artifacts'}
$run=Join-Path $ArtifactsDirectory ('integration-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($run)
$path=Join-Path $run 'native-math.xlsx'
$planPath=Join-Path $root 'examples\equations.json'
$checks=New-Object Collections.Generic.List[string]
function Assert($condition,[string]$message){if(-not $condition){throw $message};$checks.Add($message)}
function Expect-Failure([scriptblock]$action,[string]$pattern){
 $caught=$null
 try{& $action|Out-Null}catch{$caught=$_}
 Assert ($null -ne $caught -and $caught.Exception.Message -match $pattern) ('Expected refusal: '+$pattern)
}
$book=$null
try{
 $invalidCreate=Join-Path $run 'invalid-create.json'
 @{equations=@(@{name='Bad';sheet='Math';latex='x=1'})}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $invalidCreate -Encoding UTF8
 Expect-Failure {& $script -Operation Apply -Workbook (Join-Path $run 'must-not-exist.xlsx') -PlanPath $invalidCreate -Create} 'missing anchor'
 Assert (-not (Test-Path -LiteralPath (Join-Path $run 'must-not-exist.xlsx'))) 'Invalid create plan rejected without creating a file'
 $result=(& $script -Operation Apply -Workbook $path -PlanPath $planPath -Create -ArtifactsDirectory $run -Capture|Out-String|ConvertFrom-Json)
 Assert ($result.status -eq 'saved_and_verified' -and $result.equation_count -eq 5) 'Created five saved native formulas'
 $excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
 $book=$excel.Workbooks.Item('native-math.xlsx');$sheet=$book.Worksheets.Item('Math')
 $sheet.Range('A:F').ColumnWidth=16
 foreach($shape in $sheet.Shapes){$shape.Width=$sheet.Range('A1:F1').Width-6}
 $sheet.Range('A1').Value2='保留原有内容'
 $sheet.Range('G2').Formula='=2+3'
 $sheet.Range('A3').Value2='仅替换此处原文本'
 $book.Save()
 $update=@{equations=@(@{name='HeatExchange';sheet='Math';anchor='A3:F3';row_height=60;clear_text=$true;latex='D=\frac{T_A-T_B}{4}\times k_A\times k_B'})}
 $updatePath=Join-Path $run 'update.json'
 $update|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $updatePath -Encoding UTF8
 $updated=(& $script -Operation Apply -Workbook $path -PlanPath $updatePath -ArtifactsDirectory $run -Capture|Out-String|ConvertFrom-Json)
 Assert ($updated.status -eq 'saved_and_verified' -and $sheet.Shapes.Count -eq 5) 'Same-name update creates no duplicates'
 Assert ($sheet.Range('A1').Value2 -eq '保留原有内容' -and $sheet.Range('G2').Formula -eq '=2+3') 'Unrelated text and calculation preserved'
 Assert ($null -eq $sheet.Range('A3').Value2) 'Explicit clear_text clears only requested source text'
 Assert (Test-Path -LiteralPath $updated.backup) 'Backup is present before update'
 $stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
 $zip=New-Object IO.Compression.ZipArchive($stream,[IO.Compression.ZipArchiveMode]::Read,$false)
 try{
  $entry=@($zip.Entries|Where-Object {$_.FullName -match '^xl/drawings/[^/]+\.xml$'})[0]
  $reader=New-Object IO.StreamReader($entry.Open())
  try{[xml]$xml=$reader.ReadToEnd()}finally{$reader.Dispose()}
  $ns=New-Object Xml.XmlNamespaceManager($xml.NameTable)
  $ns.AddNamespace('m','http://schemas.openxmlformats.org/officeDocument/2006/math')
  $ns.AddNamespace('xdr','http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing')
  $fraction=$xml.SelectSingleNode('//xdr:sp[xdr:nvSpPr/xdr:cNvPr/@name="HeatExchange"][.//m:oMath]//m:f/m:den//m:t',$ns)
  Assert ($null -ne $fraction -and $fraction.InnerText -eq '4') 'Updated fraction denominator persisted as native math'
 }finally{$zip.Dispose()}
 $verify=(& $script -Operation Verify -Workbook $path -PlanPath $planPath|Out-String|ConvertFrom-Json)
 Assert ($verify.native_equation_count -eq 5) 'Compatibility fallback shapes are not double counted'
 $sheet.Range('A1').Value2='unsaved sentinel'
 Expect-Failure {& $script -Operation Apply -Workbook $path -PlanPath $updatePath -ArtifactsDirectory $run} 'no unsaved changes'
 Assert ($sheet.Range('A1').Value2 -eq 'unsaved sentinel') 'Unstored user edit was not overwritten'
 $sheet.Range('A1').Value2='保留原有内容';$book.Save()
 $unmanaged=$sheet.Shapes.AddTextbox(1,600,10,100,30)
 $unmanaged.Name='ForeignFormula';$unmanaged.TextFrame2.TextRange.Text='keep this'
 $book.Save()
 $bad=@{equations=@(@{name='ForeignFormula';sheet='Math';anchor='A20:F20';latex='x=1'})}
 $badPath=Join-Path $run 'unmanaged.json'
 $bad|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $badPath -Encoding UTF8
 Expect-Failure {& $script -Operation Apply -Workbook $path -PlanPath $badPath -ArtifactsDirectory $run} 'unmanaged shape'
 Assert ($sheet.Shapes.Item('ForeignFormula').TextFrame2.TextRange.Text -eq 'keep this') 'Unmanaged shape preserved'
 $unmanaged.Delete();$book.Save()
 $bad.equations=@($update.equations[0],$update.equations[0])
 $bad|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $badPath -Encoding UTF8
 Expect-Failure {& $script -Operation Apply -Workbook $path -PlanPath $badPath -ArtifactsDirectory $run} 'Names must be unique'
 Assert ($sheet.Shapes.Count -eq 5 -and $book.Saved) 'Invalid batch rejected before mutation'
 $book.Close($false);$book=$null
 $book=$excel.Workbooks.Open($path);$sheet=$book.Worksheets.Item('Math')
 Assert ($sheet.Shapes.Count -eq 5 -and $sheet.Shapes.Item('HeatExchange').AlternativeText -eq 'excel-latex/v1') 'Native formulas and management marker survive reopen'
 $inspected=(& $script -Operation Inspect -Workbook $path|Out-String|ConvertFrom-Json)
 Assert ($inspected.workbook -eq $path -and $inspected.saved) 'Inspection attaches to the exact reopened workbook'
 $finalVerify=(& $script -Operation Verify -Workbook $path -PlanPath $planPath|Out-String|ConvertFrom-Json)
 Assert ($finalVerify.native_equation_count -eq 5) 'Saved formulas verified after reopening'
 $report=@{status='passed';checks=$checks.ToArray();workbook=$path;artifacts=$run;screenshot=$updated.screenshots}
 $report|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $run 'integration-result.json') -Encoding UTF8
 $report|ConvertTo-Json -Depth 6
}finally{
 if($null -eq $book){
  try{$excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application');foreach($candidate in $excel.Workbooks){if($candidate.FullName -eq $path){$book=$candidate;break}}}catch{}
 }
 if($null -ne $book -and $book.FullName -eq $path){$book.Close($false)}
}
