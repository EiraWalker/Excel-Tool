#requires -Version 5.1
[CmdletBinding()]
param(
 [Parameter(Mandatory=$true)][ValidateSet('Inspect','Apply','Verify')][string]$Operation,
 [Parameter(Mandatory=$true)][string]$Workbook,
 [string]$PlanPath,
 [switch]$Open,
 [switch]$Create,
 [string]$ArtifactsDirectory,
 [switch]$Capture
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression
$mathNs='http://schemas.openxmlformats.org/officeDocument/2006/math'
$drawingNs='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
$marker='excel-latex/v1'
function Get-Value($object,[string]$name,$fallback) {
 if($null -ne $object -and $null -ne $object.PSObject.Properties[$name]){return $object.$name}
 return $fallback
}
function Read-NativeMath([string]$path) {
 $stream=[IO.File]::Open($path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
 $zip=New-Object IO.Compression.ZipArchive($stream,[IO.Compression.ZipArchiveMode]::Read,$false)
 $items=New-Object Collections.Generic.List[object]
 try{
  foreach($entry in $zip.Entries){
   if($entry.FullName -notmatch '^xl/drawings/[^/]+\.xml$'){continue}
   $reader=New-Object IO.StreamReader($entry.Open())
   try{[xml]$xml=$reader.ReadToEnd()}finally{$reader.Dispose()}
   $ns=New-Object Xml.XmlNamespaceManager($xml.NameTable)
   $ns.AddNamespace('m',$mathNs);$ns.AddNamespace('xdr',$drawingNs)
   foreach($shape in $xml.SelectNodes('//xdr:sp[.//m:oMath]',$ns)){
    $name=$shape.SelectSingleNode('xdr:nvSpPr/xdr:cNvPr',$ns).GetAttribute('name')
    $items.Add([pscustomobject]@{
     name=$name;part=$entry.FullName;math=$shape.SelectNodes('.//m:oMath',$ns).Count
     fractions=$shape.SelectNodes('.//m:f',$ns).Count
     subscripts=$shape.SelectNodes('.//m:sSub|.//m:sSubSup',$ns).Count
     superscripts=$shape.SelectNodes('.//m:sSup|.//m:sSubSup',$ns).Count
     matrices=$shape.SelectNodes('.//m:m',$ns).Count
     equation_arrays=$shape.SelectNodes('.//m:eqArr',$ns).Count
    })
   }
  }
 }finally{$zip.Dispose()}
 return $items.ToArray()
}
function Assert-NativeMath([string]$path,$expected) {
 $native=@(Read-NativeMath $path)
 foreach($item in $expected){
  $match=@($native|Where-Object {$_.name -ceq $item.name})
  if($match.Count -ne 1 -or $match[0].math -lt 1){throw "Native OfficeMath not saved uniquely: $($item.name)"}
  if($item.latex -match '\\frac\b' -and $match[0].fractions -lt 1){throw "Fraction not converted: $($item.name)"}
  if($item.latex -match '\\begin\{cases\}' -and $match[0].matrices -lt 1){throw "Cases not converted: $($item.name)"}
 }
 return $native
}
$workbookPath=[IO.Path]::GetFullPath($Workbook)
if([IO.Path]::GetExtension($workbookPath) -notin @('.xlsx','.xlsm')){throw 'Only .xlsx and .xlsm workbooks are supported'}
$plan=$null
if($PlanPath){$plan=([IO.File]::ReadAllText([IO.Path]::GetFullPath($PlanPath))|ConvertFrom-Json)}
if($Operation -eq 'Apply'){
 if($null -eq $plan -or $null -eq $plan.PSObject.Properties['equations'] -or @($plan.equations).Count -eq 0){throw 'Apply requires a nonempty equations plan'}
 $validatedNames=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
 foreach($item in @($plan.equations)){
  foreach($required in @('name','sheet','anchor','latex')){if(-not (Get-Value $item $required $null)){throw "Equation missing $required"}}
  if($item.name -notmatch '^[A-Za-z][A-Za-z0-9_]{0,79}$' -or -not $validatedNames.Add($item.name)){throw 'Names must be unique ASCII identifiers (maximum 80 characters)'}
  if($item.sheet.Length -gt 31 -or $item.sheet -match '[\\/\?\*\[\]:]'){throw 'Invalid Excel worksheet name'}
  if($item.anchor -notmatch '^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?$'){throw 'Only A1 rectangles are allowed as anchors'}
  if($null -ne $item.PSObject.Properties['clear_text'] -and $item.clear_text -isnot [bool]){throw 'clear_text must be a JSON boolean'}
  $latex=[string]$item.latex
  if($latex.Length -gt 20000 -or $latex -match '\\(documentclass|usepackage|input|include|write|begin\{document\})' -or $latex.StartsWith('\[') -or $latex.StartsWith('$')){throw 'Provide a math expression without document commands or outer delimiters'}
  $size=[double](Get-Value $item 'font_size' 17)
  if($size -lt 8 -or $size -gt 72){throw 'font_size must be between 8 and 72 points'}
  foreach($dimension in @('width','height')){if($null -ne $item.PSObject.Properties[$dimension] -and [double]$item.$dimension -lt 16){throw "$dimension must be at least 16 points"}}
  $height=Get-Value $item 'row_height' $null
  if($null -ne $height -and ([double]$height -lt 16 -or [double]$height -gt 409)){throw 'row_height must be 16..409 points'}
 }
}
if($Operation -eq 'Verify'){
 if(-not (Test-Path -LiteralPath $workbookPath -PathType Leaf)){throw 'Workbook file not found'}
 if($plan){$native=@(Assert-NativeMath $workbookPath $plan.equations)}else{$native=@(Read-NativeMath $workbookPath)}
 [pscustomobject]@{operation='Verify';workbook=$workbookPath;native_equation_count=$native.Count;equations=$native}|ConvertTo-Json -Depth 8
 return
}
if([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA'){throw 'Run with Windows PowerShell 5.1: powershell.exe -NoProfile -STA -File ...'}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
if(-not ('ExcelLatexNativeWindow' -as [type])){
 Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ExcelLatexNativeWindow {
 [StructLayout(LayoutKind.Sequential)] public struct RECT {public int left,top,right,bottom;}
 [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h,out RECT r);
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
 [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h,int c);
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
 [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a,uint b,bool c);
 [DllImport("user32.dll")] public static extern IntPtr SetFocus(IntPtr h);
}
'@
}
try{$excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')}
catch{
 if(-not ($Open -or $Create)){throw 'Excel is not running. Use -Open or -Create when application startup is intended.'}
 $excel=New-Object -ComObject Excel.Application
 $excel.Visible=$true
}
$book=$null
foreach($candidate in $excel.Workbooks){if($candidate.FullName -ieq $workbookPath){$book=$candidate;break}}
$created=$false
if($null -eq $book){
 if($Create){
  if(Test-Path -LiteralPath $workbookPath){throw '-Create refuses to overwrite an existing file'}
  if([IO.Path]::GetExtension($workbookPath) -ne '.xlsx'){throw 'New workbooks must use .xlsx'}
  if($Operation -ne 'Apply' -or $null -eq $plan){throw '-Create requires Apply and a plan'}
  if(-not (Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($workbookPath)))){throw 'Workbook parent directory must exist'}
  $book=$excel.Workbooks.Add();$created=$true
 }elseif($Open){
  if(-not (Test-Path -LiteralPath $workbookPath -PathType Leaf)){throw 'Workbook not found'}
  $book=$excel.Workbooks.Open($workbookPath)
 }else{throw 'Exact workbook is not open in the attached Excel instance. Use -Open to open the named file.'}
}
function Assert-Target {
 if(-not [object]::ReferenceEquals($excel.ActiveWorkbook,$book)){
  if($excel.ActiveWorkbook.FullName -ine $book.FullName){throw 'Active workbook changed'}
 }
}
function Focus-Target {
 [void]$book.Activate()
 $window=[IntPtr]$excel.ActiveWindow.Hwnd
 $pidValue=[uint32]0
 $targetThread=[ExcelLatexNativeWindow]::GetWindowThreadProcessId($window,[ref]$pidValue)
 $foregroundThread=[ExcelLatexNativeWindow]::GetWindowThreadProcessId([ExcelLatexNativeWindow]::GetForegroundWindow(),[ref]$pidValue)
 $currentThread=[ExcelLatexNativeWindow]::GetCurrentThreadId()
 $attachedForeground=$false;$attachedTarget=$false
 try{
  if($foregroundThread -ne $currentThread){$attachedForeground=[ExcelLatexNativeWindow]::AttachThreadInput($currentThread,$foregroundThread,$true)}
  if($targetThread -ne $currentThread -and $targetThread -ne $foregroundThread){$attachedTarget=[ExcelLatexNativeWindow]::AttachThreadInput($currentThread,$targetThread,$true)}
  [void][ExcelLatexNativeWindow]::ShowWindow($window,9)
  [void][ExcelLatexNativeWindow]::BringWindowToTop($window)
  [void][ExcelLatexNativeWindow]::SetForegroundWindow($window)
  [void][ExcelLatexNativeWindow]::SetFocus($window)
 }finally{
  if($attachedTarget){[void][ExcelLatexNativeWindow]::AttachThreadInput($currentThread,$targetThread,$false)}
  if($attachedForeground){[void][ExcelLatexNativeWindow]::AttachThreadInput($currentThread,$foregroundThread,$false)}
 }
 Start-Sleep -Milliseconds 150
 Assert-Target
 if([ExcelLatexNativeWindow]::GetForegroundWindow() -ne $window){throw "Actual Excel foreground focus not established; no paste attempted (expected=$window actual=$([ExcelLatexNativeWindow]::GetForegroundWindow()) target_thread=$targetThread foreground_thread=$foregroundThread current_thread=$currentThread attached_foreground=$attachedForeground attached_target=$attachedTarget)"}
}
function Get-Shapes($sheet) {
 $result=New-Object Collections.Generic.List[object]
 foreach($shape in $sheet.Shapes){
  $text='';try{$text=$shape.TextFrame2.TextRange.Text}catch{}
  $result.Add([pscustomobject]@{name=$shape.Name;managed=($shape.AlternativeText.StartsWith($marker));text=$text;left=$shape.Left;top=$shape.Top;width=$shape.Width;height=$shape.Height})
 }
 return $result.ToArray()
}
if($Operation -eq 'Inspect'){
 $sheets=@();foreach($sheet in $book.Worksheets){$sheets+=@{name=$sheet.Name;protected=[bool]$sheet.ProtectContents;used_range=$sheet.UsedRange.Address();shapes=@(Get-Shapes $sheet)}}
 [pscustomobject]@{operation='Inspect';workbook=$book.FullName;saved=[bool]$book.Saved;read_only=[bool]$book.ReadOnly;excel_version=$excel.Version;excel_build=$excel.Build;sheets=$sheets}|ConvertTo-Json -Depth 8
 return
}
if($null -eq $plan -or $null -eq $plan.PSObject.Properties['equations']){throw 'Apply requires a JSON plan containing equations'}
$equations=@($plan.equations)
if($equations.Count -eq 0){throw 'Plan has no equations'}
if(-not $created -and (-not $book.Saved -or $book.ReadOnly)){throw 'Workbook must be writable and have no unsaved changes'}
if([int]$excel.Build -lt 20131){throw 'Excel build is below the documented native LaTeX minimum (20131)'}
try{if($book.AutoSaveOn){throw 'Turn off AutoSave for this workbook before a batch edit'}}catch{if($_.Exception.Message -like 'Turn off*'){throw}}
$nameSet=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$operations=New-Object Collections.Generic.List[object]
$rows=New-Object Collections.Generic.List[object]
foreach($item in $equations){
 foreach($required in @('name','sheet','anchor','latex')){if(-not (Get-Value $item $required $null)){throw "Equation missing $required"}}
 if($item.name -notmatch '^[A-Za-z][A-Za-z0-9_]{0,79}$' -or -not $nameSet.Add($item.name)){throw 'Names must be unique ASCII identifiers (maximum 80 characters)'}
 $latex=[string]$item.latex
 if($latex.Length -gt 20000 -or $latex -match '\\(documentclass|usepackage|input|include|write|begin\{document\})' -or $latex.StartsWith('\[') -or $latex.StartsWith('$')){throw 'Provide a math expression without document commands or outer delimiters'}
 $sheet=$null;try{$sheet=$book.Worksheets.Item([string]$item.sheet)}catch{}
 if($null -eq $sheet){
  if(-not $created){throw "Worksheet not found: $($item.sheet)"}
  if($operations.Count -eq 0){$sheet=$book.Worksheets.Item(1);$sheet.Name=$item.sheet}
  else{$sheet=$book.Worksheets.Add([Type]::Missing,$book.Worksheets.Item($book.Worksheets.Count));$sheet.Name=$item.sheet}
 }
 if($sheet.ProtectContents -or $sheet.ProtectDrawingObjects){throw "Worksheet is protected: $($item.sheet)"}
 $anchor=[string]$item.anchor
 if($anchor -notmatch '^\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?$'){throw "Only A1 rectangles are allowed as anchors: $anchor"}
 $range=$sheet.Range($anchor)
 if($range.Cells.Count -eq 1 -and $range.MergeCells){$range=$range.MergeArea}
 if($range.Cells.Count -gt 10000){throw 'Anchor is too large'}
 $fontSize=[double](Get-Value $item 'font_size' 17)
 if($fontSize -lt 8 -or $fontSize -gt 72){throw 'font_size must be between 8 and 72 points'}
 foreach($dimension in @('width','height')){if($null -ne $item.PSObject.Properties[$dimension] -and [double]$item.$dimension -lt 16){throw "$dimension must be at least 16 points"}}
 $rowHeight=Get-Value $item 'row_height' $null
 if($null -ne $rowHeight -and ([double]$rowHeight -lt 16 -or [double]$rowHeight -gt 409)){throw 'row_height must be 16..409 points'}
 $existing=$null
 foreach($shape in $sheet.Shapes){if($shape.Name -ieq $item.name){$existing=$shape;break}}
 if($existing -and -not $existing.AlternativeText.StartsWith($marker)){throw "Refusing to replace an unmanaged shape: $($item.name)"}
 $operations.Add([pscustomobject]@{item=$item;sheet=$sheet;range=$range;existing=$existing;new_shape=$null;font_size=$fontSize;row_height=$rowHeight})
}
if($created){$book.SaveAs($workbookPath,51)}
if(-not $ArtifactsDirectory){$ArtifactsDirectory=Join-Path ([IO.Path]::GetDirectoryName($workbookPath)) '.excel-latex-artifacts'}
$artifactRoot=[IO.Path]::GetFullPath($ArtifactsDirectory)
$runDirectory=Join-Path $artifactRoot ((Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[guid]::NewGuid().ToString('N').Substring(0,8))
[void][IO.Directory]::CreateDirectory($runDirectory)
$backupPath=Join-Path $runDirectory ('before'+[IO.Path]::GetExtension($workbookPath))
if(-not $created){$book.SaveCopyAs($backupPath)}else{$backupPath=$null}
$stagedPath=Join-Path $runDirectory ('staged'+[IO.Path]::GetExtension($workbookPath))
$clipboard=$null;$clipboardRestored=$false;$phase='stage';$report=$null
try{
 $clipboard=[Windows.Forms.Clipboard]::GetDataObject()
 foreach($op in $operations){
  [void]$op.sheet.Activate()
  if($null -ne $op.row_height){
   foreach($row in $op.range.Rows){$rows.Add(@{sheet=$op.sheet;index=$row.Row;height=$row.RowHeight})}
   $op.range.EntireRow.RowHeight=[double]$op.row_height
  }
  for($attempt=1;$attempt -le 2;$attempt++){
   try{
    Focus-Target
    [void]$op.sheet.Activate()
    [void]$op.range.Cells.Item(1,1).Select()
    $beforeCount=$op.sheet.Shapes.Count
    if(-not $excel.CommandBars.GetEnabledMso('EquationInsertNew')){throw 'EquationInsertNew is not enabled'}
    $excel.CommandBars.ExecuteMso('EquationInsertNew')
    if($op.sheet.Shapes.Count -ne ($beforeCount+1)){throw 'Equation did not create exactly one shape'}
    $op.new_shape=$op.sheet.Shapes.Item($op.sheet.Shapes.Count)
    $op.new_shape.Name='XLatexStage_'+[guid]::NewGuid().ToString('N')
    if(-not $excel.CommandBars.GetEnabledMso('EquationProfessional')){throw 'Native equation editing context missing'}
    [Windows.Forms.Clipboard]::SetText(('\['+$op.item.latex+'\]'),[Windows.Forms.TextDataFormat]::UnicodeText)
    Assert-Target
    if(-not $excel.CommandBars.GetEnabledMso('EquationProfessional')){throw 'Equation editing context changed before paste; paste cancelled'}
    $excel.CommandBars.ExecuteMso('Paste')
    Start-Sleep -Milliseconds 200
    $excel.CommandBars.ExecuteMso('EquationProfessional')
    [void]$op.range.Cells.Item(1,1).Select()
    break
   }catch{
    if($attempt -eq 2 -or $_.Exception.Message -notmatch 'focus|context|not enabled|E_FAIL|HRESULT'){throw}
    if($op.new_shape){$op.new_shape.Delete();$op.new_shape=$null}
    Start-Sleep -Milliseconds 150
   }
  }
  $shape=$op.new_shape
  if($shape.TextFrame2.TextRange.Text -match '\\[a-zA-Z]+|在此处键入公式|Type equation here'){throw "LaTeX remained unconverted: $($op.item.name)"}
  $shape.TextFrame2.TextRange.Font.Name='Cambria Math'
  $shape.TextFrame2.TextRange.Font.Size=$op.font_size
  $shape.TextFrame2.MarginLeft=8;$shape.TextFrame2.MarginRight=8
  $shape.TextFrame2.MarginTop=3;$shape.TextFrame2.MarginBottom=3
  $shape.TextFrame2.AutoSize=0;$shape.TextFrame2.WordWrap=0
  $shape.Left=$op.range.Left+3;$shape.Top=$op.range.Top+2
  $shape.Width=[double](Get-Value $op.item 'width' ([Math]::Max(16,$op.range.Width-6)))
  $shape.Height=[double](Get-Value $op.item 'height' ([Math]::Max(16,$op.range.Height-4)))
  $shape.Placement=1;$shape.Fill.Visible=0;$shape.Line.Visible=0
  $shape.AlternativeText=$marker
 }
 $book.SaveCopyAs($stagedPath)
 $stagedExpected=@($operations|ForEach-Object {@{name=$_.new_shape.Name;latex=$_.item.latex}})
 [void](Assert-NativeMath $stagedPath $stagedExpected)
 $phase='commit'
 foreach($op in $operations){
  if($op.existing){$op.existing.Delete()}
  $op.new_shape.Name=$op.item.name
  if([bool](Get-Value $op.item 'clear_text' $false)){[void]$op.range.ClearContents()}
 }
 if($created){$book.SaveAs($workbookPath,51)}else{$book.Save()}
 if(-not $book.Saved){throw 'Excel did not confirm saving'}
 $native=@(Assert-NativeMath $workbookPath $equations)
 $phase='capture'
 $screenshots=New-Object Collections.Generic.List[string]
 if($Capture){
  $seenSheets=New-Object 'Collections.Generic.HashSet[string]'
  foreach($op in $operations){
   if(-not $seenSheets.Add($op.sheet.Name)){continue}
   Focus-Target;[void]$op.sheet.Activate()
   $excel.ActiveWindow.ScrollRow=[Math]::Max(1,$op.range.Row-1)
   $excel.ActiveWindow.ScrollColumn=[Math]::Max(1,$op.range.Column)
   Start-Sleep -Milliseconds 300
   $rect=New-Object ExcelLatexNativeWindow+RECT
   [void][ExcelLatexNativeWindow]::GetWindowRect([IntPtr]$excel.ActiveWindow.Hwnd,[ref]$rect)
   $bounds=[Windows.Forms.SystemInformation]::VirtualScreen
   $left=[Math]::Max($rect.left,$bounds.Left);$top=[Math]::Max($rect.top,$bounds.Top)
   $width=[Math]::Min($rect.right,$bounds.Right)-$left;$height=[Math]::Min($rect.bottom,$bounds.Bottom)-$top
   $bitmap=New-Object Drawing.Bitmap($width,$height)
   $graphics=[Drawing.Graphics]::FromImage($bitmap)
   $file=Join-Path $runDirectory ('sheet-'+$screenshots.Count+'.png')
   try{$graphics.CopyFromScreen($left,$top,0,0,$bitmap.Size);$bitmap.Save($file,[Drawing.Imaging.ImageFormat]::Png)}finally{$graphics.Dispose();$bitmap.Dispose()}
   $screenshots.Add($file)
  }
 }
 $report=[pscustomobject]@{operation='Apply';workbook=$workbookPath;status='saved_and_verified';excel_build=$excel.Build;backup=$backupPath;staged_file=$stagedPath;equation_count=$equations.Count;native_equations=$native;screenshots=$screenshots.ToArray();artifacts=$runDirectory}
}catch{
 $failure=$_
 if($phase -eq 'stage'){
  foreach($op in $operations){if($op.new_shape){try{$op.new_shape.Delete()}catch{}}}
  for($i=$rows.Count-1;$i -ge 0;$i--){try{$rows[$i].sheet.Rows.Item($rows[$i].index).RowHeight=$rows[$i].height}catch{}}
 }
 $errorReport=@{status='failed';phase=$phase;workbook=$workbookPath;backup=$backupPath;artifacts=$runDirectory;message=$failure.Exception.Message;target_saved=$book.Saved}
 $errorReport|ConvertTo-Json -Depth 6|Set-Content -LiteralPath (Join-Path $runDirectory 'failure.json') -Encoding UTF8
 throw "Excel LaTeX failed during $phase. $($failure.Exception.Message) Backup: $backupPath; details: $runDirectory"
}finally{
 if($clipboard){try{[Windows.Forms.Clipboard]::SetDataObject($clipboard,$true);$clipboardRestored=$true}catch{}}
}
$report|Add-Member -NotePropertyName clipboard_restored -NotePropertyValue $clipboardRestored
$report|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $runDirectory 'result.json') -Encoding UTF8
$report|ConvertTo-Json -Depth 8
