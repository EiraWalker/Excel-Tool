#requires -Version 5.1
[CmdletBinding()]
param(
 [Parameter(Mandatory=$true)][ValidateSet('Compile','Inspect','Apply','Verify')][string]$Operation,
 [Parameter(Mandatory=$true)][string]$Workbook,
 [string]$PlanPath,
 [switch]$Open,
 [switch]$Create,
 [string]$ArtifactsDirectory
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2
if(-not ('CellMathParser' -as [type])){Add-Type -Path (Join-Path $PSScriptRoot 'CellMathParser.cs')}
function Value($obj,[string]$key,$default){
 if($null -ne $obj -and $null -ne $obj.PSObject.Properties[$key]){return $obj.$key}
 return $default
}
function Color([string]$hex){
 if($hex -notmatch '^#[0-9a-fA-F]{6}$'){throw 'Colors must use #RRGGBB'}
 return [Convert]::ToInt32($hex.Substring(1,2),16)+256*[Convert]::ToInt32($hex.Substring(3,2),16)+65536*[Convert]::ToInt32($hex.Substring(5,2),16)
}
function Assert-CellAddress([string]$address){
 if($address -notmatch '^([A-Za-z]{1,3})([1-9][0-9]{0,6})$'){throw 'cell must be a single A1 address without dollar signs'}
 $column=0;foreach($char in $matches[1].ToUpper().ToCharArray()){$column=26*$column+[int]$char-64}
 if($column -gt 16384 -or [int]$matches[2] -gt 1048576){throw 'Cell is outside Excel worksheet limits'}
}
$path=[IO.Path]::GetFullPath($Workbook)
if([IO.Path]::GetExtension($path) -notin @('.xlsx','.xlsm')){throw 'Only .xlsx and .xlsm are supported'}
if($Open -and $Create){throw 'Use either -Open or -Create'}
$compiled=@()
if($PlanPath){
 $plan=([IO.File]::ReadAllText([IO.Path]::GetFullPath($PlanPath))|ConvertFrom-Json)
 $cells=@(Value $plan 'cells' @())
 if($cells.Count -eq 0){throw 'Plan requires nonempty cells'}
 $seen=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
 foreach($item in $cells){
  $sheet=[string](Value $item 'sheet' '');$address=[string](Value $item 'cell' '')
  if(-not $sheet -or $sheet.Length -gt 31 -or $sheet -match '[\\/\?\*\[\]:]'){throw 'Invalid sheet name'}
  Assert-CellAddress $address
  if(-not $seen.Add($sheet+'!'+$address)){throw 'Duplicate target cell'}
  $textFont=[string](Value $item 'text_font' (Value $plan 'text_font' 'MiSans'))
  $mathFont=[string](Value $item 'math_font' (Value $plan 'math_font' 'Cambria'))
  $textSize=[double](Value $item 'text_size' (Value $plan 'text_size' 12))
  $mathSize=[double](Value $item 'math_size' (Value $plan 'math_size' 14))
  foreach($size in @($textSize,$mathSize)){if([double]::IsNaN($size) -or [double]::IsInfinity($size) -or $size -lt 8 -or $size -gt 72){throw 'Font sizes must be 8..72 points'}}
  $textColor=[string](Value $item 'text_color' (Value $plan 'text_color' '#334155'))
  $mathColor=[string](Value $item 'math_color' (Value $plan 'math_color' '#147568'))
  [void](Color $textColor);[void](Color $mathColor)
  $height=Value $item 'row_height' $null
  if($null -ne $height -and ([double]::IsNaN([double]$height) -or [double]$height -lt 16 -or [double]$height -gt 409)){throw 'row_height must be 16..409 points'}
  $replace=Value $item 'replace_text' $false
  if($replace -isnot [bool]){throw 'replace_text must be a JSON boolean'}
  $segments=@(Value $item 'segments' @())
  if($segments.Count -eq 0){throw 'Each cell requires segments'}
  $runs=@();$text=''
  foreach($segment in $segments){
   $hasText=$null -ne $segment.PSObject.Properties['text'];$hasMath=$null -ne $segment.PSObject.Properties['latex']
   if($hasText -eq $hasMath){throw 'Each segment must have exactly one of text or latex'}
   if($hasMath){$parts=@([CellMathParser]::Parse([string]$segment.latex))}
   else{$parts=@([pscustomobject]@{Text=[string]$segment.text;Kind='text'})}
   foreach($part in $parts){
    if($part.Text.Length -eq 0){continue}
    foreach($char in $part.Text.ToCharArray()){if([char]::IsSurrogate($char) -or ([char]::IsControl($char) -and $char -ne "`n")){throw 'Use BMP text; only LF line breaks are supported'}}
    $math=$part.Kind -ne 'text'
    $runs+=[pscustomobject]@{text=$part.Text;kind=$part.Kind;start=$text.Length+1;length=$part.Text.Length;font=$(if($math){$mathFont}else{$textFont});size=$(if($math){$mathSize}else{$textSize});color=$(if($math){$mathColor}else{$textColor})}
    $text+=$part.Text
   }
  }
  if($text.Length -eq 0 -or $text.Length -gt 32767){throw 'Compiled cell text must be 1..32767 characters'}
  $compiled+=[pscustomobject]@{sheet=$sheet;cell=$address.ToUpper();text=$text;runs=$runs;text_font=$textFont;text_size=$textSize;text_color=$textColor;row_height=$height;replace_text=$replace}
 }
}
if($Operation -in @('Compile','Apply','Verify') -and $compiled.Count -eq 0){throw "$Operation requires -PlanPath"}
if($Operation -eq 'Compile'){
 [pscustomobject]@{operation='Compile';mode='cell-richtext';cells=$compiled}|ConvertTo-Json -Depth 10
 return
}
function Verify-SavedCells([string]$file,$expected){
 Add-Type -AssemblyName System.IO.Compression
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $stream=[IO.File]::Open($file,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
 $zip=New-Object IO.Compression.ZipArchive($stream,[IO.Compression.ZipArchiveMode]::Read,$false)
 try{
  function XmlPart([string]$name){
   $entry=$zip.GetEntry($name);if($null -eq $entry){return $null}
   $reader=New-Object IO.StreamReader($entry.Open())
   try{return [xml]$reader.ReadToEnd()}finally{$reader.Dispose()}
  }
  $wb=XmlPart 'xl/workbook.xml';$rels=XmlPart 'xl/_rels/workbook.xml.rels';$strings=XmlPart 'xl/sharedStrings.xml';$styles=XmlPart 'xl/styles.xml'
  $n=New-Object Xml.XmlNamespaceManager($wb.NameTable)
  $n.AddNamespace('s','http://schemas.openxmlformats.org/spreadsheetml/2006/main')
  $n.AddNamespace('r','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
  $items=@();$fallbacks=@()
  foreach($item in $expected){
   $sheet=$null;foreach($candidate in $wb.SelectNodes('/s:workbook/s:sheets/s:sheet',$n)){if($candidate.GetAttribute('name') -ceq $item.sheet){$sheet=$candidate;break}}
   if($null -eq $sheet){throw "Saved sheet missing: $($item.sheet)"}
   $rid=$sheet.GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
   $rel=$null;foreach($candidate in $rels.DocumentElement.ChildNodes){if($candidate.GetAttribute('Id') -eq $rid){$rel=$candidate;break}}
   if($null -eq $rel){throw 'Saved sheet relationship missing'}
   $target=$rel.GetAttribute('Target')
   if($target.StartsWith('/')){$target=$target.TrimStart('/')}else{$target='xl/'+$target}
   $xml=XmlPart $target
   $cell=$xml.SelectSingleNode("/s:worksheet/s:sheetData/s:row/s:c[@r='$($item.cell)']",$n)
   if($null -eq $cell -or $null -ne $cell.SelectSingleNode('s:f',$n)){throw "Saved text cell missing or is a calculation: $($item.cell)"}
   $styleIndex=0;if($cell.HasAttribute('s')){$styleIndex=[int]$cell.GetAttribute('s')}
   $xf=$styles.SelectNodes('/s:styleSheet/s:cellXfs/s:xf',$n).Item($styleIndex)
   $baseFont=$styles.SelectNodes('/s:styleSheet/s:fonts/s:font',$n).Item([int]$xf.GetAttribute('fontId'))
   $baseName=$baseFont.SelectSingleNode('s:name',$n).GetAttribute('val')
   $si=$null
   if($cell.GetAttribute('t') -eq 's'){$si=$strings.SelectNodes('/s:sst/s:si',$n).Item([int]$cell.SelectSingleNode('s:v',$n).InnerText)}
   elseif($cell.GetAttribute('t') -eq 'inlineStr'){$si=$cell.SelectSingleNode('s:is',$n)}
   if($null -eq $si){throw 'Target cell is not a saved string'}
   $actual=@();$savedText=''
   $xmlRuns=@($si.SelectNodes('s:r',$n))
   if($xmlRuns.Count -eq 0){$xmlRuns=@($si)}
   foreach($run in $xmlRuns){
    $t=$run.SelectSingleNode('s:t',$n).InnerText
    $rp=$run.SelectSingleNode('s:rPr',$n)
    $font=$baseName;$kind='baseline';$italic=$null -ne $baseFont.SelectSingleNode('s:i[not(@val) or @val="1"]',$n)
    if($null -ne $rp){
     $f=$rp.SelectSingleNode('s:rFont',$n);if($null -ne $f){$font=$f.GetAttribute('val')}
     $v=$rp.SelectSingleNode('s:vertAlign',$n);if($null -ne $v){$kind=$v.GetAttribute('val')}
     $italic=$null -ne $rp.SelectSingleNode('s:i[not(@val) or @val="1"]',$n)
    }
    foreach($ch in $t.ToCharArray()){$actual+=@{font=$font;kind=$kind;italic=$italic;char=[string]$ch}}
    $savedText+=$t
   }
   if($savedText -cne $item.text){throw "Saved text differs: $($item.sheet)!$($item.cell)"}
   foreach($run in $item.runs){
    for($i=0;$i -lt $run.length;$i++){
     $a=$actual[$run.start+$i-1];$char=$run.text[$i]
     $kind='baseline';if($run.kind -eq 'sub'){$kind='subscript'}elseif($run.kind -eq 'sup'){$kind='superscript'}
     if($a.kind -ne $kind -or $a.italic){throw "Saved upright/script style differs: $($item.cell)"}
     if($a.font -ne $run.font){
      if([char]::IsLetterOrDigit($char)){throw "Saved font differs for letter/digit: $($item.cell)"}
      $fallbacks+=@{sheet=$item.sheet;cell=$item.cell;character=[string]$char;font=$a.font}
     }
    }
   }
   $items+=@{sheet=$item.sheet;cell=$item.cell;text=$savedText;verified=$true}
  }
  return [pscustomobject]@{cells=$items;fallback_symbols=$fallbacks}
 }finally{$zip.Dispose()}
}
if($Operation -eq 'Verify'){
 $verified=Verify-SavedCells $path $compiled
 [pscustomobject]@{operation='Verify';mode='cell-richtext';workbook=$path;verified=$true;cells=$verified.cells;fallback_symbols=$verified.fallback_symbols}|ConvertTo-Json -Depth 8
 return
}
if([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA'){throw 'Use powershell.exe -NoProfile -STA -File ...'}
Add-Type -AssemblyName System.Drawing
$fonts=New-Object Drawing.Text.InstalledFontCollection
foreach($name in @($compiled | ForEach-Object {$_.runs} | ForEach-Object {$_.font} | Select-Object -Unique)){
 if($name -notin @($fonts.Families.Name)){throw "Requested font is not installed: $name"}
}
$excel=$null;$book=$null;$created=$false
if($Create){
 if($Operation -ne 'Apply' -or [IO.Path]::GetExtension($path) -ne '.xlsx' -or (Test-Path -LiteralPath $path)){throw '-Create needs Apply and a new .xlsx path'}
 if(-not(Test-Path -LiteralPath ([IO.Path]::GetDirectoryName($path)))){throw 'Workbook parent directory must exist'}
 $excel=New-Object -ComObject Excel.Application;$excel.Visible=$true;$book=$excel.Workbooks.Add();$created=$true
}else{
 try{$excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')}catch{}
 if($null -ne $excel){foreach($candidate in $excel.Workbooks){if($candidate.FullName -ieq $path){$book=$candidate;break}}}
 if($null -eq $book -and $Open){
  if(-not(Test-Path -LiteralPath $path -PathType Leaf)){throw 'Workbook not found'}
  $book=[Runtime.InteropServices.Marshal]::BindToMoniker($path);$excel=$book.Application
 }
 if($null -eq $book){throw 'Exact workbook not found in the attached instance; use -Open for the specified file'}
 if($book.FullName -ine $path){throw 'Workbook identity mismatch'}
}
if($Operation -eq 'Inspect'){
 [pscustomobject]@{operation='Inspect';mode='cell-richtext';workbook=$book.FullName;saved=$book.Saved;read_only=$book.ReadOnly;sheets=@($book.Worksheets|ForEach-Object {$_.Name})}|ConvertTo-Json -Depth 5
 return
}
if($null -eq $book.Name -or (-not $created -and (-not $book.Saved -or $book.ReadOnly))){throw 'Exit cell editing; workbook must be writable with no unsaved changes'}
try{if($book.AutoSaveOn){throw 'Turn off AutoSave before editing'}}catch{if($_.Exception.Message -like 'Turn off*'){throw}}
$ops=@()
foreach($item in $compiled){
 $sheet=$null;foreach($candidate in $book.Worksheets){if($candidate.Name -ieq $item.sheet){$sheet=$candidate;break}}
 if($null -eq $sheet){
  if(-not $created){throw "Worksheet not found: $($item.sheet)"}
  if($ops.Count -eq 0){$sheet=$book.Worksheets.Item(1)}else{$sheet=$book.Worksheets.Add([Type]::Missing,$book.Worksheets.Item($book.Worksheets.Count))}
  $sheet.Name=$item.sheet
 }
 if($sheet.ProtectContents){throw 'Target worksheet is protected'}
 $cell=$sheet.Range($item.cell)
 if($cell.MergeCells -and $cell.MergeArea.Cells.Item(1,1).Address() -ne $cell.Address()){throw 'Use the top-left cell of a merged area'}
 if($cell.HasFormula){throw 'Cell-math display refuses to replace a calculation formula'}
 if($null -ne $cell.Value2 -and [string]$cell.Value2 -ne '' -and -not $item.replace_text){throw "Nonempty cell requires replace_text=true: $($item.cell)"}
 $ops+=@{item=$item;sheet=$sheet;cell=$cell;shapes=$sheet.Shapes.Count}
}
if(-not $ArtifactsDirectory){$ArtifactsDirectory=Join-Path ([IO.Path]::GetDirectoryName($path)) '.excel-latex-artifacts'}
$run=Join-Path ([IO.Path]::GetFullPath($ArtifactsDirectory)) ('cell-math-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($run)
$backup=$null;$stage=Join-Path $run ('staged'+[IO.Path]::GetExtension($path));$phase='backup'
try{
 if(-not $created){$backup=Join-Path $run ('backup'+[IO.Path]::GetExtension($path));$book.SaveCopyAs($backup)}
 $phase='write'
 foreach($op in $ops){
  $item=$op.item;$cell=$op.cell
  $numberFormat=$cell.NumberFormat
  try{$cell.NumberFormat='@';$cell.Value2=[string]$item.text}finally{$cell.NumberFormat=$numberFormat}
  $cell.Font.Name=$item.text_font;$cell.Font.Size=$item.text_size;$cell.Font.Color=Color $item.text_color
  $cell.Font.Bold=$false;$cell.Font.Italic=$false;$cell.Font.Subscript=$false;$cell.Font.Superscript=$false
  $cell.WrapText=$true;$cell.VerticalAlignment=-4160
  foreach($runStyle in $item.runs){
   $font=$cell.Characters($runStyle.start,$runStyle.length).Font
   $font.Name=$runStyle.font;$font.Size=$runStyle.size;$font.Color=Color $runStyle.color;$font.Bold=$false;$font.Italic=$false
   if($runStyle.kind -eq 'sub'){$font.Subscript=$true}elseif($runStyle.kind -eq 'sup'){$font.Superscript=$true}
  }
  if($null -ne $item.row_height){$cell.EntireRow.RowHeight=[double]$item.row_height}
  if($op.sheet.Shapes.Count -ne $op.shapes){throw 'Unexpected floating-object count change'}
 }
 $phase='staged_verify';$book.SaveCopyAs($stage)
 [void](Verify-SavedCells $stage $compiled)
 $phase='save'
 if($created){$book.SaveAs($path,51)}else{$book.Save()}
 $verified=Verify-SavedCells $path $compiled
 [pscustomobject]@{status='saved_and_verified';mode='cell-richtext';workbook=$path;cell_count=$compiled.Count;backup=$backup;staged=$stage;artifacts=$run;cells=$verified.cells;fallback_symbols=$verified.fallback_symbols}|ConvertTo-Json -Depth 8
}catch{
 [pscustomobject]@{status='failed';phase=$phase;workbook=$path;backup=$backup;error=$_.Exception.Message;note='Review the open workbook and backup; partial in-memory changes may remain.'}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $run 'failure.json') -Encoding UTF8
 throw
}
