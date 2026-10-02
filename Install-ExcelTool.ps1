#requires -Version 5.1
[CmdletBinding()]
param([string]$CodexSkillsDirectory,[string]$SkillshareSource)
$ErrorActionPreference='Stop'
$source=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot 'excel-tool'))
if(-not (Test-Path -LiteralPath (Join-Path $source 'SKILL.md'))){throw 'Skill source is missing'}
if(-not $CodexSkillsDirectory){
 $codexRoot=$env:CODEX_HOME
 if(-not $codexRoot){$codexRoot=Join-Path $env:USERPROFILE '.codex'}
 $CodexSkillsDirectory=Join-Path $codexRoot 'skills'
}
if(-not $SkillshareSource -and (Get-Command skillshare -ErrorAction SilentlyContinue)){
 $status=(& skillshare status -g --json|Out-String|ConvertFrom-Json)
 if($LASTEXITCODE -eq 0 -and $status.source.exists){$SkillshareSource=$status.source.path}
}
function Register-Junction([string]$root,[string]$linkSource){
 [void][IO.Directory]::CreateDirectory($root)
 $destination=Join-Path ([IO.Path]::GetFullPath($root)) 'excel-tool'
 if(Test-Path -LiteralPath $destination){
  $existing=Get-Item -LiteralPath $destination -Force
  if($existing.LinkType -ne 'Junction'){throw "Existing skill is not this tool's junction: $destination"}
  $existingTarget=[IO.Path]::GetFullPath([string]$existing.Target[0])
  if($existingTarget -ieq $linkSource){return @{path=$destination;status='already_registered';source=$linkSource}}
  if($existingTarget -ine $source){throw "Existing skill points to an unrelated source: $destination"}
  [IO.Directory]::Delete($destination)
 }
 [void](New-Item -ItemType Junction -Path $destination -Target $linkSource)
 return @{path=$destination;status='registered';source=$linkSource}
}
$registered=@()
if($SkillshareSource){
 $installed=Join-Path ([IO.Path]::GetFullPath($SkillshareSource)) 'excel-tool'
 if(Test-Path -LiteralPath $installed){
  $existing=Get-Item -LiteralPath $installed -Force
  if($existing.LinkType){
   if($existing.LinkType -ne 'Junction' -or [IO.Path]::GetFullPath([string]$existing.Target[0]) -ine $source){throw 'Skillshare source has an unrelated link'}
   [IO.Directory]::Delete($installed)
  }elseif(-not (Test-Path -LiteralPath (Join-Path $installed '.excel-tool-install.json'))){throw 'Skillshare contains an unrelated existing skill'}
 }
 [void][IO.Directory]::CreateDirectory($installed)
 $manifestPath=Join-Path $installed '.excel-tool-install.json'
 if(Test-Path -LiteralPath $manifestPath){
  $old=([IO.File]::ReadAllText($manifestPath)|ConvertFrom-Json)
  if($old.source -ine $source){throw 'Installed skill originated from a different repository'}
  foreach($file in $old.files){
   $local=Join-Path $installed $file.path
   if(-not (Test-Path -LiteralPath $local) -or (Get-FileHash -LiteralPath $local -Algorithm SHA256).Hash -ne $file.sha256){throw "Preserving locally modified installed skill: $local"}
  }
 }
 $files=@()
 foreach($file in Get-ChildItem -LiteralPath $source -File -Recurse){
  $relative=$file.FullName.Substring($source.Length+1)
  $destination=Join-Path $installed $relative
  [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($destination))
  Copy-Item -LiteralPath $file.FullName -Destination $destination
  $files+=@{path=$relative;sha256=(Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash}
 }
 @{source=$source;files=$files}|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $manifestPath -Encoding UTF8
 $registered+=@{path=$installed;status='installed_for_skillshare';source=$source}
 $registered+=Register-Junction $CodexSkillsDirectory $installed
}else{$registered+=Register-Junction $CodexSkillsDirectory $source}
@{skill='excel-tool';display_name='Excel Tool';implicit_invocation=$true;registrations=$registered;note='Start a new Codex conversation for skill discovery. Preview Skillshare sync before synchronizing other targets.'}|ConvertTo-Json -Depth 5
