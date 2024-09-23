<#PSScriptInfo .VERSION 1.0.0#>

using namespace System.IO
using namespace System.Runtime.InteropServices
[CmdletBinding()]
param ()

& {
  Import-Module "$PSScriptRoot\tools"
  Format-ProjectCode @('*.js','*.ps*1','.gitignore'| ForEach-Object { "$PSScriptRoot\$_" })
  Remove-Module tools
  $shell = New-Object -ComObject 'WScript.Shell'
  $shell.CreateShortcut("$PSScriptRoot\MsgBox.lnk") | ForEach-Object {
    $_.TargetPath = 'C:\Windows\System32\wscript.exe'
    $_.Arguments = [Path]::Combine($PSScriptRoot, 'src\messageBox.js')
    $_.IconLocation = [Path]::Combine($PSScriptRoot, 'rsc\menu.ico')
    $_.Save()
    [void][Marshal]::FinalReleaseComObject($_)
    $_ = $null
  }
  [void][Marshal]::FinalReleaseComObject($shell)
  $shell = $null
  [GC]::Collect()
}