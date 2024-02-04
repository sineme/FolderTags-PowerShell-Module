[CmdletBinding()]
param()
$ModuleName = 'FolderTags'
$ModuleFile = $ModuleName + '.psm1'
if (Get-Module $ModuleName){
    Remove-Module $ModuleName
}
Join-Path $PSScriptRoot $ModuleFile | Import-Module 
Write-Verbose 'Module imported.'