[CmdletBinding()]
param()
$ModuleName = 'FolderTags'
$ModuleFile = $ModuleName + '.psm1'
$ModuleManifest = $ModuleName + '.psd1'
$Path = $env:PSModulePath -split ";"
$ModulePath = Join-Path $Path[0] $ModuleName
if (Test-Path $ModulePath) {
    Write-Verbose 'Removing old module.'
    Remove-Item $ModulePath -Force -Recurse
}
Write-Verbose 'Creating module directory.'
$null = New-Item $ModulePath -ItemType Directory
Write-Verbose 'Installing module.'
Join-Path $PSScriptRoot $ModuleFile | Copy-Item -Destination $ModulePath -Force
Join-Path $PSScriptRoot $ModuleManifest | Copy-Item -Destination $ModulePath -Force
Write-Verbose 'Module installed.'