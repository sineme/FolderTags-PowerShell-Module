<#  .SYNOPSIS
    Add-Tag function
    .DESCRIPTION
    Adds a tag to the specified directory
#>
function Add-Tag {

    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelinebyPropertyName = $true)]
        [PSDefaultValue(Help = 'Current directory', Value = '.')]
        [SupportsWildcards()]
        [ValidateScript({ Test-Path $_ -PathType Container }, ErrorMessage = "Path {0} does not exist.")]
        [String[]]$Path = '.',
        [Parameter(
            Mandatory = $true,
            ValueFromPipelinebyPropertyName = $true)]
        [String[]]$Tag
    )
    begin {}
    process {
        $Path | ForEach-Object {
            $iniFilePath = Join-Path $_ 'desktop.ini'
            # create ini if it does not exist
            if (-Not (Test-Path $iniFilePath)) {
                $iniFile = New-Item $iniFilePath -ItemType File
                $iniFile.Attributes += "Hidden"
            }

            # read ini
            $iniContent = Get-IniContent $iniFilePath | Initialize-DesktopIni

            $null = $iniContent['{F29F85E0-4FF9-1068-AB91-08002B27B3D9}']['Prop5'] -match '31,(.*)'
            $iniTag = ($Matches[1] -split '\s*;\s*' | Where-Object { $_ })
            if ($iniTag) { 
                $Tag = $iniTag + $Tag
            }
            $iniContent['{F29F85E0-4FF9-1068-AB91-08002B27B3D9}']['Prop5'] = ($Tag | Select-Object -Unique | Join-String -Separator '; ' -OutputPrefix '31,')
            Set-IniContent $iniFilePath $iniContent
        }
    }
    end {}
}

<#  .SYNOPSIS
    Remove-Tag function
    .DESCRIPTION
    Removes a tag to the specified directory
#>
function Remove-Tag {

    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelinebyPropertyName = $true)]
        [PSDefaultValue(Help = 'Current directory', Value = '.')]
        [SupportsWildcards()]
        [ValidateScript({ Test-Path $_ -PathType Container }, ErrorMessage = "Path {0} does not exist.")]
        [String[]]$Path = '.',
        [Parameter(
            Mandatory = $true,
            ValueFromPipelinebyPropertyName = $true)]
        [String[]]$Tag
    )
    begin {}
    process {
        $Path | ForEach-Object {
            $iniFilePath = Join-Path $_ 'desktop.ini'
            # create ini if it does not exist
            if (-Not (Test-Path $iniFilePath)) {
                $iniFile = New-Item $iniFilePath -ItemType File
                $iniFile.Attributes += "Hidden"
            }

            # read ini
            $iniContent = Get-IniContent $iniFilePath | Initialize-DesktopIni

            $null = $iniContent['{F29F85E0-4FF9-1068-AB91-08002B27B3D9}']['Prop5'] -match '31,(.*)'
            $iniTag = $Matches[1] -split '\s*;\s*'
            $iniTag = $iniTag | Select-Object -Unique | Where-Object { $_ -and $_ -notin $Tag }
            if ($iniTag.count -gt 0) {
                $iniContent['{F29F85E0-4FF9-1068-AB91-08002B27B3D9}']['Prop5'] = ( $iniTag | Join-String -Separator '; ' -OutputPrefix '31,')
            }
            else {
                # no tags, remove property
                $iniContent['{F29F85E0-4FF9-1068-AB91-08002B27B3D9}'].Remove('Prop5')
                if ($iniContent['{F29F85E0-4FF9-1068-AB91-08002B27B3D9}'].count -eq 0) {
                    $iniContent.Remove('{F29F85E0-4FF9-1068-AB91-08002B27B3D9}')
                }
            }

            # check if the file would either be empty or only contain [.ShellClassInfo]
            if ($null -eq $iniContent -or ( `
                        $iniContent.count -eq 1 `
                        -and -not $null -eq $iniContent['.ShellClassInfo'] `
                        -and $iniContent['.ShellClassInfo'].count -eq 0)) {
                # the ini file is empty, so remove it
                Remove-Item $iniFilePath -Force
            }
            else {
                Set-IniContent $iniFilePath $iniContent
            }

        }
    }
    end {}
}

<#  .SYNOPSIS
    .DESCRIPTION
    Ensures the necessary entries exist in the desktop ini
#>
function Initialize-DesktopIni ([Parameter(ValueFromPipeline = $true)]$iniContent) {
    # add section .ShellClassInfo if it does not exist
    if ($null -eq $iniContent['.ShellClassInfo']) {
        $iniContent['.ShellClassInfo'] = @{}
    }

    # add section {F29F85E0-4FF9-1068-AB91-08002B27B3D9} if it does not exist
    if ($null -eq $iniContent['{F29F85E0-4FF9-1068-AB91-08002B27B3D9}']) {
        $iniContent['{F29F85E0-4FF9-1068-AB91-08002B27B3D9}'] = @{}
    }

    # add key Prop5 containing the tags if it does not exist
    if ($null -eq $iniContent['{F29F85E0-4FF9-1068-AB91-08002B27B3D9}']['Prop5']) {
        $iniContent['{F29F85E0-4FF9-1068-AB91-08002B27B3D9}']['Prop5'] = '31,'
    }

    return $iniContent
}

<#
    Code from
    Use PowerShell to Work with Any INI File - by Doctor Scripto, August 20th, 2011
    https://devblogs.microsoft.com/scripting/use-powershell-to-work-with-any-ini-file/
#>
function Get-IniContent ($filePath) {
    $ini = @{}
    switch -regex -file $FilePath {
        “^\[(.+)\]” {
            # Section
            $section = $matches[1]
            $ini[$section] = @{}
            $CommentCount = 0
        }
        “^(;.*)$” {
            # Comment
            $value = $matches[1]
            $CommentCount = $CommentCount + 1
            $name = “Comment” + $CommentCount
            $ini[$section][$name] = $value
        }
        “(.+?)\s*=(.*)” {
            # Key
            $name, $value = $matches[1..2]
            $ini[$section][$name] = $value
        }
    }
    return $ini
}

function Set-IniContent ($filePath, $iniContent) {
    $fileContent = '[.ShellClassInfo]'
    if ($iniContent['.ShellClassInfo']) {
        foreach ($propertyHash in $iniContent['.ShellClassInfo'].GetEnumerator() ) {
            $fileContent += "`r`n" + $propertyHash.Name + '=' + $propertyHash.Value
        }
    }
    foreach ($sectionHash in $iniContent.GetEnumerator()) {
        if ($sectionHash.Name -eq '.ShellClassInfo') {
            continue;
        }
        $fileContent += "`r`n[" + $sectionHash.Name + ']'
        foreach ($propertyHash in $sectionHash.Value.GetEnumerator() ) {
            $fileContent += "`r`n" + $propertyHash.Name + '=' + $propertyHash.Value
        }
    }
    Set-Content $filePath $fileContent
}