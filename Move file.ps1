#Requires -Version 7

Param (
    [Parameter(Mandatory)]
    [String]$SourceFolder,
    [Parameter(Mandatory)]
    [String]$DestinationFolder,
    [Parameter(Mandatory)]
    [ValidateSet('Year', 'Year\Month', 'Year-Month', 'YYYYMM')]
    [String]$DestinationChildFolder,
    [Parameter(Mandatory)]
    [ValidateSet('Day', 'Month', 'Year')]
    [String]$OlderThanUnit,
    [Parameter(Mandatory)]
    [Int]$OlderThanQuantity,
    [Parameter(Mandatory)]
    [Boolean]$Recurse,
    [ValidateSet($null, 'OverwriteFile', 'RenameFile')]
    [String]$DuplicateFile
)

#region Test source folder
if (-not (Test-Path -LiteralPath $SourceFolder -PathType Container)) {
    throw "SourceFolder '$($SourceFolder)' not found"
}
#endregion

#region Create filter
Write-Verbose "Create filter for files with a creation date older than '$OlderThanQuantity $OlderThanUnit'"

if ($OlderThanQuantity -eq 0) {
    Filter Select-FileHC {
        Write-Output $_
    }
}
else {
    $today = Get-Date

    Switch ($OlderThanUnit) {
        'Day' {
            Filter Select-FileHC {
                if (
                    $_.CreationTime.Date.ToString('yyyyMMdd') -le $(($today.AddDays( - $OlderThanQuantity)).Date.ToString('yyyyMMdd'))
                ) {
                    Write-Output $_
                }
            }

            break
        }
        'Month' {
            Filter Select-FileHC {
                if (
                    $_.CreationTime.Date.ToString('yyyyMM') -le $(($today.AddMonths( - $OlderThanQuantity)).Date.ToString('yyyyMM'))
                ) {
                    Write-Output $_
                }
            }

            break
        }
        'Year' {
            Filter Select-FileHC {
                if (
                    $_.CreationTime.Date.ToString('yyyy') -le $(($today.AddYears( - $OlderThanQuantity)).Date.ToString('yyyy'))
                ) {
                    Write-Output $_
                }
            }

            break
        }
        Default {
            throw "OlderThan.Unit '$_' not supported"
        }
    }
}
#endregion

$getParams = @{
    LiteralPath = $SourceFolder
    File        = $true
    Recurse     = $Recurse
}

foreach (
    $file in
    Get-ChildItem @getParams | Select-FileHC
) {
    try {
        Write-Verbose "File '$File'"

        $result = [PSCustomObject]@{
            Action                = $null
            FileName              = $file.Name
            FileCreationTime      = $file.CreationTime
            DestinationFolderPath = $DestinationFolder
            Error                 = $null
        }

        $childPath = Switch ($DestinationChildFolder) {
            'Year' {
                [String]$file.CreationTime.Year
                break
            }
            'Year\Month' {
                [String]$file.CreationTime.Year + '\' + $file.CreationTime.ToString('MM')
                break
            }
            'Year-Month' {
                [String]$file.CreationTime.Year + '-' + $file.CreationTime.ToString('MM')
                break
            }
            'YYYYMM' {
                [String]$file.CreationTime.Year + $file.CreationTime.ToString('MM')
                break
            }
            Default {
                throw "Destination.ChildFolder '$_' not supported"
            }
        }

        $joinParams = @{
            Path      = $DestinationFolder
            ChildPath = $childPath
        }
        $result.DestinationFolderPath = Join-Path @joinParams

        $newParams = @{
            Path        = $result.DestinationFolderPath
            Type        = 'Directory'
            ErrorAction = 'Ignore'
        }
        $null = New-Item @newParams

        $moveParams = @{
            Path        = $file.FullName
            Destination = $result.DestinationFolderPath
            ErrorAction = 'Stop'
        }
        Write-Verbose "Move to '$($moveParams.Destination)'"
        Move-Item @moveParams

        $result.Action = 'File moved'
        Write-Verbose $result.Action
    }
    catch {
        if ($_ -match 'file already exists') {
            Write-Verbose 'Duplicate file name in destination folder'

            $error.RemoveAt(0)

            switch ($DuplicateFile) {
                'OverwriteFile' {
                    try {
                        Write-Verbose 'Overwrite destination file'

                        Move-Item @moveParams -Force
                        $result.Action = 'File moved and overwritten'

                        Write-Verbose $result.Action
                    }
                    catch {
                        $result.Error = "Failed to overwrite file: $_"
                        $error.RemoveAt(0)
                    }
                    Break
                }
                'RenameFile' {
                    try {
                        Write-Verbose 'Create new name for destination file'

                        $newFileName = '{0}_{1}_{2}{3}' -f
                        $file.BaseName,
                        $file.CreationTime.ToString('yyyy-MM-dd-HHmmss'),
                        $(Get-Random -Maximum 999),
                        $file.Extension

                        $joinParams = @{
                            Path      = $moveParams.Destination
                            ChildPath = $newFileName
                        }
                        $moveParams.Destination = Join-Path @joinParams

                        Move-Item @moveParams

                        $result.Action = "File moved with new name '$newFileName' due to duplicate file name"

                        Write-Verbose $result.Action
                    }
                    catch {
                        $result.Error = "Failed to move file with new name '$newFileName': $_"
                        $error.RemoveAt(0)
                    }
                    Break
                }
                Default {
                    $result.Error = "Duplicate file name in destination folder. (See 'Option.DuplicateFile: OverwriteFile or RenameFile')"
                }
            }
        }
        else {
            $result.Error = $_
            $error.RemoveAt(0)
        }

        if ($result.Error) {
            Write-Warning "Error: $($result.Error)"
        }
    }
    finally {
        $result
    }
}