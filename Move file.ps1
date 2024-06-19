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
    [String]$DuplicateFile
)

#region Test source folder
if (-not (Test-Path -LiteralPath $SourceFolder -PathType Container)) {
    throw "SourceFolder '$($SourceFolder)' not found"
}
#endregion

#region Create filter
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
        }
        'Month' {
            Filter Select-FileHC {
                if (
                    $_.CreationTime.Date.ToString('yyyyMM') -le $(($today.AddMonths( - $OlderThanQuantity)).Date.ToString('yyyyMM'))
                ) {
                    Write-Output $_
                }
            }
        }
        'Year' {
            Filter Select-FileHC {
                if (
                    $_.CreationTime.Date.ToString('yyyy') -le $(($today.AddYears( - $OlderThanQuantity)).Date.ToString('yyyy'))
                ) {
                    Write-Output $_
                }
            }
        }
        Default {
            throw "OlderThan.Unit '$_' not supported"
        }
    }
}
#endregion

foreach (
    $file in
    Get-ChildItem $SourceFolder -File | Select-FileHC
) {
    try {
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

        Try {
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
            Move-Item @moveParams

            $result.Action = 'File moved'
        }
        Catch {
            Switch ($_) {
                { $_ -match 'file already exists' } {
                    Move-Item @moveParams -Force
                    $result.Action = 'File moved and overwritten'
                    $error.RemoveAt(0)
                    break
                }
                default {
                    $result.Error = $_
                    $error.RemoveAt(0)
                }
            }
        }
    }
    catch {
        $result.Error = $_
        $error.RemoveAt(0)
    }
    finally {
        $result
    }
}