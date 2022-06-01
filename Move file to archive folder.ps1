<# 
    .SYNOPSIS   
        Move files based on their creation date to folders where the folder
        names are based on the file's creation date. 

    .DESCRIPTION
        When a file in the source folder is older than x days/months/years 
        the file is moved to the destination folder. The name of the destination
        folder is defined by the 'Structure' argument.
        
        The file search is non recursive, only files in the root folder are 
        treated.

    .PARAMETER SourceFolderPath
        Path of the source folder where the file are located.

    .PARAMETER DestinationFolderPath
        Path of the destination folder where the files need be moved too.

    .PARAMETER DestinationFolderStructure
        Name of the folder that needs to be created based on the CreationDate 
        of the file. 
        
        Valid options:
        - Year-Month  : Parent folder '2022-01'
        - Year\\Month : Parent folder '2022' child folder '01'
        - Year        : Parent folder '2022'
        - YYYYMM      : Parent folder '202201'

    .PARAMETER OlderThanUnit
        Combined with OlderThanQuantity this reads:
        OlderThanQuantity = 5
        OlderThanUnit     = 'Day'
        All files older than 5 days will be moved

        Valid options:
        - Day
        - Month
        - Year

    .PARAMETER OlderThanQuantity
        A number to be used in combination with OlderThanUnit
#>

[CmdLetBinding()]
Param (
    [String]$ScriptName = 'Auto Archive',
    [String]$ImportFile,
    [String]$LogFolder = "\\$env:COMPUTERNAME\Log",
    [String]$ScriptAdmin = 'Brecht.Gijbels@heidelbergcement.com'
)

Begin {
    Function Move-ToArchiveHC {
        <#
            .SYNOPSIS
                Moves files to folders based on their creation date.
    
            .DESCRIPTION
                Moves files to folders, where the destination folder names are automatically created based on the file's creation date (year/month). This is useful in situation where files need to be archived by year for example. When a file already exists on the destination it will be overwritten. When a file is in use by another process, we can't move it so we only report that it's in use, no error is thrown.
    
            .PARAMETER Source
                The source folder where we will pick up the files to move them to the destination folder. This folder is only used to pick up files on the root directory, so not recursively.
    
            .PARAMETER Destination
                The destination folder where the files will be moved to. When left empty, the files will be moved to sub folders that will be created in the source folder.
    
            .PARAMETER Structure
                The folder structure that will be used on the destination. The files will be moved based on their creation date. The default value is 'Year-Month'. Valid options are:
    
                'Year'
                C\SourceFolder\2014
                C\SourceFolder\2014\File december.txt
    
                'Year\Month'
                C\SourceFolder\2014
                C\SourceFolder\2014\12
                C\SourceFolder\2014\12\File december.txt
    
                'Year-Month'
                C\SourceFolder\2014-12
                C\SourceFolder\2014-12\File december.txt
    
                'YYYYMM'
                C\SourceFolder\201504
                C\SourceFolder\201504\File.txt
    
            .PARAMETER OlderThan
                This is a filter to only archive files that are older than x days/months/years, where 'x' is defined by the parameter 'Quantity'. When 'OlderThan' and 'Quantity' are not used, all files will be moved an no filtering will take place. Valid options are:
                'Day'
                'Month'
                'Year'
    
            .PARAMETER Quantity
                Quantity defines the number of days/months/years defined for 'OlderThan'. Valid options are only numbers.
    
                -OlderThan Day -Quantity '3'    > All files older than 3 days will be moved
                -OlderThan Month -Quantity '1'  > All files older than 1 month will be moved (all files older than this month will be moved)
                <blanc>                         > All files will be moved, regardless of their creation date
    
            .EXAMPLE
                Move-ToArchiveHC -Source 'T:\Truck movements' -Verbose
                Moves all files based on their creation date from the folder 'T:\Truck movements' to the folders:
                'T:\Truck movements\2014-01\File Jan 2014.txt', 'T:\Truck movements\2014-02\File Feb 2014.txt', ..
    
            .EXAMPLE
                Move-ToArchiveHC -Source 'T:\GPS' -Destination 'C:\Archive' -Structure Year\Month -OlderThan Day -Quantity '3' -Verbose
                Moves all files older than 3 days, based on their creation date, from the folder 'T:\GPS' to the folders:
                'C:\Archive\2014\01\2014-01-01.xml', 'C:\Archive\2014\01\2014-01-02.xml', 'C:\Archive\2014\01\2014-01-03.xml' ..
    
            .NOTES
                CHANGELOG
                2014/09/12 Function born
                2014/09/16 Added filter options 'OlderThan' & 'Quantity'
                2014/09/17 Improved error handling when the destination file is already present
                2014/09/18 Improved help info
                2015/01/20 Improved error reporting
                2015/02/03 Added output in case there is nothing to copy
                2015/04/08 Added 'Structure YYYYMM'
                2017/05/30 Changed to overwrite file in case it's already present on the destination
                           Simplified logging and code
                2017/07/17 Improved logging of destination folder
    
                AUTHOR Brecht.Gijbels@heidelbergcement.com #>
    
        [CmdletBinding(SupportsShouldProcess = $True, DefaultParameterSetName = 'A')]
        Param (
            [parameter(Mandatory = $true, Position = 0, ParameterSetName = 'A')]
            [parameter(Mandatory = $true, Position = 0, ParameterSetName = 'B')]
            [ValidateNotNullOrEmpty()]
            [ValidateScript( { Test-Path $_ -PathType Container })]
            [String]$Source,
            [parameter(Mandatory = $false, Position = 1, ParameterSetName = 'A')]
            [parameter(Mandatory = $false, Position = 1, ParameterSetName = 'B')]
            [ValidateNotNullOrEmpty()]
            [ValidateScript( { Test-Path $_ -PathType Container })]
            [String]$Destination = $Source,
            [parameter(Mandatory = $false, ParameterSetName = 'A')]
            [parameter(Mandatory = $false, ParameterSetName = 'B')]
            [ValidateSet('Year', 'Year\Month', 'Year-Month', 'YYYYMM')]
            [String]$Structure = 'Year-Month',
            [parameter(Mandatory = $true, ParameterSetName = 'B')]
            [ValidateSet('Day', 'Month', 'Year')]
            [String]$OlderThan,
            [parameter(Mandatory = $true, ParameterSetName = 'B')]
            [Int]$Quantity
        )
    
        Begin {
            $Today = Get-Date
    
            Switch ($OlderThan) {
                'Day' {
                    Filter Select-Stuff {
                        Write-Verbose "Found file '$_' with CreationTime '$($_.CreationTime.ToString('dd/MM/yyyy'))'"
                        if ($_.CreationTime.Date.ToString('yyyyMMdd') -le $(($Today.AddDays( - $Quantity)).Date.ToString('yyyyMMdd'))) {
                            Write-Output $_
                        }
                    }
                }
                'Month' {
                    Filter Select-Stuff {
                        Write-Verbose "Found file '$_' with CreationTime '$($_.CreationTime.ToString('dd/MM/yyyy'))'"
                        if ($_.CreationTime.Date.ToString('yyyyMM') -le $(($Today.AddMonths( - $Quantity)).Date.ToString('yyyyMM'))) {
                            Write-Output $_
                        }
                    }
                }
                'Year' {
                    Filter Select-Stuff {
                        Write-Verbose "Found file '$_' with CreationTime '$($_.CreationTime.ToString('dd/MM/yyyy'))'"
                        if ($_.CreationTime.Date.ToString('yyyy') -le $(($Today.AddYears( - $Quantity)).Date.ToString('yyyy'))) {
                            Write-Output $_
                        }
                    }
                }
                Default {
                    Filter Select-Stuff {
                        Write-Verbose "Found file '$_' with CreationTime '$($_.CreationTime.ToString('dd/MM/yyyy'))'"
                        Write-Output $_
                    }
                }
            }
    
            Write-Output @"
        ComputerName: $Env:COMPUTERNAME
        Source:       $Source
        Destination:  $Destination
        Structure:    $Structure
        OlderThan:    $OlderThan
        Quantity:     $Quantity
        Date:         $($Today.ToString('dd/MM/yyyy hh:mm:ss'))
    
        Moved file:
"@
        }
    
        Process {
            $File = $null
    
            Get-ChildItem $Source -File | Select-Stuff | ForEach-Object {
                $File = $_
    
                $ChildPath = Switch ($Structure) {
                    'Year' { 
                        [String]$File.CreationTime.Year 
                        break
                    }
                    'Year\Month' { 
                        [String]$File.CreationTime.Year + '\' + $File.CreationTime.ToString('MM') 
                        break
                    }
                    'Year-Month' { 
                        [String]$File.CreationTime.Year + '-' + $File.CreationTime.ToString('MM') 
                        break
                    }
                    'YYYYMM' { 
                        [String]$File.CreationTime.Year + $File.CreationTime.ToString('MM') 
                        break
                    }
                    Default {
                        throw ""
                    }
                }
                $Target = Join-Path -Path $Destination -ChildPath $ChildPath
    
                Try {
                    $null = New-Item $Target -Type Directory -EA Ignore
                    Move-Item -Path $File.FullName -Destination $Target -EA Stop
                    Write-Output "- '$File' > '$ChildPath'"
                }
                Catch {
                    Switch ($_) {
                        { $_ -match 'cannot access the file because it is being used by another process' } {
                            Write-Output "- '$File' WARNING $_"
                            $Global:Error.RemoveAt(0)
                            break
                        }
                        { $_ -match 'file already exists' } {
                            Move-Item -Path $File.FullName -Destination $Target -Force
                            Write-Output "- '$File' WARNING File already existed on the destination but has now been overwritten"
                            $Global:Error.RemoveAt(0)
                            break
                        }
                        default {
                            Write-Error "Error moving file '$($File.FullName)': $_"
                            $Global:Error.RemoveAt(1)
                            Write-Output "- '$File' ERROR $_"
                        }
                    }
                }
            }
    
            if (-not $File) {
                Write-Output '- INFO No files found that match the filter, nothing moved'
            }
        }
    }
    
    
    Try {
        $HTMLList = $HTMLTargets = $MailTo = $null

        $null = Get-ScriptRuntimeHC -Start

        $ImportFileName = (Get-Item $ImportFile -EA Stop).BaseName
        $ScriptName += ' (' + $ImportFileName + ')'

        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        $File = Get-Content $ImportFile | Remove-CommentsHC
        
        $FunctionFeed = $File | Get-ValueFromArrayHC -Exclude MailTo | 
        ConvertFrom-Csv -Delimiter ',' -Header 'Source', 'Destination', 
        'Structure', 'OlderThan', 'Quantity'
        
        if (-not ($MailTo = $File | Get-ValueFromArrayHC MailTo -Delimiter ',')) {
            throw "No 'MailTo' found in the input file."
        }

        if (-not $FunctionFeed) {
            throw 'The input file was empty.'
        }

        $LogFolder = New-FolderHC -Path $LogFolder -ChildPath "Auto Archive\$ScriptName"
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    ForEach ($Line in $FunctionFeed) {
        Write-Verbose "Source '$($Line.Source)'"
        Write-Verbose "Destination '$($Line.Destination)'"
        Write-Verbose "Structure '$($Line.Structure)'"
        Write-Verbose "OlderThan '$($Line.OlderThan)'"
        Write-Verbose "Quantity '$($Line.Quantity)'`n"

        [Array]$HTMLList += "Source: $(ConvertTo-HTMLlinkHC -Path $Line.Source -Name $Line.Source)<br>
                    Destination: $(ConvertTo-HTMLlinkHC -Path $Line.Destination -Name $Line.Destination)<br>
                    Folder structure: $($Line.Structure)<br>
                    Older than: $($Line.Quantity) $($Line.OlderThan)(s)"

        $MoveParams = @{
            Source      = $Line.Source
            Destination = $Line.Destination
            Structure   = $Line.Structure
            OlderThan   = $Line.OlderThan
            Quantity    = $Line.Quantity
        }

        $MoveParams.Values | ForEach-Object {
            if ($_ -eq $null) {
                Write-Error "Incomplete parameter set, check the input file: $Line."
                Continue
            }
        }
        
        $LogFileDetail = New-LogFileNameHC -LogFolder $LogFolder -Name "$($Line.Source).log" -Date ScriptStartTime -Unique
        
        Try {
            Move-ToArchiveHC @MoveParams *>> $LogFileDetail
        }
        Catch {
            "Move-ToArchiveHC | $(Get-Date -Format "dd/MM/yyyy HH:mm:ss") | ERROR: $_" | Out-File $LogFileDetail
            Write-EventLog @EventErrorParams -Message "Failure:`n`n $_"
        }
    }
}

End {
    Try {
        $HTMLTargets = $HTMLList | ConvertTo-HtmlListHC -Spacing Wide -Header 'The parameters were:'`
            -FootNote "Files are moved from the source to the destination based on their creation date compared to the 'Older than' parameter."
        $HTMLErrors = $Error | ConvertTo-HtmlListHC -Spacing Wide -Header 'Errors detected:'`
            -FootNote "The most common error is that the destination already contains the same file name as the source file, 
            and we don't overwrite files. All other files are correctly moved except for these.<br>Other options for errors are 
            when the source or destination folder is unavailable. This can be caused due to the server being offline, 
            DFS issues or other network related problems."

        if ($Error) {
            $OutParams = @{
                Name    = "$ScriptName - FAILURE"
                Message = $HTMLTargets, $HTMLErrors
            }
            $MailParams = @{
                Message  = $HTMLTargets, $HTMLErrors
                Priority = 'High'
                Subject  = "FAILURE"
            }
        }
        else {
            $OutParams = @{
                Name    = "$ScriptName - Success"
                Message = $HTMLTargets, 'No errors detected.'
            }
            $MailParams = @{
                Message  = $HTMLTargets, 'No errors found'
                Priority = 'Normal'
                Subject  = "Success"
            }
        }

        $null = Get-ScriptRuntimeHC -Stop
        Out-HtmlFileHC @OutParams -Path $LogFolder -NamePrefix ScriptStartTime
        Send-MailHC @MailParams -To $MailTo -LogFolder $LogFolder -Header $ScriptName
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"; Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}