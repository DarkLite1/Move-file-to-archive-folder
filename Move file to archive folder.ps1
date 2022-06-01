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
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [Int]$MaxConcurrentJobs = 4,
    [String]$LogFolder = $env:POWERSHELL_LOG_FOLDER,
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    $scriptBlock = {    
        Param (
            [Parameter(Mandatory)]
            [ValidateScript( { Test-Path $_ -PathType Container })]
            [String]$Source,
            [Parameter(Mandatory)]
            [ValidateScript( { Test-Path $_ -PathType Container })]
            [String]$Destination,
            [Parameter(Mandatory)]
            [ValidateSet('Year', 'Year\Month', 'Year-Month', 'YYYYMM')]
            [String]$Structure,
            [Parameter(Mandatory)]
            [ValidateSet('Day', 'Month', 'Year')]
            [String]$OlderThan,
            [Parameter(Mandatory)]
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
    
            [PSCustomObject]@{
                ComputerName = $Env:COMPUTERNAME
                Source       = $Source
                Destination  = $Destination
                Structure    = $Structure
                OlderThan    = $OlderThan
                Quantity     = $Quantity
                Date         = $($Today.ToString('dd/MM/yyyy hh:mm:ss'))
            }

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
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        #region Logging
        try {
            $LogParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $LogFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        $file = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json
        #endregion

        #region Test .json file properties
        if (-not ($MailTo = $file.MailTo)) {
            throw "Input file '$ImportFile': No 'MailTo' addresses found."
        }
        if (-not ($Tasks = $file.Tasks)) {
            throw "Input file '$ImportFile': No 'Tasks' found."
        }
        foreach ($task in $Tasks) {
            #region SourceFolderPath
            if (-not $task.SourceFolderPath) {
                throw "Input file '$ImportFile': No 'SourceFolderPath' found in one of the 'Tasks'."
            }
            #endregion

            #region DestinationFolderPath
            if (-not $task.DestinationFolderPath) {
                throw "Input file '$ImportFile': No 'DestinationFolderPath' found in one of the 'Tasks'."
            }
            #endregion

            #region DestinationFolderStructure
            if (-not $task.DestinationFolderStructure) {
                throw "Input file '$ImportFile': No 'DestinationFolderStructure' found in one of the 'Tasks'."
            }

            if ($task.DestinationFolderStructure -notMatch '^Year-Month$|^Year\\Month$|^Year$|^YYYYMM$') {
                throw "Input file '$ImportFile': Value '$($task.DestinationFolderStructure)' is not supported by 'DestinationFolderStructure'. Valid options are 'Year-Month', 'Year\Month', 'Year' or 'YYYYMM'."
            }
            #endregion

            #region OlderThanUnit
            if (-not $task.OlderThanUnit) {
                throw "Input file '$ImportFile': No 'OlderThanUnit' found in one of the 'Tasks'."
            }

            if ($task.OlderThanUnit -notMatch '^Day$|^Month$|^Year$') {
                throw "Input file '$ImportFile': Value '$($task.OlderThanUnit)' is not supported by 'OlderThanUnit'. Valid options are 'Day', 'Month' or 'Year'."
            }
            #endregion

            #region OlderThanQuantity
            if ($task.PSObject.Properties.Name -notContains 'OlderThanQuantity') {
                throw "Input file '$ImportFile' SourceFolderPath '$($task.SourceFolderPath)': Property 'OlderThanQuantity' not found. Use value number '0' to move all files."
            }
            if (-not ($task.OlderThanQuantity -is [int])) {
                throw "Input file '$ImportFile' SourceFolderPath '$($task.SourceFolderPath)': Property 'OlderThanQuantity' needs to be a number, the value '$($task.OlderThanQuantity)' is not supported. Use value number '0' to move all files."
            }
            #endregion

            #region SourceFolderPath
            if (
                $task.ComputerName -and 
                (
                    ($task.SourceFolderPath -Match '^\\\\') -or
                    ($task.DestinationFolderPath -Match '^\\\\')
                )
            ) {
                throw "Input file '$ImportFile' with ComputerName '$($task.ComputerName)' SourceFolderPath '$($task.SourceFolderPath)' and DestinationFolderPath '$($task.DestinationFolderPath)': When ComputerName is used only local paths are allowed (to avoid the double hop issue)."
            }
            #endregion
        }
        #endregion

        $mailParams = @{ }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        #region Remove files/folders on remote machines
        $jobs = @()

        foreach ($task in $Tasks) {
            $invokeParams = @{
                ScriptBlock  = $scriptBlock
                ArgumentList = $task.SourceFolderPath, 
                $task.DestinationFolderPath, 
                $task.DestinationFolderStructure, 
                $task.OlderThanUnit, 
                $task.OlderThanQuantity
            }

            $M = "Start job on '{0}' with SourceFolderPath '{1}' DestinationFolderPath '{2}' DestinationFolderStructure '{3}' OlderThanUnit '{4}' OlderThanQuantity '{5}'" -f $env:COMPUTERNAME,
            $invokeParams.ArgumentList[0], $invokeParams.ArgumentList[1],
            $invokeParams.ArgumentList[2], $invokeParams.ArgumentList[3], 
            $invokeParams.ArgumentList[4]
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            $jobs += if ($task.ComputerName) {
                $invokeParams.ComputerName = $task.ComputerName
                $invokeParams.AsJob = $true
                Invoke-Command @invokeParams
            }
            else {
                Start-Job @invokeParams
            }
            
            # & $scriptBlock -Type $task.Remove -Path $task.Path -OlderThanDays $task.OlderThanDays -RemoveEmptyFolders $task.RemoveEmptyFolders

            Wait-MaxRunningJobsHC -Name $jobs -MaxThreads $MaxConcurrentJobs
        }

        $M = "Wait for all $($jobs.count) jobs to finish"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        # $jobResults = if ($jobs) { $jobs | Wait-Job -Force | Receive-Job }
        $jobResults = if ($jobs) { 
            Receive-Job -Job $jobs -Wait -AutoRemoveJob -Force 
        }
        #endregion

        #region Export results to Excel log file
        $exportToExcel = foreach (
            $job in 
            $jobResults | Where-Object { $_.Items }
        ) {
            $job.Items | Select-Object -Property @{
                Name       = 'ComputerName'; 
                Expression = { $job.ComputerName } 
            },
            'Type', @{
                Name       = 'Path'; 
                Expression = { $_.FullName } 
            }, 'CreationTime', 'Action', 'Error'
        }

        if ($exportToExcel) {
            $M = "Export $($exportToExcel.Count) rows to Excel"
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
            
            $excelParams = @{
                Path               = $LogFile + '- Log.xlsx'
                WorksheetName      = 'Overview'
                TableName          = 'Overview'
                NoNumberConversion = '*'
                AutoSize           = $true
                FreezeTopRow       = $true
            }
            $exportToExcel | Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
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