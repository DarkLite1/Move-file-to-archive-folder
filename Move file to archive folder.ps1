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

    .EXAMPLE
        See Example.json
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [Int]$MaxConcurrentJobs = 4,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\Move file to archive folder\$ScriptName",
    [String[]]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
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
    
            #region Create filter
            Switch ($OlderThan) {
                'Day' {
                    Filter Select-FileHC {
                        if (
                            $_.CreationTime.Date.ToString('yyyyMMdd') -le $(($Today.AddDays( - $Quantity)).Date.ToString('yyyyMMdd'))
                        ) {
                            Write-Output $_
                        }
                    }
                }
                'Month' {
                    Filter Select-FileHC {
                        if (
                            $_.CreationTime.Date.ToString('yyyyMM') -le $(($Today.AddMonths( - $Quantity)).Date.ToString('yyyyMM'))
                        ) {
                            Write-Output $_
                        }
                    }
                }
                'Year' {
                    Filter Select-FileHC {
                        if (
                            $_.CreationTime.Date.ToString('yyyy') -le $(($Today.AddYears( - $Quantity)).Date.ToString('yyyy'))
                        ) {
                            Write-Output $_
                        }
                    }
                }
                Default {
                    throw "OlderThanUnit '$_' not supported"
                }
            }

            if ($Quantity -eq 0) {
                Filter Select-FileHC {
                    Write-Output $_
                }
            }
            #endregion
        }
    
        Process {
            Get-ChildItem $Source -File | Select-FileHC | ForEach-Object {
                $file = $_

                $result = [PSCustomObject]@{
                    Action                 = $null
                    ComputerName           = $env:COMPUTERNAME
                    SourceFileCreationTime = $file.CreationTime
                    SourceFilePath         = $file.FullName
                    DestinationFolderPath  = $Destination
                    OlderThan              = "$Quantity $OlderThan{0}" -f $(
                        if ($Quantity -gt 1) {
                            's'
                        }
                    )
                    Error                  = $null
                }
    
                $childPath = Switch ($Structure) {
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
                        throw "DestinationFolderStructure '$_' not supported"
                    }
                }

                $joinParams = @{
                    Path      = $Destination
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
                            $global:error.RemoveAt(0)
                            break
                        }
                        default {
                            $result.Error = $_
                            $global:error.RemoveAt(0)
                        }
                    }
                }
                Finally {
                    $result
                }
            }
        }
    }

    Try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
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

            #region ComputerName and non UNC path
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

            #region Local paths without ComputerName
            if (
                (-not $task.ComputerName) -and 
                (
                    ($task.SourceFolderPath -notMatch '^\\\\') -or
                    ($task.DestinationFolderPath -NotMatch '^\\\\')
                )
            ) {
                throw "Input file '$ImportFile' with ComputerName '$($task.ComputerName)' SourceFolderPath '$($task.SourceFolderPath)' and DestinationFolderPath '$($task.DestinationFolderPath)': When local paths are used the ComputerName is mandatory."
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
        #region Start jobs to move files to archive folder
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
            
            $addParams = @{
                NotePropertyMembers = @{
                    Job        = $null
                    JobResults = @()
                    JobErrors  = @()
                }
            }
            $task | Add-Member @addParams

            $task.Job = if ($task.ComputerName) {
                $invokeParams.ComputerName = $task.ComputerName
                $invokeParams.AsJob = $true
                Invoke-Command @invokeParams
            }
            else {
                Start-Job @invokeParams
            }
            
            $waitParams = @{
                Name       = $Tasks.Job | Where-Object { $_ }
                MaxThreads = $MaxConcurrentJobs
            }
            Wait-MaxRunningJobsHC @waitParams
        }
        #endregion

        #region Wait for jobs to finish
        $M = "Wait for all $($Tasks.count) jobs to finish"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        $Tasks.Job | Wait-Job
        #endregion

        #region Get job results and job errors
        foreach ($task in $Tasks) {
            $jobErrors = @()
            $receiveParams = @{
                ErrorVariable = 'jobErrors'
                ErrorAction   = 'SilentlyContinue'
            }
            $task.JobResults += $task.Job | Receive-Job @receiveParams

            foreach ($e in $jobErrors) {
                $task.JobErrors += $e.ToString()
                $Error.Remove($e)

                $M = "Task error on '{0}' with SourceFolderPath '{1}' DestinationFolderPath '{2}' DestinationFolderStructure '{3}' OlderThanUnit '{4}' OlderThanQuantity '{5}': {6}" -f 
                $task.Job.Location, $task.SourceFolderPath, 
                $task.DestinationFolderPath, $task.DestinationFolderStructure, 
                $task.OlderThanUnit, $task.OlderThanQuantity, $e.ToString()
                Write-Verbose $M; Write-EventLog @EventErrorParams -Message $M
            }
        }
        #endregion

        #region Export job results to Excel file
        if ($jobResults = $Tasks.JobResults | Where-Object { $_ }) {
            $M = "Export $($jobResults.Count) rows to Excel"
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
            
            $excelParams = @{
                Path               = $logFile + '- Log.xlsx'
                WorksheetName      = 'Overview'
                TableName          = 'Overview'
                NoNumberConversion = '*'
                AutoSize           = $true
                FreezeTopRow       = $true
            }
            $jobResults | 
            Select-Object -Property * -ExcludeProperty 'PSComputerName',
            'RunSpaceId', 'PSShowComputerName' | 
            Export-Excel @excelParams

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
    try {
        #region Send mail to user

        #region Count results, errors, ...
        $counter = @{
            movedFiles      = (
                $jobResults | 
                Where-Object { $_.Action -like 'File moved*' } | 
                Measure-Object
            ).Count
            moveFilesErrors = ($jobResults.Error | Measure-Object).Count
            jobErrors       = ($Tasks.jobErrors | Measure-Object).Count
            systemErrors    = ($Error.Exception.Message | Measure-Object).Count
        }
        #endregion

        #region Mail subject and priority
        $mailParams.Priority = 'Normal'

        $mailParams.Subject = '{0} file{1} moved' -f $counter.movedFiles, $(
            if ($counter.movedFiles -gt 1) { 's' }
        )

        if (
            $totalErrorCount = $counter.moveFilesErrors + $counter.jobErrors + 
            $counter.systemErrors
        ) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $totalErrorCount error{0}" -f $(
                if ($totalErrorCount -gt 1) { 's' }
            )
        }
        #endregion

        #region Create html lists
        $systemErrorsHtmlList = if ($counter.systemErrors) {
            "<p>Detected <b>{0} non terminating error{1}:{2}</p>" -f $counter.systemErrors, 
            $(
                if ($counter.systemErrors -gt 1) { 's' }
            ),
            $(
                $Error.Exception.Message | Where-Object { $_ } | 
                ConvertTo-HtmlListHC
            )
        }

        $jobResultsHtmlListItems = foreach (
            $task in 
            $Tasks | Sort-Object -Property 'SourceFolderPath'
        ) {
            'From: {0}<br>To: {1}<br>{2}<br>Moved: {3}{4}{5}' -f 
            $(
                if ($task.SourceFolderPath -match '^\\\\') {
                    '<a href="{0}">{0}</a>' -f $task.SourceFolderPath
                }
                else {
                    $uncPath = $task.SourceFolderPath -Replace '^.{2}', (
                        '\\{0}\{1}$' -f $task.ComputerName, $task.SourceFolderPath[0]
                    )
                    '<a href="{0}">{0}</a>' -f $uncPath
                }
            ), 
            $(
                if ($task.DestinationFolderPath -match '^\\\\') {
                    '<a href="{0}">{0}</a>' -f $task.DestinationFolderPath
                }
                else {
                    $uncPath = $task.DestinationFolderPath -Replace '^.{2}', (
                        '\\{0}\{1}$' -f $task.ComputerName, $task.DestinationFolderPath[0]
                    )
                    '<a href="{0}">{0}</a>' -f $uncPath
                }
            ), 
            $(
                if ($task.OlderThanQuantity -eq 0) {
                    'Move all files regardless their creation date'
                }
                else {
                    'Move files older than {0} {1}{2}' -f 
                    $task.OlderThanQuantity,
                    $(
                        $task.OlderThanUnit.ToLower()
                    ),
                    $(
                        if ($task.OlderThanQuantity -gt 1) { 's' }
                    )
                }
            ), 
            $(
                (
                    $task.JobResults | 
                    Where-Object { $_.Action -like 'File moved*' } | 
                    Measure-Object
                ).Count
            ),
            $(
                if ($errorCount = (
                        $task.JobResults | 
                        Where-Object { $_.Error } | 
                        Measure-Object
                    ).Count + $task.JobErrors.Count) {
                    ', <b style="color:red;">errors: {0}</b>' -f $errorCount
                }
            ),
            $(
                if ($task.JobErrors) {
                    $task.JobErrors | ForEach-Object {
                        '<br><b style="color:red;">{0}</b>' -f $_
                    }
                }
            )
        }
   
        $jobResultsHtmlList = $jobResultsHtmlListItems | 
        ConvertTo-HtmlListHC -Spacing Wide
        #endregion
        
        $mailParams += @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Message   = "
                $systemErrorsHtmlList
                <p>Summary:</p>
                $jobResultsHtmlList"
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }

        if ($mailParams.Attachments) {
            $mailParams.Message += 
            "<p><i>* Check the attachment for details</i></p>"
        }
   
        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
        #endregion
    }
    catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}