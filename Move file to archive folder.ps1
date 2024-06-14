#Requires -Version 5.1
#Requires -Modules ImportExcel
#Requires -Modules Toolbox.EventLog, Toolbox.HTML, Toolbox.Remoting

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

    .PARAMETER OlderThan.Unit
        Combined with OlderThan.Quantity this reads:
        OlderThan.Quantity = 5
        OlderThan.Unit     = 'Day'
        All files older than 5 days will be moved

        Valid options:
        - Day
        - Month
        - Year

    .PARAMETER OlderThan.Quantity
        A number to be used in combination with OlderThan.Unit

    .PARAMETER PSSessionConfiguration
        The version of PowerShell on the remote endpoint as returned by
        Get-PSSessionConfiguration.

    .EXAMPLE
        See Example.json
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$PSSessionConfiguration = 'PowerShell.7',
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\Move file to archive folder\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
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
                    throw "OlderThan.Unit '$_' not supported"
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

        if (-not ($MaxConcurrentJobs = $file.MaxConcurrentJobs)) {
            throw "Property 'MaxConcurrentJobs' not found"
        }
        try {
            $null = $MaxConcurrentJobs.ToInt16($null)
        }
        catch {
            throw "Property 'MaxConcurrentJobs' needs to be a number, the value '$MaxConcurrentJobs' is not supported."
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

            #region OlderThan.Unit
            if (-not $task.OlderThan.Unit) {
                throw "Input file '$ImportFile': No 'OlderThan.Unit' found in one of the 'Tasks'."
            }

            if ($task.OlderThan.Unit -notMatch '^Day$|^Month$|^Year$') {
                throw "Input file '$ImportFile': Value '$($task.OlderThan.Unit)' is not supported by 'OlderThan.Unit'. Valid options are 'Day', 'Month' or 'Year'."
            }
            #endregion

            #region OlderThan.Quantity
            if ($task.PSObject.Properties.Name -notContains 'OlderThan') {
                throw "Input file '$ImportFile' SourceFolderPath '$($task.SourceFolderPath)': Property 'OlderThan' with 'Quantity' and 'Unit' not found."
            }

            if ($task.OlderThan.PSObject.Properties.Name -notContains 'Quantity') {
                throw "Input file '$ImportFile' SourceFolderPath '$($task.SourceFolderPath)': Property 'OlderThan.Quantity' not found. Use value number '0' to move all files."
            }

            try {
                $null = [int]$task.OlderThan.Quantity
            }
            catch {
                throw "Input file '$ImportFile' SourceFolderPath '$($task.SourceFolderPath)': Property 'OlderThan.Quantity' needs to be a number, the value '$($task.OlderThan.Quantity)' is not supported. Use value number '0' to move all files."
            }
            #endregion
        }
        #endregion

        #region Convert .json file
        foreach ($task in $Tasks) {
            #region Set ComputerName if there is none
            if (
            (-not $task.ComputerName) -or
            ($task.ComputerName -eq 'localhost') -or
            ($task.ComputerName -eq "$ENV:COMPUTERNAME.$env:USERDNSDOMAIN")
            ) {
                $task.ComputerName = $env:COMPUTERNAME
            }
            #endregion

            #region Add properties
            $task | Add-Member -NotePropertyMembers @{
                Job = @{
                    Object  = $null
                    Results = @()
                    Errors  = @()
                }
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
                $task.OlderThan.Unit,
                $task.OlderThan.Quantity
            }

            $M = "Start job on '{0}' with SourceFolderPath '{1}' DestinationFolderPath '{2}' DestinationFolderStructure '{3}' OlderThan.Unit '{4}' OlderThan.Quantity '{5}'" -f $env:COMPUTERNAME,
            $invokeParams.ArgumentList[0], $invokeParams.ArgumentList[1],
            $invokeParams.ArgumentList[2], $invokeParams.ArgumentList[3],
            $invokeParams.ArgumentList[4]
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            #region Start job
            $computerName = $task.ComputerName

            $task.Job.Object = if (
                $computerName -eq $ENV:COMPUTERNAME
            ) {
                Start-Job @invokeParams
            }
            else {
                $invokeParams += @{
                    ConfigurationName = $PSSessionConfiguration
                    ComputerName      = $computerName
                    AsJob             = $true
                }
                Invoke-Command @invokeParams
            }
            #endregion

            #region Wait for max running jobs
            $waitJobParams = @{
                Job        = $Tasks.Job.Object | Where-Object { $_ }
                MaxThreads = $MaxConcurrentJobs
            }
            Wait-MaxRunningJobsHC @waitJobParams
            #endregion
        }
        #endregion

        #region Wait for all jobs to finish
        Write-Verbose 'Wait for all jobs to finish'

        $null = $Tasks.Job.Object | Wait-Job
        #endregion

        #region Get job results and job errors
        foreach ($task in $Tasks) {
            $jobErrors = @()
            $receiveParams = @{
                ErrorVariable = 'jobErrors'
                ErrorAction   = 'SilentlyContinue'
            }
            $task.Job.Results += $task.Job.Object | Receive-Job @receiveParams

            foreach ($e in $jobErrors) {
                $task.Job.Errors += $e.ToString()
                $Error.Remove($e)

                $M = "Task error on '{0}' with SourceFolderPath '{1}' DestinationFolderPath '{2}' DestinationFolderStructure '{3}' OlderThan.Unit '{4}' OlderThan.Quantity '{5}': {6}" -f
                $task.ComputerName, $task.SourceFolderPath,
                $task.DestinationFolderPath, $task.DestinationFolderStructure,
                $task.OlderThan.Unit, $task.OlderThan.Quantity, $e.ToString()
                Write-Verbose $M; Write-EventLog @EventErrorParams -Message $M
            }
        }
        #endregion

        #region Export job results to Excel file
        if ($jobResults = $Tasks.Job.Results | Where-Object { $_ }) {
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
            jobErrors       = ($Tasks.job.Errors | Measure-Object).Count
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
        #region System errors HTML list
        $systemErrorsHtmlList = if ($counter.systemErrors) {
            "<p>Detected <b>{0} non terminating error{1}</b>:{2}</p>" -f $counter.systemErrors,
            $(
                if ($counter.systemErrors -gt 1) { 's' }
            ),
            $(
                $Error.Exception.Message | Where-Object { $_ } |
                ConvertTo-HtmlListHC
            )
        }
        #endregion

        #region Job results HTML list
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
                if ($task.OlderThan.Quantity -eq 0) {
                    'Move all files regardless their creation date'
                }
                else {
                    'Move files older than {0} {1}{2}' -f
                    $task.OlderThan.Quantity,
                    $(
                        $task.OlderThan.Unit.ToLower()
                    ),
                    $(
                        if ($task.OlderThan.Quantity -gt 1) { 's' }
                    )
                }
            ),
            $(
                (
                    $task.Job.Results |
                    Where-Object { $_.Action -like 'File moved*' } |
                    Measure-Object
                ).Count
            ),
            $(
                if ($errorCount = (
                        $task.Job.Results |
                        Where-Object { $_.Error } |
                        Measure-Object
                    ).Count + $task.Job.Errors.Count) {
                    ', <b style="color:red;">errors: {0}</b>' -f $errorCount
                }
            ),
            $(
                if ($task.Job.Errors) {
                    $task.Job.Errors | ForEach-Object {
                        '<br><b style="color:red;">{0}</b>' -f $_
                    }
                }
            )
        }

        $jobResultsHtmlList = $jobResultsHtmlListItems |
        ConvertTo-HtmlListHC -Spacing Wide
        #endregion

        #region Job errors HTML list
        $jobErrorsHtmlList = if ($counter.jobErrors) {
            $errorList = foreach (
                $task in
                $Tasks | Where-Object { $_.Job.Errors }
            ) {
                foreach ($e in $task.Job.Errors) {
                    "Failed task with ComputerName '{0}' with SourceFolderPath '{1}' DestinationFolderPath '{2}' DestinationFolderStructure '{3}' OlderThan.Unit '{4}' OlderThan.Quantity '{5}': {6}" -f
                    $task.ComputerName, $task.SourceFolderPath,
                    $task.DestinationFolderPath,
                    $task.DestinationFolderStructure,
                    $task.OlderThan.Unit, $task.OlderThan.Quantity, $e
                }
            }

            $errorList |
            ConvertTo-HtmlListHC -Spacing Wide -Header 'Job errors:'
        }
        #endregion
        #endregion

        $mailParams += @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Message   = "
                $systemErrorsHtmlList
                $jobErrorsHtmlList
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