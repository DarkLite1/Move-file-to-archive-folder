#Requires -Version 7
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

    .PARAMETER SourceFolder
        Path of the source folder where the file are located.

    .PARAMETER Destination.Folder
        Path of the destination folder where the files need be moved too.

    .PARAMETER Destination.ChildFolder
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
    [String]$MoveScript = "$PSScriptRoot\Move file.ps1",
    [String]$PSSessionConfiguration = 'PowerShell.7',
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\File or folder\Move file to archive folder\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
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

        #region Test script path exists
        try {
            $params = @{
                Path        = $MoveScript
                ErrorAction = 'Stop'
            }
            $moveScriptPath = (Get-Item @params).FullName
        }
        catch {
            throw "Move script with path '$($MoveScript)' not found"
        }
        #endregion

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        $file = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json
        #endregion

        #region Test .json file properties
        try {
            @(
                'SendMail', 'MaxConcurrentJobs', 'Tasks'
            ).where(
                { -not $file.$_ }
            ).foreach(
                { throw "Property '$_' not found" }
            )

            @(
                'To', 'When'
            ).where(
                { -not $file.SendMail.$_ }
            ).foreach(
                { throw "Property 'SendMail.$_' not found" }
            )

            $MaxConcurrentJobs = $file.MaxConcurrentJobs
            try {
                $null = $MaxConcurrentJobs.ToInt16($null)
            }
            catch {
                throw "Property 'MaxConcurrentJobs' needs to be a number, the value '$MaxConcurrentJobs' is not supported."
            }

            $Tasks = $file.Tasks
            foreach ($task in $Tasks) {
                @(
                    'SourceFolder', 'Destination', 'OlderThan', 'Option'
                ).where(
                    { -not $task.$_ }
                ).foreach(
                    { throw "Property 'Tasks.$_' not found" }
                )

                @(
                    'Folder', 'ChildFolder'
                ).where(
                    { -not $task.Destination.$_ }
                ).foreach(
                    { throw "Property 'Tasks.Destination.$_' not found" }
                )

                if ($task.Destination.ChildFolder -notMatch '^Year-Month$|^Year\\Month$|^Year$|^YYYYMM$') {
                    throw "Input file '$ImportFile': Value '$($task.Destination.ChildFolder)' is not supported by 'Destination.ChildFolder'. Valid options are 'Year-Month', 'Year\Month', 'Year' or 'YYYYMM'."
                }

                #region OlderThan
                if (-not $task.OlderThan.Unit) {
                    throw "Input file '$ImportFile': No 'OlderThan.Unit' found in one of the 'Tasks'."
                }

                if ($task.OlderThan.Unit -notMatch '^Day$|^Month$|^Year$') {
                    throw "Input file '$ImportFile': Value '$($task.OlderThan.Unit)' is not supported by 'OlderThan.Unit'. Valid options are 'Day', 'Month' or 'Year'."
                }

                if ($task.PSObject.Properties.Name -notContains 'OlderThan') {
                    throw "Input file '$ImportFile' SourceFolder '$($task.SourceFolder)': Property 'OlderThan' with 'Quantity' and 'Unit' not found."
                }

                if ($task.OlderThan.PSObject.Properties.Name -notContains 'Quantity') {
                    throw "Input file '$ImportFile' SourceFolder '$($task.SourceFolder)': Property 'OlderThan.Quantity' not found. Use value number '0' to move all files."
                }

                try {
                    $null = [int]$task.OlderThan.Quantity
                }
                catch {
                    throw "Input file '$ImportFile' SourceFolder '$($task.SourceFolder)': Property 'OlderThan.Quantity' needs to be a number, the value '$($task.OlderThan.Quantity)' is not supported. Use value number '0' to move all files."
                }
                #endregion

                #region Option
                if ($task.PSObject.Properties.Name -notContains 'Option') {
                    throw "Input file '$ImportFile' SourceFolder '$($task.SourceFolder)': Property 'Option' not found."
                }

                if ($task.Option.PSObject.Properties.Name -notContains 'DuplicateFile') {
                    throw "Input file '$ImportFile' SourceFolder '$($task.SourceFolder)': Property 'Option.DuplicateFile' not found."
                }

                if (
                    ($task.Option.DuplicateFile) -and
                    ($task.Option.DuplicateFile -notMatch '^OverwriteFile$|^RenameFile$')
                ) {
                    throw "Input file '$ImportFile': Value '$($task.Option.DuplicateFile)' is not supported by 'Option.DuplicateFile'. Valid options are NULL, 'OverwriteFile' or 'RenameFile'."
                }
                #endregion
            }
        }
        catch {
            throw "Input file '$ImportFile': $_"
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
                    Results = @()
                    Errors  = @()
                }
            }
            #endregion
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

Process {
    Try {
        $scriptBlock = {
            try {
                $task = $_

                #region Declare variables for code running in parallel
                if (-not $MaxConcurrentJobs) {
                    $moveScriptPath = $using:moveScriptPath
                    $PSSessionConfiguration = $using:PSSessionConfiguration
                    $EventVerboseParams = $using:EventVerboseParams
                    $EventErrorParams = $using:EventErrorParams
                    $EventOutParams = $using:EventOutParams
                }
                #endregion

                #region Create job parameters
                $invokeParams = @{
                    FilePath     = $moveScriptPath
                    ArgumentList = $task.SourceFolder,
                    $task.Destination.Folder,
                    $task.Destination.ChildFolder,
                    $task.OlderThan.Unit,
                    $task.OlderThan.Quantity,
                    $task.Option.DuplicateFile
                }

                $M = "Start job on '{0}' with SourceFolder '{1}' Destination.Folder '{2}' Destination.ChildFolder '{3}' OlderThan.Unit '{4}' OlderThan.Quantity '{5}' Option.DuplicateFile '{6}'" -f $env:COMPUTERNAME,
                $invokeParams.ArgumentList[0], $invokeParams.ArgumentList[1],
                $invokeParams.ArgumentList[2], $invokeParams.ArgumentList[3],
                $invokeParams.ArgumentList[4], $invokeParams.ArgumentList[5]
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
                #endregion

                #region Start job
                $computerName = $task.ComputerName

                $task.Job.Results += if (
                    $computerName -eq $ENV:COMPUTERNAME
                ) {
                    $params = $invokeParams.ArgumentList
                    & $invokeParams.FilePath @params
                }
                else {
                    $invokeParams += @{
                        ConfigurationName = $PSSessionConfiguration
                        ComputerName      = $computerName
                        ErrorAction       = 'Stop'
                    }
                    Invoke-Command @invokeParams
                }
                #endregion

                #region Verbose
                $M = "Task on '{0}' with SourceFolder '{1}' Destination.Folder '{2}' Destination.ChildFolder '{3}' OlderThan.Unit '{4}' OlderThan.Quantity '{5}' Option.DuplicateFile '{6}'. Results: {7}" -f $env:COMPUTERNAME,
                $invokeParams.ArgumentList[0], $invokeParams.ArgumentList[1],
                $invokeParams.ArgumentList[2], $invokeParams.ArgumentList[3],
                $invokeParams.ArgumentList[4], $invokeParams.ArgumentList[5],
                $task.Job.Results.Count

                if ($errorCount = $task.Job.Results.Where({ $_.Error }).Count) {
                    $M += " , Errors: {0}" -f $errorCount
                    Write-Warning $M
                    Write-EventLog @EventErrorParams -Message $M
                }
                elseif ($task.Job.Results.Count) {
                    Write-Verbose $M
                    Write-EventLog @EventOutParams -Message $M
                }
                else {
                    Write-Verbose $M
                    Write-EventLog @EventVerboseParams -Message $M
                }
                #endregion
            }
            catch {
                $task.Job.Errors += $_
                $Error.RemoveAt(0)
            }
        }

        #region Run code serial or parallel
        $foreachParams = if ($MaxConcurrentJobs -eq 1) {
            @{
                Process = $scriptBlock
            }
        }
        else {
            @{
                Parallel      = $scriptBlock
                ThrottleLimit = $MaxConcurrentJobs
            }
        }

        $Tasks | ForEach-Object @foreachParams
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
        $mailParams = @{ }

        #region Export job results to Excel file
        $excelWorksheet = @{
            Overview = @()
            Errors   = @()
        }

        foreach ($task in $Tasks) {
            if ($task.Job.Errors) {
                $excelWorksheet.Errors += $task.Job.Errors |
                Select-Object -Property @{
                    Name       = 'ComputerName'
                    Expression = { $task.ComputerName }
                },
                @{
                    Name       = 'SourceFolder'
                    Expression = { $task.SourceFolder }
                },
                @{
                    Name       = 'Error'
                    Expression = { $_ }
                }
            }

            if ($task.Job.Results) {
                $excelWorksheet.Overview += $task.Job.Results |
                Select-Object -Property @{
                    Name       = 'ComputerName'
                    Expression = { $task.ComputerName }
                },
                @{
                    Name       = 'OlderThan'
                    Expression = {
                        '{0} {1}{2}' -f
                        $task.OlderThan.Quantity,
                        $task.OlderThan.Unit,
                        $(if ($task.OlderThan.Quantity -gt 1) { 's' })
                    }
                },
                @{
                    Name       = 'SourceFolder'
                    Expression = { $task.SourceFolder }
                },
                @{
                    Name       = 'DestinationFolder'
                    Expression = { $_.DestinationFolderPath }
                },
                @{
                    Name       = 'FileName'
                    Expression = { $_.FileName }
                },
                @{
                    Name       = 'FileCreationTime'
                    Expression = { $_.FileCreationTime }
                },
                @{
                    Name       = 'Action'
                    Expression = { $_.Action }
                },
                @{
                    Name       = 'Error'
                    Expression = { $_.Error }
                }
            }
        }

        $excelParams = @{
            Path               = $logFile + '- Log.xlsx'
            NoNumberConversion = '*'
            AutoSize           = $true
            FreezeTopRow       = $true
        }

        if ($excelWorksheet.Overview) {
            $excelParams.TableName = $excelParams.WorksheetName = 'Overview'

            $M = "Export {0} rows to Excel sheet '{1}'" -f
            $excelWorksheet.Overview.Count, $excelParams.WorksheetName
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $excelWorksheet.Overview | Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }

        if ($excelWorksheet.Errors) {
            $excelParams.TableName = $excelParams.WorksheetName = 'Errors'

            $M = "Export {0} rows to Excel sheet '{1}'" -f
            $excelWorksheet.Overview.Count, $excelParams.WorksheetName
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $excelWorksheet.Errors | Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Count results & errors
        $counter = @{
            movedFiles     = $Tasks.Job.Results.where(
                { $_.Action -like 'File moved*' }).Count
            moveFileErrors = $Tasks.Job.Results.where({ $_.Error }).Count
            jobErrors      = $Tasks.Job.where({ $_.Error }).Count
            systemErrors   = ($Error.Exception.Message | Measure-Object).Count
        }

        $totalErrorCount = $counter.moveFileErrors + $counter.jobErrors +
        $counter.systemErrors
        #endregion

        #region Mail subject and priority
        $mailParams.Priority = 'Normal'

        $mailParams.Subject = '{0} file{1} moved' -f
        $counter.movedFiles, $(if ($counter.movedFiles -gt 1) { 's' })

        if ($totalErrorCount) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $totalErrorCount error{0}" -f
            $(if ($totalErrorCount -gt 1) { 's' })
        }
        #endregion

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
            $Tasks | Sort-Object -Property 'SourceFolder'
        ) {
            'From: {0}<br>To: {1}<br>{2}<br>Moved: {3}{4}{5}' -f
            $(
                if ($task.SourceFolder -match '^\\\\') {
                    '<a href="{0}">{0}</a>' -f $task.SourceFolder
                }
                else {
                    $uncPath = $task.SourceFolder -Replace '^.{2}', (
                        '\\{0}\{1}$' -f $task.ComputerName, $task.SourceFolder[0]
                    )
                    '<a href="{0}">{0}</a>' -f $uncPath
                }
            ),
            $(
                if ($task.Destination.Folder -match '^\\\\') {
                    '<a href="{0}">{0}</a>' -f $task.Destination.Folder
                }
                else {
                    $uncPath = $task.Destination.Folder -Replace '^.{2}', (
                        '\\{0}\{1}$' -f $task.ComputerName, $task.Destination.Folder[0]
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
        $jobErrorsHtml = if ($counter.jobErrors) {
            '<p>Detected <b>{0} job errors</b>.</p>' -f
            $counter.jobErrors
        }
        #endregion

        #region Check to send mail to user
        $sendMailToUser = $false

        if (
            (
                ($file.SendMail.When -eq 'Always')
            ) -or
            (
                ($file.SendMail.When -eq 'OnlyOnError') -and
                $totalErrorCount
            ) -or
            (
                ($file.SendMail.When -eq 'OnlyOnErrorOrAction') -and
                (
                    ($counter.movedFiles) -or $totalErrorCount
                )
            )
        ) {
            $sendMailToUser = $true
        }
        #endregion

        #region Send mail to user
        $mailParams += @{
            To             = $file.SendMail.To
            Bcc            = $ScriptAdmin
            Message        = "
                $systemErrorsHtmlList
                $jobErrorsHtml
                <p>Summary:</p>
                $jobResultsHtmlList"
            LogFolder      = $LogParams.LogFolder
            Header         = $ScriptName
            EventLogSource = $ScriptName
            Save           = $LogFile + ' - Mail.html'
            ErrorAction    = 'Stop'
        }

        if ($mailParams.Attachments) {
            $mailParams.Message +=
            "<p><i>* Check the attachment for details</i></p>"
        }

        Get-ScriptRuntimeHC -Stop

        if ($sendMailToUser) {
            $M = 'Send e-mail to the user'
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            if ($totalErrorCount) {
                $mailParams.Bcc = $ScriptAdmin
            }
            Send-MailHC @mailParams
        }
        else {
            $M = 'Send no e-mail to the user'
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            if ($totalErrorCount) {
                Write-Verbose 'Send e-mail to admin only with errors'

                $mailParams.To = $ScriptAdmin
                Send-MailHC @mailParams
            }
        }
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