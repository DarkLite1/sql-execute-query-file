#Requires -Version 7
#Requires -Modules SqlServer, ImportExcel
#Requires -Modules Toolbox.HTML, Toolbox.EventLog

<#
.SYNOPSIS
    Execute SQL queries

.DESCRIPTION
    This script reads a .JSON input file that contains all the parameters
    required to execute the .SQL files.

    The files are executed in order and when one file fails the next ones are
    simply not executed and marked as 'Executed = false'. Success and the reason
    for failure are reported in an Excel file that is sent by e-mail.

.PARAMETER MaxConcurrentTasks
    How many tasks are allowed to run at the same time.

.PARAMETER Tasks
    Collection of tasks to executed.

.PARAMETER Tasks.ComputerName
    Computer name where the SQL database is hosted.

.PARAMETER Tasks.DatabaseName
    Name of the database located on the server instance. In case multiple
    databases need to be addressed use 'MASTER' and the 'USE database x'
    statement within the .SQL file(s).

.PARAMETER Tasks.SqlFiles
    The .SQL file containing the SQL statements to execute.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$SqlScript = "$PSScriptRoot\Start SQL query.ps1",
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Application specific\SQL\$ScriptName",
    [String[]]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        $error.Clear()

        Get-Job | Remove-Job -Force -EA Ignore

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

        #region Test script path exists
        try {
            $params = @{
                Path        = $SqlScript
                ErrorAction = 'Stop'
            }
            $sqlScriptPath = (Get-Item @params).FullName
        }
        catch {
            throw "SQL script with path '$($SqlScript)' not found"
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
                'MaxConcurrentJobs', 'Tasks', 'MailTo'
            ).where(
                { -not $file.$_ }
            ).foreach(
                { throw "Property '$_' not found" }
            )

            try {
                $null = [int]$MaxConcurrentJobs = $file.MaxConcurrentTasks
            }
            catch {
                throw "Property 'MaxConcurrentTasks' needs to be a number, the value '$($file.MaxConcurrentTasks)' is not supported."
            }

            $Tasks = $file.Tasks

            foreach ($task in $Tasks) {
                @(
                    'ComputerNames', 'DatabaseNames', 'SqlFiles'
                ).where(
                    { -not $task.$_ }
                ).foreach(
                    { throw "Property 'Tasks.$_' not found" }
                )

                foreach ($file in $task.SqlFiles) {
                    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
                        throw "SQL file '$file' not found."
                    }
                    if ($file -notMatch '.sql$') {
                        throw "SQL file '$file' needs to have extension '.sql'."
                    }
                }
            }
        }
        catch {
            throw "Input file '$ImportFile': $_"
        }
        #endregion

        #region Create a list of tasks to execute
        $tasksToExecute = foreach ($task in $Tasks) {
            $filesContent = foreach ($file in $task.SqlFiles) {
                $fileContent = Get-Content -LiteralPath $file -Raw -EA Stop

                if (-not $fileContent) {
                    throw "No file content in query file '$file'"
                }

                $fileContent
            }

            foreach ($computerName in $task.ComputerNames) {
                #region Set ComputerName
                if (
                    (-not $computerName) -or
                    ($computerName -eq 'localhost') -or
                    ($computerName -eq "$ENV:COMPUTERNAME.$env:USERDNSDOMAIN")
                ) {
                    $computerName = $env:COMPUTERNAME
                }
                #endregion

                foreach ($databaseName in $task.DatabaseNames) {
                    [PSCustomObject]@{
                        ComputerName = $computerName.ToUpper()
                        Database     = $databaseName.ToUpper()
                        SqlFile      = @{
                            Paths    = @($task.SqlFiles)
                            Contents = @($filesContent)
                        }
                        Job          = @{
                            Results = @()
                            Errors  = @()
                        }

                    }
                }
            }
        }
        #endregion
    }
    catch {
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
                    $sqlScriptPath = $using:sqlScriptPath
                    $PSSessionConfiguration = $using:PSSessionConfiguration
                    $EventVerboseParams = $using:EventVerboseParams
                    $EventErrorParams = $using:EventErrorParams
                    $EventOutParams = $using:EventOutParams
                }
                #endregion

                #region Create job parameters
                $invokeParams = @{
                    ComputerName        = $env:COMPUTERNAME
                    FilePath            = $sqlScriptPath
                    ArgumentList        = $task.ComputerName, $task.Database,
                    $task.SqlFile.Contents, $task.SqlFile.Paths
                    EnableNetworkAccess = $true
                    ErrorAction         = 'Stop'
                }

                $M = "Execute on '{0}\{1}' {2} .SQL files" -f
                $invokeParams.ArgumentList[0],
                $invokeParams.ArgumentList[1],
                $invokeParams.ArgumentList[3].Count
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
                #endregion

                #region Start job
                $task.Job.Results += Invoke-Command @invokeParams
                #endregion

                #region Verbose
                $M = "Results on '{0}\{1}' for {2} .SQL files: {3}" -f
                $invokeParams.ArgumentList[0],
                $invokeParams.ArgumentList[1],
                $invokeParams.ArgumentList[3].Count,
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
                $M = "Error on '{0}\{1}' {2} .SQL files: $_" -f
                $invokeParams.ArgumentList[0],
                $invokeParams.ArgumentList[1],
                $invokeParams.ArgumentList[3].Count
                Write-Warning $M; Write-EventLog @EventErrorParams -Message $M

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

        $tasksToExecute | ForEach-Object @foreachParams
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
        $mailParams = @{}

        $excelParams = @{
            Path               = $logFile + ' - Log.xlsx'
            NoNumberConversion = '*'
            AutoSize           = $true
            FreezeTopRow       = $true
        }

        #region Export job results to Excel file
        if ($jobResults = $tasksToExecute.Job.Results | Where-Object { $_ }) {
            $excelParams.WorksheetName = $excelParams.TableName = 'Overview'

            $M = "Export $($jobResults.Count) rows to Excel sheet '$($excelParams.WorksheetName)'"
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $jobResults |
            Select-Object -Property 'ComputerName',
            'DatabaseName', 'StartTime', 'EndTime',
            @{
                Name       = 'Duration'
                Expression = {
                    '{0:hh}:{0:mm}:{0:ss}:{0:fff}' -f
                    (New-TimeSpan -Start $_.StartTime -End $_.EndTime)
                }
            },
            'SqlFile', 'Output', 'Error' |
            Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Export job errors to Excel file
        if ($jobErrors = $tasksToExecute | Where-Object { $_.Job.Errors }) {
            $excelParams.WorksheetName = $excelParams.TableName = 'Errors'

            $M = "Export $($jobErrors.Count) rows to Excel sheet '$($excelParams.WorksheetName)'"
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

            $jobErrors |
            Select-Object -Property 'ComputerName', 'DatabaseName',
            @{
                Name       = 'SqlFiles';
                Expression = { $_.SqlFile.Paths.Count }
            },
            @{
                Name       = 'Error';
                Expression = { $_.Job.Errors -join ', ' }
            } |
            Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Count results, errors, ...
        $counter = @{
            sqlFiles         = $tasksToExecute.SqlFile.Paths.Count
            sqlFilesExecuted = (
                $jobResults | Where-Object { $_.Executed } | Measure-Object
            ).Count
            executionErrors  = (
                $jobResults | Where-Object { $_.Error } | Measure-Object
            ).Count
            jobErrors        = (
                $jobErrors | Measure-Object
            ).Count
            systemErrors     = (
                $Error.Exception.Message | Measure-Object
            ).Count
        }
        #endregion

        #region Mail subject and priority
        $mailParams.Priority = 'Normal'

        $mailParams.Subject = '{0} {1}' -f $counter.sqlFiles, $(
            if ($counter.sqlFiles -ne 1) { 'queries' } else { 'query' }
        )

        if (
            $totalErrorCount = $counter.executionErrors + $counter.jobErrors +
            $counter.systemErrors
        ) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $totalErrorCount error{0}" -f $(
                if ($totalErrorCount -ne 1) { 's' }
            )
        }
        #endregion

        #region Create error html lists
        $systemErrorsHtmlList = if ($counter.systemErrors) {
            "<p>Detected <b>{0} non terminating error{1}:{2}</p>" -f
            $counter.systemErrors,
            $(
                if ($counter.systemErrors -ne 1) { 's' }
            ),
            $(
                $Error.Exception.Message | Where-Object { $_ } |
                ConvertTo-HtmlListHC
            )
        }
        #endregion

        #region Send mail to user
        $summaryTable = "
        <table>
            <tr>
                <th>Total queries</th>
                <td>$($counter.sqlFiles)</td>
            </tr>
            <tr>
                <th>Executed queries</th>
                <td>$($counter.sqlFilesExecuted)</td>
            </tr>
            <tr>
                <th>Not executed queries</th>
                <td>$($counter.sqlFiles - $counter.sqlFilesExecuted)</td>
            </tr>
            <tr{0}>
                <th>Failed queries</th>
                <td>$($counter.executionErrors)</td>
            </tr>
            {1}
        </table>
        " -f $(
            if ($counter.executionErrors) {
                ' style="background-color: red"'
            }
        ),
        $(
            if ($counter.JobErrors) {
                "<tr style=`"background-color: red`">
                    <th>Job errors</th>
                    <td>$($counter.JobErrors)</td>
                </tr>"
            }
        )

        $mailParams += @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Message   = "
                $systemErrorsHtmlList
                <p>Summary:</p>
                $summaryTable"
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