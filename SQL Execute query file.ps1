#Requires -Version 5.1
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

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        $file = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json
        #endregion

        #region Test .json file properties
        if (-not ($MailTo = $file.MailTo)) {
            throw "Input file '$ImportFile': Property 'MailTo' is missing"
        }

        if (-not ($MaxConcurrentJobs = $file.MaxConcurrentTasks)) {
            throw "Input file '$ImportFile': Property 'MaxConcurrentTasks' not found."
        }

        try {
            $null = [int]$file.MaxConcurrentTasks
        }
        catch {
            throw "Input file '$ImportFile': Property 'MaxConcurrentTasks' needs to be a number, the value '$($file.MaxConcurrentTasks)' is not supported."
        }

        #region Tasks
        if (-not ($Tasks = $file.Tasks)) {
            throw "Input file '$ImportFile': No 'Tasks' found."
        }

        foreach ($task in $Tasks) {
            if (-not $task.ComputerName) {
                throw "Input file '$ImportFile': No 'ComputerName' found in one of the 'Tasks'."
            }

            if (-not $task.DatabaseName) {
                throw "Input file '$ImportFile': No 'DatabaseName' found for the task on '$($task.ComputerName)'."
            }

            if (-not $task.SqlFiles) {
                throw "Input file '$ImportFile': No 'SqlFiles' found for the task on '$($task.ComputerName)'."
            }

            foreach ($q in $task.SqlFiles) {
                if (-not (Test-Path -LiteralPath $q -PathType Leaf)) {
                    throw "Input file '$ImportFile': Query file '$q' not found for the task on '$($task.ComputerName)'."
                }
                if ($q -notMatch '.sql$') {
                    throw "Input file '$ImportFile': Query file '$q' is not supported, only the extension '.sql' is supported."
                }
            }
        }
        #endregion
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

            foreach ($computerName in $task.ComputerName) {
                foreach ($databaseName in $task.DatabaseName) {
                    [PSCustomObject]@{
                        ComputerName = $computerName
                        Database     = $databaseName
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
        $sqlScriptBlock = {
            Param (
                [Parameter(Mandatory)]
                [String]$ComputerName,
                [Parameter(Mandatory)]
                [String]$DatabaseName,
                [Parameter(Mandatory)]
                [String[]]$SqlFileContents,
                [Parameter(Mandatory)]
                [String[]]$SqlFilePaths
            )

            $i = 0
            foreach ($fileContent in $SqlFileContents) {
                try {
                    $result = [PSCustomObject]@{
                        ComputerName = $ComputerName
                        DatabaseName = $DatabaseName
                        SqlFile      = $SqlFilePaths[$i]
                        StartTime    = Get-Date
                        EndTime      = $null
                        Executed     = $false
                        Error        = $null
                        Output       = @()
                    }

                    $result.StartTime = Get-Date

                    $params = @{
                        ServerInstance         = $ComputerName
                        Database               = $DatabaseName
                        Query                  = $fileContent
                        TrustServerCertificate = $true
                        QueryTimeout           = '1000'
                        ConnectionTimeout      = '20'
                        ErrorAction            = 'Stop'
                    }
                    $result.Output += Invoke-Sqlcmd @params
                    $result.Executed = $true
                }
                catch {
                    $result.Error = $_
                    $error.RemoveAt(0)
                }
                finally {
                    $result.EndTime = Get-Date

                    $i++
                    $result
                }
            }
        }

        $scriptBlock = {
            try {
                $task = $_

                #region Declare variables for code running in parallel
                if (-not $MaxConcurrentJobs) {
                    $sqlScriptBlock = $using:sqlScriptBlock
                    $PSSessionConfiguration = $using:PSSessionConfiguration
                    $EventVerboseParams = $using:EventVerboseParams
                }
                #endregion

                #region Create job parameters
                $invokeParams = @{
                    ScriptBlock  = $sqlScriptBlock
                    ArgumentList = $task.ComputerName, $task.Database,
                    $task.SqlFile.Contents, $task.SqlFile.Paths
                    ErrorAction  = 'Stop'
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
                $M = "Results on '{0}\{1}' for {2} .SQL files. Results: {3}" -f
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

        $mailParams.Subject = '{0} job{1}' -f $counter.sqlFiles, $(
            if ($counter.sqlFiles -ne 1) { 's' }
        )

        if (
            $totalErrorCount = $counter.executionErrors + $counter.jobErrors +
            $counter.systemErrors
        ) {
            $mailParams.Priority = 'High'
            $mailParams.Subject += ", $totalErrorCount error{0}" -f $(
                if ($totalErrorCount -gt 1) { 's' }
            )
        }
        #endregion

        #region Create error html lists
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
            <tr>
                <th>Failed queries</th>
                <td>$($counter.executionErrors)</td>
            </tr>
            {0}
        </table>
        " -f $(
            if ($counter.JobErrors) {
                "<tr>
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