#Requires -Version 5.1
#Requires -Modules SqlServer, ImportExcel
#Requires -Modules Toolbox.HTML, Toolbox.Remoting, Toolbox.EventLog

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

.PARAMETER Tasks.QueryFile
    The .SQL file containing the SQL statements to execute.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Application specific\SQL\$ScriptName",
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Function Get-JobDurationHC {
        [OutputType([TimeSpan])]
        Param (
            [Parameter(Mandatory)]
            [System.Management.Automation.Job]$Job,
            [Parameter(Mandatory)]
            [String]$ComputerName
        )

        $params = @{
            Start = $Job.PSBeginTime
            End   = $Job.PSEndTime
        }
        $jobDuration = New-TimeSpan @params

        $M = "'{0}' job duration '{1:hh}:{1:mm}:{1:ss}'" -f 
        $ComputerName, $jobDuration
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
      
        $jobDuration
    }
    Function Get-JobResultsAndErrorsHC {
        [OutputType([PSCustomObject])]
        Param (
            [Parameter(Mandatory)]
            [System.Management.Automation.Job]$Job,
            [Parameter(Mandatory)]
            [String]$ComputerName
        )

        $result = [PSCustomObject]@{
            Result = $null
            Errors = @()
        }

        #region Get job results
        $M = "'{0}' get job results" -f $ComputerName
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
              
        $jobErrors = @()
        $receiveParams = @{
            ErrorVariable = 'jobErrors'
            ErrorAction   = 'SilentlyContinue'
        }
        $result.Result = $Job | Receive-Job @receiveParams
        #endregion
   
        #region Get job errors
        $jobErrorsFound = $false

        foreach ($e in $jobErrors) {
            $jobErrorsFound = $true

            $M = "'{0}' job error '{1}'" -f $ComputerName, $e.ToString()
            Write-Warning $M; Write-EventLog @EventWarnParams -Message $M
                  
            $result.Errors += $M
            $error.Remove($e)
        }
        foreach ($e in $result.Result.Error | Where-Object { $_ }) {
            $jobErrorsFound = $true

            $M = "'{0}' error '{1}'" -f $ComputerName, $e
            Write-Warning $M; Write-EventLog @EventWarnParams -Message $M
        }
        #endregion

        if (-not $jobErrorsFound) {
            $M = "'{0}' job successful" -f $ComputerName
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        }

        $result
    }
    $executeQueryFiles = {
        Param (
            [Parameter(Mandatory)]
            [String]$ServerInstance,
            [Parameter(Mandatory)]
            [String]$Database,
            [Parameter(Mandatory)]
            [String[]]$Queries,
            [Parameter(Mandatory)]
            [String[]]$QueryFiles
        )

        $i = 0
        foreach ($query in $Queries) {
            try {
                $result = [PSCustomObject]@{
                    ComputerName = $ServerInstance
                    DatabaseName = $database
                    QueryFile    = $QueryFiles[$i]
                    Executed     = $false
                    Duration     = $null
                    Error        = $null
                    Output       = @()
                }

                if (-not $result.Error) {
                    $startDate = Get-Date

                    $params = @{
                        ServerInstance    = $ServerInstance
                        Database          = $database
                        Query             = $query
                        QueryTimeout      = '1000'
                        ConnectionTimeout = '20'
                        ErrorAction       = 'Stop'
                    }
                    $result.Output += Invoke-Sqlcmd @params
                    $result.Executed = $true

                    $duration = New-TimeSpan -Start $startDate -End (Get-Date)
                    $result.Duration = '{1:hh}:{1:mm}:{1:ss}' -f $duration
                }
            }
            catch {
                $duration = New-TimeSpan -Start $startDate -End (Get-Date)
                $result.Duration = '{1:hh}:{1:mm}:{1:ss}' -f $duration
                    
                $result.Error = $_
                $global:error.RemoveAt(0)
            }
            finally {
                $i++
                $result
            }
        }
    }
    $getJobResult = {
        #region Get job results
        $params = @{
            Job          = $completedTask.Job
            ComputerName = '{0}\{1}' -f $completedTask.ServerInstance, 
            $completedTask.Database
        }
        $jobOutput = Get-JobResultsAndErrorsHC @params

        $completedTask.JobDuration = Get-JobDurationHC @params
        #endregion

        #region Add job results
        $completedTask.JobResults += $jobOutput.Result

        $jobOutput.Errors | ForEach-Object { 
            $completedTask.JobErrors += $_ 
        }
        #endregion
            
        $completedTask.Job = $null
    }

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
        #region MailTo
        if (-not ($MailTo = $file.MailTo)) {
            throw "Input file '$ImportFile': Property 'MailTo' is missing"
        }
        #endregion

        #region MaxConcurrentTasks
        if ($file.PSObject.Properties.Name -notContains 'MaxConcurrentTasks') {
            throw "Input file '$ImportFile': Property 'MaxConcurrentTasks' not found."
        }
        if (-not ($file.MaxConcurrentTasks -is [int])) {
            throw "Input file '$ImportFile': Property 'MaxConcurrentTasks' needs to be a number, the value '$($file.MaxConcurrentTasks)' is not supported."
        }

        $MaxConcurrentJobs = [int]$file.MaxConcurrentTasks
        #endregion

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

            if (-not $task.QueryFile) {
                throw "Input file '$ImportFile': No 'QueryFile' found for the task on '$($task.ComputerName)'."
            }

            foreach ($q in $task.QueryFile) {
                if (-not (Test-Path -LiteralPath $q -PathType Leaf)) {
                    throw "Input file '$ImportFile': Query file '$q' not found for the task on '$($task.ComputerName)'."
                }
                if ($q -notMatch '.sql$') {
                    throw "Input file '$ImportFile': Query file '$q' is not supported, only the extension '.sql' is supported."
                }
            }
        }
        #endregion

        #region Create a list of tasks to execute
        $tasksToExecute = foreach ($task in $Tasks) {
            $queries = foreach ($queryFile in $task.QueryFile) {
                Get-Content -LiteralPath $queryFile -Raw -EA Stop
            }

            foreach ($computerName in $task.ComputerName) {
                foreach ($databaseName in $task.DatabaseName) {
                    [PSCustomObject]@{
                        ServerInstance = $computerName
                        Database       = $databaseName
                        QueryFiles     = @($task.QueryFile)
                        Queries        = @($queries)
                        Job            = $null
                        JobDuration    = $null
                        JobResults     = @()
                        JobErrors      = @()
                    }
                }
            }
        }
        #endregion

        $mailParams = @{}
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
        #region Start jobs to execute queries
        foreach ($task in $tasksToExecute) {
            $invokeParams = @{
                ScriptBlock  = $executeQueryFiles
                ArgumentList = $task.ServerInstance, $task.Database,
                $task.Queries, $task.QueryFiles
            }

            $M = "'{0}\{1}' Execute '{2}' .SQL files" -f 
            $invokeParams.ArgumentList[0], $invokeParams.ArgumentList[1],
            ($task.Queries).Count
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            
            $task.Job = Start-Job @invokeParams
            
            #region Wait for max running jobs
            $waitParams = @{
                Name       = $tasksToExecute.Job | Where-Object { $_ }
                MaxThreads = $MaxConcurrentJobs
            }
            Wait-MaxRunningJobsHC @waitParams
            #endregion

            #region Get job results
            foreach (
                $completedTask in 
                $tasksToExecute | Where-Object {
                    ($_.Job) -and
                    ($_.Job.State -match 'Completed|Failed')
                }
            ) {
                & $getJobResult
            }
            #endregion
        }
        #endregion

        #region Wait for jobs to finish and get results
        while (
            $runningTasks = $tasksToExecute | Where-Object { $_.Job }
        ) {
            #region Verbose progress
            $runningJobCounter = ($runningTasks | Measure-Object).Count
            if ($runningJobCounter -eq 1) {
                $M = 'Wait for the last running job to finish'
            }
            else {
                $M = "Wait for one of '{0}' running jobs to finish" -f $runningJobCounter
            }
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            #endregion

            $finishedJob = $runningTasks.Job | Wait-Job -Any
        
            $completedTask = $runningTasks | Where-Object {
                $_.Job.Id -eq $finishedJob.Id
            }
        
            & $getJobResult
        }
        #endregion

        #region Export job results to Excel file
        if ($jobResults = $tasksToExecute.JobResults | Where-Object { $_ }) {
            $M = "Export $($jobResults.Count) rows to Excel"
            Write-Verbose $M; Write-EventLog @EventOutParams -Message $M
            
            $excelParams = @{
                Path               = $logFile + ' - Log.xlsx'
                WorksheetName      = 'Overview'
                TableName          = 'Overview'
                NoNumberConversion = '*'
                AutoSize           = $true
                FreezeTopRow       = $true
            }
            $jobResults | 
            Select-Object -Property * -ExcludeProperty 'PSComputerName',
            'RunSpaceId', 'PSShowComputerName', 'Output' | 
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
            queriesTotal    = ($jobResults | Measure-Object).Count
            queriesExecuted = (
                $jobResults | Where-Object { $_.Executed } | Measure-Object
            ).Count
            queryErrors     = (
                $jobResults | Where-Object { $_.Error } | Measure-Object
            ).Count
            jobErrors       = (
                $tasksToExecute | Where-Object { $_.JobErrors } | Measure-Object
            ).Count
            systemErrors    = ($Error.Exception.Message | Measure-Object).Count
        }
        #endregion

        #region Mail subject and priority
        $mailParams.Priority = 'Normal'

        $mailParams.Subject = '{0} {1}' -f $counter.queriesTotal, $(
            if ($counter.queriesTotal -gt 1) { 'queries' } else { 'query' }
        )

        if (
            $totalErrorCount = $counter.queryErrors + $counter.jobErrors + 
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

        $summaryTable = "
        <table>
            <tr>
                <th>Total queries</th>
                <td>$($counter.queriesTotal)</td>
            </tr>
            <tr>
                <th>Executed queries</th>
                <td>$($counter.queriesExecuted)</td>
            </tr>
            <tr>
                <th>Not executed queries</th>
                <td>$($counter.queriesTotal - $counter.queriesExecuted)</td>
            </tr>
            <tr>
                <th>Failed queries</th>
                <td>$($counter.queryErrors)</td>
            </tr>
        </table>
        "
        
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