#Requires -Version 5.1
#Requires -Modules ImportExcel, Toolbox.HTML, Toolbox.Remoting, Toolbox.EventLog

<#
.SYNOPSIS
    Execute SQL queries

.DESCRIPTION
    This script reads a .JSON input file that contains all the information
    required to run SQL queries.

.PARAMETER MaxConcurrentTasks
    How many tasks can be running at the same time.

.PARAMETER Tasks
    Collection of tasks to executed

.PARAMETER Tasks.ComputerName
    Computer name where the SQL database is hosted

.PARAMETER Tasks.DatabaseName
    Name of the database located on the server instance

.PARAMETER Tasks.Query
    The queries to execute against the databases on the server instances
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
    $scriptBlock = {
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
                    Error        = $null
                    Output       = @()
                }

                if (-not $result.Error) {
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
                }
            }
            catch {
                $result.Error = $_
                $global:error.RemoveAt(0)
            }
            finally {
                $i++
                $result
            }
        }
    }

    try {
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams
        Get-ScriptRuntimeHC -Start

        $error.Clear()

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
                if ($q -notMatch '.sql$|.txt$') {
                    throw "Input file '$ImportFile': Query file '$q' is not supported, only extensions '.txt' or '.sql' are supported."
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
                ScriptBlock  = $scriptBlock
                ArgumentList = $task.ServerInstance, $task.Database,
                $task.Queries, $task.QueryFiles
            }

            $M = "Start job for server instance '{0}' database '{1}' queries '{2}''" -f 
            $invokeParams.ArgumentList[0], $invokeParams.ArgumentList[1],
            ($task.Queries).Count
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            
            $task.Job = Start-Job @invokeParams
            
            $waitParams = @{
                Name       = $tasksToExecute.Job | Where-Object { $_ }
                MaxThreads = $MaxConcurrentJobs
            }
            Wait-MaxRunningJobsHC @waitParams
        }
        #endregion

        #region Wait for jobs to finish
        $M = "Wait for all $($tasksToExecute.count) jobs to finish"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        $null = $tasksToExecute.Job | Wait-Job
        #endregion

        #region Get job results and job errors
        foreach ($task in $tasksToExecute) {
            $jobErrors = @()
            $receiveParams = @{
                ErrorVariable = 'jobErrors'
                ErrorAction   = 'SilentlyContinue'
            }
            $task.JobResults += $task.Job | Receive-Job @receiveParams

            foreach ($e in $jobErrors) {
                $task.JobErrors += $e.ToString()
                $Error.Remove($e)

                $M = "Task error on '{0}\{1}': {2}" -f 
                $task.ServerInstance, $task.Database, $e.ToString()
                Write-Warning $M; Write-EventLog @EventErrorParams -Message $M
            }

            if (-not $jobErrors) {
                $M = 'No job errors'
                Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
            }
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

        $summaryHtmlList = "
        <table>
            <tr>
                <th>Total queries</th>
                <td>$($counter.queriesTotal)</td>
            <tr>
            <tr>
                <th>Executed queries</th>
                <td>$($counter.queriesExecuted)</td>
            <tr>
            <tr>
                <th>Not executed queries</th>
                <td>$($counter.queriesTotal - $counter.queriesExecuted)</td>
            <tr>
            <tr>
                <th>Failed queries</th>
                <td>$($counter.queryErrors)</td>
            <tr>
        </table>
        "
        
        $mailParams += @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Message   = "
                $systemErrorsHtmlList
                <p>Summary:</p>
                $summaryHtmlList"
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