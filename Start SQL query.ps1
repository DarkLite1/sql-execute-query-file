#Requires -Version 7
#Requires -Modules SqlServer

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