#Requires -Version 7
#Requires -Modules Pester, SqlServer

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ComputerName    = 'PC1'
        DatabaseName    = 'db1'
        SqlFileContents = @('c1', 'c2')
        SqlFilePaths    = @('p1', 'p2')
    }

    Mock Invoke-Sqlcmd
}

Describe 'call Invoke-Sqlcmd' {
    BeforeAll {
        Mock Invoke-Sqlcmd {
            'r1'
        } -ParameterFilter {
            $Query -eq $testParams.SqlFileContents[0]
        }
        Mock Invoke-Sqlcmd {
            'r2'
        } -ParameterFilter {
            $Query -eq $testParams.SqlFileContents[1]
        }

        $actual = .$testScript @testParams

        $expected = @(
            [PSCustomObject]@{
                ComputerName = $testParams.ComputerName
                DatabaseName = $testParams.DatabaseName
                SqlFile      = $testParams.SqlFilePaths[0]
                StartTime    = Get-Date
                EndTime      = Get-Date
                Executed     = $true
                Error        = $null
                Output       = 'r1'
            }
            [PSCustomObject]@{
                ComputerName = $testParams.ComputerName
                DatabaseName = $testParams.DatabaseName
                SqlFile      = $testParams.SqlFilePaths[1]
                StartTime    = Get-Date
                EndTime      = Get-Date
                Executed     = $true
                Error        = $null
                Output       = 'r2'
            }
        )
    }
    It 'with the correct arguments' {
        foreach ($testQuery in $testParams.SqlFileContents) {
            Should -Invoke Invoke-Sqlcmd -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($ServerInstance -eq $testParams.ComputerName) -and
                ($Database -eq $testParams.DatabaseName) -and
                ($Query -eq $testQuery) -and
                ($TrustServerCertificate) -and
                ($QueryTimeout -eq '1000') -and
                ($ConnectionTimeout -eq '20') -and
                ($ErrorAction -eq 'Stop')
            }
        }
    }
    It 'return results' {
        $actual | Should -HaveCount $expected.Count

        foreach ($expectedResult in $expected) {
            $actualResult = $actual | Where-Object {
                ($_.SqlFile -eq $expectedResult.SqlFile)
            }

            $actualResult.ComputerName | Should -Be $expectedResult.ComputerName
            $actualResult.DatabaseName | Should -Be $expectedResult.DatabaseName
            $actualResult.Executed | Should -Be $expectedResult.Executed
            $actualResult.Error | Should -Be $expectedResult.Error
            $actualResult.Output | Should -Be $expectedResult.Output
            $actualResult.StartTime.ToString('yyyy') |
            Should -Be $expectedResult.StartTime.ToString('yyyy')
            $actualResult.EndTime.ToString('yyyy') |
            Should -Be $expectedResult.EndTime.ToString('yyyy')
        }
    }
}