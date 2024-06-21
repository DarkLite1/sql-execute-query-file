#Requires -Version 7
#Requires -Modules Pester, ImportExcel

BeforeAll {
    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/params.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testData = @{
        sqlFile = @(
            @{
                Path    = (New-Item 'TestDrive:/s1.sql' -ItemType File).FullName
                Content = '-- SQL instructions file 1'
            }
            @{
                Path    = (New-Item 'TestDrive:/s2.sql' -ItemType File).FullName
                Content = '-- SQL instructions file 2'
            }
        )
    }

    $testData.sqlFile.foreach(
        { $_.Content | Out-File -FilePath $_.Path -NoNewline }
    )

    $testInputFile = @{
        MailTo             = 'bob@contoso.com'
        MaxConcurrentTasks = 1
        Tasks              = @(
            @{
                ComputerNames = @('PC1')
                DatabaseNames = @('db1')
                SqlFiles      = $testData.sqlFile.Path
            }
        )
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        ImportFile = $testOutParams.FilePath
        SqlScript  = (New-Item 'TestDrive:/s.ps1' -ItemType File).FullName
        LogFolder  = New-Item 'TestDrive:/log' -ItemType Directory
    }

    Mock Invoke-Sqlcmd
    Mock Invoke-Command
    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @('ImportFile', 'ScriptName') {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and
            ($Subject -eq 'FAILURE')
        }
    }
    It 'the log folder cannot be created' {
        $testNewParams = Copy-ObjectHC $testParams
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and
            ($Message -like '*Failed creating the log folder*')
        }
    }
    It 'the file SqlScript cannot be found' {
        $testNewParams = Copy-ObjectHC $testParams
        $testNewParams.SqlScript = 'c:\upDoesNotExist.ps1'

        $testInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*SQL script with path '$($testNewParams.SqlScript)' not found*")
        }
        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
            $EntryType -eq 'Error'
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = Copy-ObjectHC $testParams
            $testNewParams.ImportFile = 'nonExisting.json'

            .$testScript @testNewParams

            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            It '<_> not found' -ForEach @(
                'MaxConcurrentTasks', 'Tasks', 'MailTo'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property '$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.<_> not found' -ForEach @(
                'ComputerNames', 'DatabaseNames', 'SqlFiles'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and
                    ($Message -like "*$ImportFile*Property 'Tasks.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context 'SqlFiles' {
                It 'file not found' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].SqlFiles = @('notExisting.sql')

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams


                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*SQL file 'notExisting.sql' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'file extension not .sql' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].SqlFiles = @((
                            New-Item 'TestDrive:/a.txt' -ItemType File).FullName
                    )

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*SQL file '$($testNewInputFile.Tasks[0].SqlFiles)' needs to have extension '.sql'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'file empty' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].SqlFiles = @((
                            New-Item 'TestDrive:/b.sql' -ItemType File).FullName
                    )

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*No file content in SQL file '$($testNewInputFile.Tasks[0].SqlFiles[0])'*")
                    }
                }
            }
            Context 'MaxConcurrentTasks' {
                It 'is not a number' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.MaxConcurrentTasks = 'a'

                    $testNewInputFile | ConvertTo-Json -Depth 7 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'MaxConcurrentTasks' needs to be a number, the value 'a' is not supported*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
        }
    }
}
Describe 'execute the SQL script with Invoke-Command' {
    It 'once for 1 computer name and 1 database name' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].ComputerNames = 'PC1'
        $testNewInputFile.Tasks[0].DatabaseNames = 'db1'

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Invoke-Command -Times 1 -Exactly -ParameterFilter {
            ($ComputerName -eq $env:COMPUTERNAME) -and
            ($FilePath -eq $testParams.SqlScript) -and
            ($EnableNetworkAccess) -and
            ($ErrorAction -eq 'Stop') -and
            ($ArgumentList[0] -eq $testNewInputFile.Tasks[0].ComputerNames) -and
            ($ArgumentList[1] -eq $testNewInputFile.Tasks[0].DatabaseNames) -and
            ($ArgumentList[2][0] -eq $testData.sqlFile[0].Content) -and
            ($ArgumentList[2][1] -eq $testData.sqlFile[1].Content) -and
            ($ArgumentList[3][0] -eq $testData.sqlFile[0].Path) -and
            ($ArgumentList[3][1] -eq $testData.sqlFile[1].Path)
        }

        Should -Invoke Invoke-Command -Times 1 -Exactly
    }
    It 'once for each computer name and each database name' {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].ComputerNames = @('PC1', 'PC2')
        $testNewInputFile.Tasks[0].DatabaseNames = @('db1', 'db2')

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams

        foreach (
            $testComputer in
            $testNewInputFile.Tasks[0].ComputerNames
        ) {
            foreach (
                $testDatabase in
                $testNewInputFile.Tasks[0].DatabaseNames
            ) {
                Should -Invoke Invoke-Command -Times 1 -Exactly -ParameterFilter {
                    ($ComputerName -eq $env:COMPUTERNAME) -and
                    ($FilePath -eq $testParams.SqlScript) -and
                    ($EnableNetworkAccess) -and
                    ($ErrorAction -eq 'Stop') -and
                    ($ArgumentList[0] -eq $testComputer) -and
                    ($ArgumentList[1] -eq $testDatabase) -and
                    ($ArgumentList[2][0] -eq $testData.sqlFile[0].Content) -and
                    ($ArgumentList[2][1] -eq $testData.sqlFile[1].Content) -and
                    ($ArgumentList[3][0] -eq $testData.sqlFile[0].Path) -and
                    ($ArgumentList[3][1] -eq $testData.sqlFile[1].Path)
                }
            }
        }

        Should -Invoke Invoke-Command -Times 4 -Exactly
    }
} -Tag test
Describe 'when a query is slow and MaxConcurrentTasks is 6' {
    BeforeAll {
        @{
            MailTo             = 'bob@contoso.com'
            MaxConcurrentTasks = 6
            Tasks              = @(
                @{
                    ComputerName = @('PC1', 'PC2')
                    DatabaseName = @('a', 'b')
                    SqlFiles     = $testData.sqlFile.Path
                }
            )
        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        $testExportedExcelRows = @(
            [PSCustomObject]@{
                ComputerName = 'PC1'
                DatabaseName = 'a'
                SqlFile      = 'c:\query1.sql'
                Executed     = $true
                Duration     = '00:00:00:1a1'
                Error        = 'problem'
            }
            [PSCustomObject]@{
                ComputerName = 'PC1'
                DatabaseName = 'a'
                SqlFile      = 'c:\query2.sql'
                Executed     = $false
                Duration     = $null
                Error        = $null
            }
            [PSCustomObject]@{
                ComputerName = 'PC1'
                DatabaseName = 'b'
                SqlFile      = 'c:\query1.sql'
                Executed     = $true
                Duration     = '00:00:00:1b1'
                Error        = $null
            }
            [PSCustomObject]@{
                ComputerName = 'PC1'
                DatabaseName = 'b'
                SqlFile      = 'c:\query2.sql'
                Executed     = $true
                Duration     = '00:00:00:1b2'
                Error        = $null
            }
            [PSCustomObject]@{
                ComputerName = 'PC2'
                DatabaseName = 'a'
                SqlFile      = 'c:\query1.sql'
                Executed     = $true
                Duration     = '00:00:00:2a1'
                Error        = $null
            }
            [PSCustomObject]@{
                ComputerName = 'PC2'
                DatabaseName = 'a'
                SqlFile      = 'c:\query2.sql'
                Executed     = $true
                Duration     = '00:00:00:2a2'
                Error        = $null
            }
            [PSCustomObject]@{
                ComputerName = 'PC2'
                DatabaseName = 'b'
                SqlFile      = 'c:\query1.sql'
                Executed     = $true
                Duration     = '00:00:00:2b1'
                Error        = $null
            }
            [PSCustomObject]@{
                ComputerName = 'PC2'
                DatabaseName = 'b'
                SqlFile      = 'c:\query2.sql'
                Executed     = $true
                Duration     = '00:00:00:2b2'
                Error        = $null
            }
        )

        Mock Start-Job -MockWith {
            & $realStartJobCommand -Scriptblock {
                Start-Sleep -Seconds 2
                @(
                    [PSCustomObject]@{
                        ComputerName = 'PC1'
                        DatabaseName = 'a'
                        SqlFile      = 'c:\query1.sql'
                        Executed     = $true
                        Duration     = '00:00:00:1a1'
                        Error        = 'problem'
                    }
                    [PSCustomObject]@{
                        ComputerName = 'PC1'
                        DatabaseName = 'a'
                        SqlFile      = 'c:\query2.sql'
                        Executed     = $false
                        Duration     = $null
                        Error        = $null
                    }
                )
            }
        } -ParameterFilter {
            ($ArgumentList[0] -eq 'PC1') -and
            ($ArgumentList[1] -eq 'a')
        }
        Mock Start-Job -MockWith {
            & $realStartJobCommand -Scriptblock {
                @(
                    [PSCustomObject]@{
                        ComputerName = 'PC2'
                        DatabaseName = 'a'
                        SqlFile      = 'c:\query1.sql'
                        Executed     = $true
                        Duration     = '00:00:00:2a1'
                        Error        = $null
                    }
                    [PSCustomObject]@{
                        ComputerName = 'PC2'
                        DatabaseName = 'a'
                        SqlFile      = 'c:\query2.sql'
                        Executed     = $true
                        Duration     = '00:00:00:2a2'
                        Error        = $null
                    }
                )
            }
        } -ParameterFilter {
            ($ArgumentList[0] -eq 'PC2') -and
            ($ArgumentList[1] -eq 'a')
        }
        Mock Start-Job -MockWith {
            & $realStartJobCommand -Scriptblock {
                @(
                    [PSCustomObject]@{
                        ComputerName = 'PC1'
                        DatabaseName = 'b'
                        SqlFile      = 'c:\query1.sql'
                        Executed     = $true
                        Duration     = '00:00:00:1b1'
                        Error        = $null
                    }
                    [PSCustomObject]@{
                        ComputerName = 'PC1'
                        DatabaseName = 'b'
                        SqlFile      = 'c:\query2.sql'
                        Executed     = $true
                        Duration     = '00:00:00:1b2'
                        Error        = $null
                    }
                )
            }
        } -ParameterFilter {
            ($ArgumentList[0] -eq 'PC1') -and
            ($ArgumentList[1] -eq 'b')
        }
        Mock Start-Job -MockWith {
            & $realStartJobCommand -Scriptblock {
                @(
                    [PSCustomObject]@{
                        ComputerName = 'PC2'
                        DatabaseName = 'b'
                        SqlFile      = 'c:\query1.sql'
                        Executed     = $true
                        Duration     = '00:00:00:2b1'
                        Error        = $null
                    }
                    [PSCustomObject]@{
                        ComputerName = 'PC2'
                        DatabaseName = 'b'
                        SqlFile      = 'c:\query2.sql'
                        Executed     = $true
                        Duration     = '00:00:00:2b2'
                        Error        = $null
                    }
                )
            }
        } -ParameterFilter {
            ($ArgumentList[0] -eq 'PC2') -and
            ($ArgumentList[1] -eq 'b')
        }

        .$testScript @testParams
    }
    It 'Start-Job is called for each ComputerName and each database' {
        Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($ScriptBlock) -and
            ($ArgumentList[0] -eq 'PC1') -and
            ($ArgumentList[1] -eq 'a') -and
            ($ArgumentList[2].Count -eq 2)
        }
        Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($ScriptBlock) -and
            ($ArgumentList[0] -eq 'PC1') -and
            ($ArgumentList[1] -eq 'b') -and
            ($ArgumentList[2].Count -eq 2)
        }
        Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($ScriptBlock) -and
            ($ArgumentList[0] -eq 'PC2') -and
            ($ArgumentList[1] -eq 'a') -and
            ($ArgumentList[2].Count -eq 2)
        }
        Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($ScriptBlock) -and
            ($ArgumentList[0] -eq 'PC2') -and
            ($ArgumentList[1] -eq 'b') -and
            ($ArgumentList[2].Count -eq 2)
        }
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    ($_.ComputerName -eq $testRow.ComputerName) -and
                    ($_.DatabaseName -eq $testRow.DatabaseName) -and
                    ($_.SqlFile -eq $testRow.SqlFile)
                }
                $actualRow.Executed | Should -Be $testRow.Executed
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.Duration | Should -Be $testRow.Duration
            }
        }
    }
    It 'send a summary mail to the user' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq 'bob@contoso.com') -and
            ($Bcc -eq $ScriptAdmin) -and
            ($Priority -eq 'High') -and
            ($Subject -eq '8 jobs, 1 error') -and
            ($Attachments.Count -eq 1) -and
            ($Attachments -like '* - Log.xlsx') -and
            ($Message -like "*<th>Total queries</th>*<td>8</td>*<th>Executed queries</th>*<td>7</td>*<th>Not executed queries</th>*<td>1</td>*<th>Failed queries</th>*<td>1</td>*<p><i>* Check the attachment for details</i></p>*")
        }
    }
}
Describe 'when all queries are fast and MaxConcurrentTasks is 1' {
    BeforeAll {
        @{
            MailTo             = 'bob@contoso.com'
            MaxConcurrentTasks = 1
            Tasks              = @(
                @{
                    ComputerName = @('PC1', 'PC2')
                    DatabaseName = @('a', 'b')
                    SqlFiles     = $testData.sqlFile.Path
                }
            )
        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        $testExportedExcelRows = @(
            [PSCustomObject]@{
                ComputerName = 'PC1'
                DatabaseName = 'a'
                SqlFile      = 'c:\query1.sql'
                Executed     = $true
                Duration     = '00:00:00:1a1'
                Error        = 'problem'
            }
            [PSCustomObject]@{
                ComputerName = 'PC1'
                DatabaseName = 'a'
                SqlFile      = 'c:\query2.sql'
                Executed     = $false
                Duration     = $null
                Error        = $null
            }
            [PSCustomObject]@{
                ComputerName = 'PC1'
                DatabaseName = 'b'
                SqlFile      = 'c:\query1.sql'
                Executed     = $true
                Duration     = '00:00:00:1b1'
                Error        = $null
            }
            [PSCustomObject]@{
                ComputerName = 'PC1'
                DatabaseName = 'b'
                SqlFile      = 'c:\query2.sql'
                Executed     = $true
                Duration     = '00:00:00:1b2'
                Error        = $null
            }
            [PSCustomObject]@{
                ComputerName = 'PC2'
                DatabaseName = 'a'
                SqlFile      = 'c:\query1.sql'
                Executed     = $true
                Duration     = '00:00:00:2a1'
                Error        = $null
            }
            [PSCustomObject]@{
                ComputerName = 'PC2'
                DatabaseName = 'a'
                SqlFile      = 'c:\query2.sql'
                Executed     = $true
                Duration     = '00:00:00:2a2'
                Error        = $null
            }
            [PSCustomObject]@{
                ComputerName = 'PC2'
                DatabaseName = 'b'
                SqlFile      = 'c:\query1.sql'
                Executed     = $true
                Duration     = '00:00:00:2b1'
                Error        = $null
            }
            [PSCustomObject]@{
                ComputerName = 'PC2'
                DatabaseName = 'b'
                SqlFile      = 'c:\query2.sql'
                Executed     = $true
                Duration     = '00:00:00:2b2'
                Error        = $null
            }
        )

        Mock Start-Job -MockWith {
            & $realStartJobCommand -Scriptblock {
                @(
                    [PSCustomObject]@{
                        ComputerName = 'PC1'
                        DatabaseName = 'a'
                        SqlFile      = 'c:\query1.sql'
                        Executed     = $true
                        Duration     = '00:00:00:1a1'
                        Error        = 'problem'
                    }
                    [PSCustomObject]@{
                        ComputerName = 'PC1'
                        DatabaseName = 'a'
                        SqlFile      = 'c:\query2.sql'
                        Executed     = $false
                        Duration     = $null
                        Error        = $null
                    }
                )
            }
        } -ParameterFilter {
            ($ArgumentList[0] -eq 'PC1') -and
            ($ArgumentList[1] -eq 'a')
        }
        Mock Start-Job -MockWith {
            & $realStartJobCommand -Scriptblock {
                @(
                    [PSCustomObject]@{
                        ComputerName = 'PC2'
                        DatabaseName = 'a'
                        SqlFile      = 'c:\query1.sql'
                        Executed     = $true
                        Duration     = '00:00:00:2a1'
                        Error        = $null
                    }
                    [PSCustomObject]@{
                        ComputerName = 'PC2'
                        DatabaseName = 'a'
                        SqlFile      = 'c:\query2.sql'
                        Executed     = $true
                        Duration     = '00:00:00:2a2'
                        Error        = $null
                    }
                )
            }
        } -ParameterFilter {
            ($ArgumentList[0] -eq 'PC2') -and
            ($ArgumentList[1] -eq 'a')
        }
        Mock Start-Job -MockWith {
            & $realStartJobCommand -Scriptblock {
                @(
                    [PSCustomObject]@{
                        ComputerName = 'PC1'
                        DatabaseName = 'b'
                        SqlFile      = 'c:\query1.sql'
                        Executed     = $true
                        Duration     = '00:00:00:1b1'
                        Error        = $null
                    }
                    [PSCustomObject]@{
                        ComputerName = 'PC1'
                        DatabaseName = 'b'
                        SqlFile      = 'c:\query2.sql'
                        Executed     = $true
                        Duration     = '00:00:00:1b2'
                        Error        = $null
                    }
                )
            }
        } -ParameterFilter {
            ($ArgumentList[0] -eq 'PC1') -and
            ($ArgumentList[1] -eq 'b')
        }
        Mock Start-Job -MockWith {
            & $realStartJobCommand -Scriptblock {
                @(
                    [PSCustomObject]@{
                        ComputerName = 'PC2'
                        DatabaseName = 'b'
                        SqlFile      = 'c:\query1.sql'
                        Executed     = $true
                        Duration     = '00:00:00:2b1'
                        Error        = $null
                    }
                    [PSCustomObject]@{
                        ComputerName = 'PC2'
                        DatabaseName = 'b'
                        SqlFile      = 'c:\query2.sql'
                        Executed     = $true
                        Duration     = '00:00:00:2b2'
                        Error        = $null
                    }
                )
            }
        } -ParameterFilter {
            ($ArgumentList[0] -eq 'PC2') -and
            ($ArgumentList[1] -eq 'b')
        }

        .$testScript @testParams
    }
    It 'Start-Job is called for each ComputerName and each database' {
        Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($ScriptBlock) -and
            ($ArgumentList[0] -eq 'PC1') -and
            ($ArgumentList[1] -eq 'a') -and
            ($ArgumentList[2].Count -eq 2)
        }
        Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($ScriptBlock) -and
            ($ArgumentList[0] -eq 'PC1') -and
            ($ArgumentList[1] -eq 'b') -and
            ($ArgumentList[2].Count -eq 2)
        }
        Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($ScriptBlock) -and
            ($ArgumentList[0] -eq 'PC2') -and
            ($ArgumentList[1] -eq 'a') -and
            ($ArgumentList[2].Count -eq 2)
        }
        Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($ScriptBlock) -and
            ($ArgumentList[0] -eq 'PC2') -and
            ($ArgumentList[1] -eq 'b') -and
            ($ArgumentList[2].Count -eq 2)
        }
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    ($_.ComputerName -eq $testRow.ComputerName) -and
                    ($_.DatabaseName -eq $testRow.DatabaseName) -and
                    ($_.SqlFile -eq $testRow.SqlFile)
                }
                $actualRow.Executed | Should -Be $testRow.Executed
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.Duration | Should -Be $testRow.Duration
            }
        }
    }
    It 'send a summary mail to the user' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq 'bob@contoso.com') -and
            ($Bcc -eq $ScriptAdmin) -and
            ($Priority -eq 'High') -and
            ($Subject -eq '8 queries, 1 error') -and
            ($Attachments.Count -eq 1) -and
            ($Attachments -like '* - Log.xlsx') -and
            ($Message -like "*<th>Total queries</th>*<td>8</td>*<th>Executed queries</th>*<td>7</td>*<th>Not executed queries</th>*<td>1</td>*<th>Failed queries</th>*<td>1</td>*<p><i>* Check the attachment for details</i></p>*")
        }
    }
}
Describe 'when a job fails' {
    BeforeAll {
        @{
            MailTo             = 'bob@contoso.com'
            MaxConcurrentTasks = 6
            Tasks              = @(
                @{
                    ComputerName = @('PC1')
                    DatabaseName = @('a')
                    SqlFiles     = $testData.sqlFile.Path
                }
            )
        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        $testExportedExcelRows = @(
            [PSCustomObject]@{
                ComputerName = 'PC1'
                DatabaseName = 'a'
                SqlFiles     = $testData.sqlFile.Path.Count
                Error        = "'PC1\a' job error 'oops'"
            }
        )

        Mock Start-Job -MockWith {
            & $realStartJobCommand -Scriptblock {
                throw 'oops'
            }
        }

        .$testScript @testParams
    }
    It 'Start-Job is called' {
        Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe -ParameterFilter {
            ($ScriptBlock) -and
            ($ArgumentList[0] -eq 'PC1') -and
            ($ArgumentList[1] -eq 'a') -and
            ($ArgumentList[2].Count -eq 2)
        }
        Should -Invoke Start-Job -Times 1 -Exactly -Scope Describe
    }
    Context 'export errors to an Excel file' {
        BeforeAll {
            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

            {
                Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
            } | Should -Throw

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'JobErrors' -EA Ignore
        }
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    ($_.ComputerName -eq $testRow.ComputerName) -and
                    ($_.DatabaseName -eq $testRow.DatabaseName) -and
                    ($_.SqlFile -eq $testRow.SqlFile)
                }
                $actualRow.Executed | Should -Be $testRow.Executed
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.Duration | Should -Be $testRow.Duration
            }
        }
    }
    It 'send a summary mail to the user' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq 'bob@contoso.com') -and
            ($Bcc -eq $ScriptAdmin) -and
            ($Priority -eq 'High') -and
            ($Subject -eq '2 queries, 1 error') -and
            ($Attachments.Count -eq 1) -and
            ($Attachments -like '* - Log.xlsx') -and
            ($Message -like "*<th>Total queries</th>*<td>2</td>*<th>Executed queries</th>*<td>0</td>*<th>Not executed queries</th>*<td>2</td>*<th>Failed queries</th>*<td>0</td>*<th>Job errors</th>*<td>1</td>*<p><i>* Check the attachment for details</i></p>*")
        }
    }
}