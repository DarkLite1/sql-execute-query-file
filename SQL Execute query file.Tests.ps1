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
        ScriptName  = 'Test (Brecht)'
        ImportFile  = $testOutParams.FilePath
        SqlScript   = (New-Item 'TestDrive:/s.ps1' -ItemType File).FullName
        LogFolder   = New-Item 'TestDrive:/log' -ItemType Directory
        ScriptAdmin = '007@contoso.com'
    }

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
}
Describe 'on a successful run' {
    BeforeAll {
        $Error.Clear()

        $testData.ScriptOutput = @(
            [PSCustomObject]@{
                ComputerName = 'PC1'
                DatabaseName = 'dn1'
                SqlFile      = $testData.sqlFile[0].Path
                StartTime    = Get-Date
                EndTime      = (Get-Date).AddMinutes(5)
                Executed     = $true
                Error        = $null
                Output       = @('a', 'b')
            }
            [PSCustomObject]@{
                ComputerName = 'PC2'
                DatabaseName = 'db1'
                SqlFile      = $testData.sqlFile[1].Path
                StartTime    = Get-Date
                EndTime      = (Get-Date).AddMinutes(10)
                Executed     = $false
                Error        = 'prob'
                Output       = $null
            }
        )

        Mock Invoke-Command {
            $testData.ScriptOutput
        }

        $testExportedExcelRows = $testData.ScriptOutput |
        Select-Object -Property 'ComputerName',
        'DatabaseName', 'StartTime', 'EndTime', 'Executed',
        @{
            Name       = 'Duration'
            Expression = {
                New-TimeSpan -Start $_.StartTime -End $_.EndTime
            }
        },
        'SqlFile', @{
            Name       = 'Output'
            Expression = {
                $_.Output -join ', '
            }
        }, 'Error'

        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams

        $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

        $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
    }
    Context 'export an Excel file' {
        It 'to the log folder' {
            $testExcelLogFile | Should -Not -BeNullOrEmpty
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                ($_.ComputerName -eq $testRow.ComputerName)
                }
                $actualRow.DatabaseName | Should -Be $testRow.DatabaseName
                $actualRow.StartTime.ToString('yyyyMMddhhmm') |
                Should -Be $testRow.StartTime.ToString('yyyyMMddhhmm')
                $actualRow.EndTime.ToString('yyyyMMddhhmm') |
                Should -Be $testRow.EndTime.ToString('yyyyMMddhhmm')
                $actualRow.Duration | Should -Not -BeNullOrEmpty
                $actualRow.Executed | Should -Be $testRow.Executed
                $actualRow.SqlFile | Should -Be $testRow.SqlFile
                $actualRow.Output | Should -Be $testRow.Output
                $actualRow.Error | Should -Be $testRow.Error
            }
        }
    }
    Context 'send an e-mail ' {
        It 'to the user' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq $testInputFile.MailTo) -and
            ($Bcc -eq $testParams.ScriptAdmin) -and
            ($Priority -eq 'High') -and
            ($Subject -eq '2 queries, 1 error') -and
            ($Attachments.Count -eq 1) -and
            ($Attachments -like '* - Log.xlsx') -and
            ($Message -like "*Total queries*2*Executed queries*1*Not executed queries*1*Failed queries*1*Check the attachment for details*")
            }
        }
    }
}