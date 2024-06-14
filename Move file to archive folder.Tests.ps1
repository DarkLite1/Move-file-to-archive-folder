#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testFolder = @{
        Source      = (New-Item 'TestDrive:/Source' -ItemType Directory).FullName
        Destination = (New-Item 'TestDrive:/Destination' -ItemType Directory).FullName
    }

    $testInputFile = @{
        MailTo            = 'bob@contoso.com'
        MaxConcurrentJobs = 5
        Tasks             = @(
            @{
                SourceFolder           = '\\contoso\folderA'
                DestinationFolderPath      = '\\contoso\folderB'
                DestinationFolderStructure = 'Year\Month'
                OlderThan                  = @{
                    Quantity = 1
                    Unit     = 'Month'
                }
            }
        )
    }

    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test (Brecht)'
        ImportFile  = $testOutParams.FilePath
        LogFolder   = New-Item 'TestDrive:/log' -ItemType Directory
        ScriptAdmin = 'admin@contoso.com'
    }

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
            ($To -eq $testParams.ScriptAdmin) -and ($Priority -eq 'High') -and
            ($Subject -eq 'FAILURE')
        }
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx:://notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
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
            It 'MailTo is missing' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.MailTo = $null

                $testNewInputFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'MailTo' addresses found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks is missing' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks = $null

                $testNewInputFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'Tasks' found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'SourceFolder is missing' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].SourceFolder = $null

                $testNewInputFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'SourceFolder' found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'DestinationFolderPath is missing' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].DestinationFolderPath = $null

                $testNewInputFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'DestinationFolderPath' found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context 'DestinationFolderStructure' {
                It 'is missing' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks = @(
                        @{
                            SourceFolder      = '\\contoso\folderA'
                            DestinationFolderPath = '\\contoso\folderB'
                            # DestinationFolderStructure = "Year\\Month"
                            OlderThan             = @{
                                Quantity = 1
                                Unit     = 'Month'
                            }
                        }
                    )

                    $testNewInputFile | ConvertTo-Json -Depth 5 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'DestinationFolderStructure' found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'is not supported' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks = @(
                        @{
                            SourceFolder           = '\\contoso\folderA'
                            DestinationFolderPath      = '\\contoso\folderB'
                            DestinationFolderStructure = "wrong"
                            OlderThan                  = @{
                                Quantity = 1
                                Unit     = 'Month'
                            }
                        }
                    )

                    $testNewInputFile | ConvertTo-Json -Depth 5 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Value 'wrong' is not supported by 'DestinationFolderStructure'. Valid options are 'Year-Month', 'Year\Month', 'Year' or 'YYYYMM'.*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context 'OlderThan' {
                Context 'Unit' {
                    It 'is missing' {
                        $testNewInputFile = Copy-ObjectHC $testInputFile
                        $testNewInputFile.Tasks = @(
                            @{
                                SourceFolder           = '\\contoso\folderA'
                                DestinationFolderPath      = '\\contoso\folderB'
                                DestinationFolderStructure = 'Year\Month'
                                OlderThan                  = @{
                                    Quantity = 1
                                    # Unit     = 'Month'
                                }
                            }
                        )

                        $testNewInputFile | ConvertTo-Json -Depth 5 |
                        Out-File @testOutParams

                        .$testScript @testParams

                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'OlderThan.Unit' found*")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                    It 'is not supported' {
                        $testNewInputFile = Copy-ObjectHC $testInputFile
                        $testNewInputFile.Tasks = @(
                            @{
                                SourceFolder           = '\\contoso\folderA'
                                DestinationFolderPath      = '\\contoso\folderB'
                                DestinationFolderStructure = 'Year\Month'
                                OlderThan                  = @{
                                    Quantity = 1
                                    Unit     = 'notSupported'
                                }
                            }
                        )

                        $testNewInputFile | ConvertTo-Json -Depth 5 |
                        Out-File @testOutParams

                        .$testScript @testParams

                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Value 'notSupported' is not supported by 'OlderThan.Unit'. Valid options are 'Day', 'Month' or 'Year'*")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                }
                Context 'Quantity' {
                    It 'is missing' {
                        $testNewInputFile = Copy-ObjectHC $testInputFile
                        $testNewInputFile.Tasks[0].OlderThan.Remove("Quantity")

                        $testNewInputFile | ConvertTo-Json -Depth 5 |
                        Out-File @testOutParams

                        .$testScript @testParams

                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'OlderThan.Quantity' not found. Use value number '0' to move all files*")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                    It 'is not a number' {
                        $testNewInputFile = Copy-ObjectHC $testInputFile
                        $testNewInputFile.Tasks[0].OlderThan.Quantity = 'a'

                        $testNewInputFile | ConvertTo-Json -Depth 5 |
                        Out-File @testOutParams

                        .$testScript @testParams

                        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'OlderThan.Quantity' needs to be a number, the value 'a' is not supported*")
                        }
                        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                            $EntryType -eq 'Error'
                        }
                    }
                }
            }
        }
    }
}
Describe 'a file in the source folder' {
    Context 'is not moved when it is created more recently than' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks = @(
                @{
                    ComputerName               = $env:COMPUTERNAME
                    SourceFolder           = $testFolder.Source
                    DestinationFolderPath      = $testFolder.Destination
                    DestinationFolderStructure = 'Year\Month'
                    OlderThan                  = @{
                        Quantity = 3
                        Unit     = 'Day'
                    }
                }
            )

            $testFile = (New-Item -Path "$($testFolder.source)\file.txt" -ItemType File).FullName
        }
        It 'Day' {
            $testNewInputFile.Tasks[0].OlderThan.Unit = 'Day'

            $testNewInputFile | ConvertTo-Json -Depth 5 |
            Out-File @testOutParams

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddDays(-2)
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 1
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 0
        }
        It 'Month' {
            $testNewInputFile.Tasks[0].OlderThan.Unit = 'Month'

            $testNewInputFile | ConvertTo-Json -Depth 5 |
            Out-File @testOutParams

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddMonths(-2)
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 1
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 0
        }
        It 'Year' {
            $testNewInputFile.Tasks[0].OlderThan.Unit = 'Year'

            $testNewInputFile | ConvertTo-Json -Depth 5 |
            Out-File @testOutParams

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddYears(-2)
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 1
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 0
        }
    }
    Context 'is moved when it is older than' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks = @(
                @{
                    ComputerName               = $env:COMPUTERNAME
                    SourceFolder           = $testFolder.Source
                    DestinationFolderPath      = $testFolder.Destination
                    DestinationFolderStructure = 'Year\Month'
                    OlderThan                  = @{
                        Quantity = 3
                        Unit     = 'Day'
                    }
                }
            )
        }
        BeforeEach {
            @($testFolder.Source, $testFolder.Destination) | ForEach-Object {
                Remove-Item "$_\*" -Recurse -Force
            }
            $testFile = (New-Item -Path "$($testFolder.source)\file.txt" -ItemType File).FullName
        }
        It 'Day' {
            $testNewInputFile.Tasks[0].OlderThan.Unit = 'Day'

            $testNewInputFile | ConvertTo-Json -Depth 5 |
            Out-File @testOutParams

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddDays(-4)
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 1
        }
        It 'Month' {
            $testNewInputFile.Tasks[0].OlderThan.Unit = 'Month'

            $testNewInputFile | ConvertTo-Json -Depth 5 |
            Out-File @testOutParams

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddMonths(-4)
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 1
        }
        It 'Year' {
            $testNewInputFile.Tasks[0].OlderThan.Unit = 'Year'

            $testNewInputFile | ConvertTo-Json -Depth 5 |
            Out-File @testOutParams

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddYears(-4)
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 1
        }
    }
    Context 'is moved to a folder with structure' {
        BeforeAll {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.Tasks = @(
                @{
                    ComputerName               = $env:COMPUTERNAME
                    SourceFolder           = $testFolder.Source
                    DestinationFolderPath      = $testFolder.Destination
                    DestinationFolderStructure = 'Year'
                    OlderThan                  = @{
                        Quantity = 3
                        Unit     = 'Day'
                    }
                }
            )
        }
        BeforeEach {
            @($testFolder.Source, $testFolder.Destination) | ForEach-Object {
                Remove-Item "$_\*" -Recurse -Force
            }
            $testFile = (New-Item -Path "$($testFolder.source)\file.txt" -ItemType File).FullName
        }
        It 'Year' {
            $testNewInputFile.Tasks[0].DestinationFolderStructure = 'Year'

            $testNewInputFile | ConvertTo-Json -Depth 5 |
            Out-File @testOutParams

            $testFileCreationDate = (Get-Date).AddDays(-4)

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = $testFileCreationDate
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path (
                $testFolder.Destination + '\' +
                $testFileCreationDate.ToString('yyyy')
            ) | Should -HaveCount 1
        }
        It 'Year-Month' {
            $testNewInputFile.Tasks[0].DestinationFolderStructure = 'Year-Month'

            $testNewInputFile | ConvertTo-Json -Depth 5 |
            Out-File @testOutParams

            $testFileCreationDate = (Get-Date).AddDays(-4)

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = $testFileCreationDate
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path (
                $testFolder.Destination + '\' +
                $testFileCreationDate.ToString('yyyy') + '-' +
                $testFileCreationDate.ToString('MM')
            ) | Should -HaveCount 1
        }
        It 'Year\Month' {
            $testNewInputFile.Tasks[0].DestinationFolderStructure = 'Year\Month'

            $testNewInputFile | ConvertTo-Json -Depth 5 |
            Out-File @testOutParams

            $testFileCreationDate = (Get-Date).AddDays(-4)

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = $testFileCreationDate
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path (
                $testFolder.Destination + '\' +
                $testFileCreationDate.ToString('yyyy') + '\' +
                $testFileCreationDate.ToString('MM')
            ) | Should -HaveCount 1
        }
        It 'YYYYMM' {
            $testNewInputFile.Tasks[0].DestinationFolderStructure = 'YYYYMM'

            $testNewInputFile | ConvertTo-Json -Depth 5 |
            Out-File @testOutParams

            $testFileCreationDate = (Get-Date).AddDays(-4)

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = $testFileCreationDate
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path (
                $testFolder.Destination + '\' +
                $testFileCreationDate.ToString('yyyyMM')
            ) | Should -HaveCount 1
        }
    }
}
Describe 'on a successful run' {
    BeforeAll {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks = @(
            @{
                ComputerName               = $env:COMPUTERNAME
                SourceFolder           = $testFolder.Source
                DestinationFolderPath      = $testFolder.Destination
                DestinationFolderStructure = 'Year'
                OlderThan                  = @{
                    Quantity = 3
                    Unit     = 'Day'
                }
            }
        )

        $testNewInputFile | ConvertTo-Json -Depth 5 |
        Out-File @testOutParams

        $testFileCreationDate = (Get-Date).AddDays(-4)

        $testFile = (New-Item -Path "$($testFolder.source)\file.txt" -ItemType File).FullName

        Get-Item -Path $testFile | ForEach-Object {
            $_.CreationTime = $testFileCreationDate
        }

        $Error.Clear()
        . $testScript @testParams
    }
    Context 'export an Excel file' {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    Action                 = 'File moved'
                    ComputerName           = $env:COMPUTERNAME
                    SourceFileCreationTime = $testFileCreationDate
                    SourceFilePath         = $testFile
                    DestinationFolderPath  = "$($testFolder.Destination)\{0}" -f $testFileCreationDate.ToString('yyyy')
                    OlderThan              = '3 Days'
                    Error                  = $null
                }
            )

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
                    $_.SourceFilePath -eq $testRow.SourceFilePath
                }
                $actualRow.ComputerName | Should -Be $testRow.ComputerName
                $actualRow.SourceFileCreationTime.ToString('yyyyMMdd HHmmss') |
                Should -Be $testRow.SourceFileCreationTime.ToString('yyyyMMdd HHmmss')
                $actualRow.DestinationFolderPath | Should -Be $testRow.DestinationFolderPath
                $actualRow.OlderThan | Should -Be $testRow.OlderThan
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.Action | Should -Be $testRow.Action
            }
        }
    }
    It 'send a summary mail to the user' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq 'bob@contoso.com') -and
            ($Bcc -eq $ScriptAdmin) -and
            ($Priority -eq 'Normal') -and
            ($Subject -eq '1 file moved') -and
            ($Attachments -like '*log.xlsx') -and
            ($Message -like (
                "*From: <a href=`"{0}`">{0}</a><br>To: <a href=`"{1}`">{1}</a><br>Move files older than 3 days<br>Moved: 1*" -f $(
                    "\\$env:COMPUTERNAME\C$\$($testFolder.Source.Substring(3))"
                ),
                $(
                    "\\$env:COMPUTERNAME\C$\$($testFolder.Destination.Substring(3))"
                )
            ))
        }
    }
}