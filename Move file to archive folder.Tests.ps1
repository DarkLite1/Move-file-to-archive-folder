#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $testFolder = @{
        Source      = (New-Item 'TestDrive:/Source' -ItemType Directory).FullName
        Destination = (New-Item 'TestDrive:/Destination' -ItemType Directory).FullName
    }

    $testData = @(
        [PSCustomObject]@{
            FileName              = 'File1.txt'
            FileCreationTime      = Get-Date
            DestinationFolderPath = 'z:\Folder'
            Action                = 'File moved'
            Error                 = $null
        }
        [PSCustomObject]@{
            FileName              = 'File2.txt'
            FileCreationTime      = Get-Date
            DestinationFolderPath = 'z:\Folder'
            Action                = $null
            Error                 = 'Failed to move'
        }
    )

    $testInputFile = @{
        SendMail          = @{
            To   = 'bob@contoso.com'
            When = 'Always'
        }
        MaxConcurrentJobs = 1
        Tasks             = @(
            @{
                ComputerName = 'PC1'
                SourceFolder = $testFolder.Source
                Destination  = @{
                    Folder      = $testFolder.Destination
                    ChildFolder = 'Year\Month'
                }
                OlderThan    = @{
                    Quantity = 1
                    Unit     = 'Month'
                }
                Option       = @{
                    DuplicateFile = $null
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
        MoveScript  = (New-Item 'TestDrive:/script.ps1' -ItemType File).FullName
        ImportFile  = $testOutParams.FilePath
        LogFolder   = New-Item 'TestDrive:/log' -ItemType Directory
        ScriptAdmin = 'admin@contoso.com'
    }

    Mock Invoke-Command {
        $testData
    } -ParameterFilter {
        $FilePath -eq $testParams.MoveScript
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
    It 'the file MoveScript cannot be found' {
        $testNewParams = Copy-ObjectHC $testParams
        $testNewParams.MoveScript = 'c:\upDoesNotExist.ps1'

        $testInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*Move script with path '$($testNewParams.MoveScript)' not found*")
        }
        Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
            $EntryType -eq 'Error'
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
            It '<_> not found' -ForEach @(
                'SendMail', 'MaxConcurrentJobs', 'Tasks'
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
            It 'SendMail.<_> not found' -ForEach @(
                'To', 'When'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.SendMail.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'SendMail.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks.<_> not found' -ForEach @(
                'SourceFolder', 'Destination', 'OlderThan', 'Option'
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
            It 'Tasks.Destination.<_> not found' -ForEach @(
                'Folder', 'ChildFolder'
            ) {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Destination.$_ = $null

                $testNewInputFile | ConvertTo-Json -Depth 7 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and
                        ($Message -like "*$ImportFile*Property 'Tasks.Destination.$_' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Destination.ChildFolder not supported' {
                $testNewInputFile = Copy-ObjectHC $testInputFile
                $testNewInputFile.Tasks[0].Destination.ChildFolder = 'Wrong'

                $testNewInputFile | ConvertTo-Json -Depth 5 |
                Out-File @testOutParams

                .$testScript @testParams

                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and ($Message -like "*$ImportFile*Value 'wrong' is not supported by 'Destination.ChildFolder'. Valid options are 'Year-Month', 'Year\Month', 'Year' or 'YYYYMM'.*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context 'OlderThan' {
                Context 'OlderThan.Unit' {
                    It 'not found' {
                        $testNewInputFile = Copy-ObjectHC $testInputFile
                        $testNewInputFile.Tasks[0].OlderThan.Remove("Unit")

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
                        $testNewInputFile.Tasks[0].OlderThan.Unit = 'notSupported'

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
                Context 'OlderThan.Quantity' {
                    It 'not found' {
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
            Context 'Option.DuplicateFile' {
                It 'not found' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].Option.Remove('DuplicateFile')

                    $testNewInputFile | ConvertTo-Json -Depth 5 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'Option.DuplicateFile' not found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'is not supported' {
                    $testNewInputFile = Copy-ObjectHC $testInputFile
                    $testNewInputFile.Tasks[0].Option.DuplicateFile = 'wrong'

                    $testNewInputFile | ConvertTo-Json -Depth 5 |
                    Out-File @testOutParams

                    .$testScript @testParams

                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Value 'wrong' is not supported by 'Option.DuplicateFile'. Valid options are NULL, 'OverwriteFile' or 'RenameFile'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
        }
    }
}
Describe 'execute the move script' {
    BeforeAll {
        $testNewInputFile = Copy-ObjectHC $testInputFile

        $testJobArguments = {
            ($FilePath -eq $testParams.MoveScript) -and
            ($ArgumentList[0] -eq $testNewInputFile.Tasks[0].SourceFolder) -and
            ($ArgumentList[1] -eq $testNewInputFile.Tasks[0].Destination.Folder) -and
            ($ArgumentList[2] -eq $testNewInputFile.Tasks[0].Destination.ChildFolder) -and
            ($ArgumentList[3] -eq $testNewInputFile.Tasks[0].OlderThan.Unit) -and
            ($ArgumentList[4] -eq $testNewInputFile.Tasks[0].OlderThan.Quantity) -and
            ($ArgumentList[5] -eq $testNewInputFile.Tasks[0].Option.DuplicateFile)
        }
    }
    It 'with Invoke-Command when Tasks.ComputerName is not the localhost' {
        $testNewInputFile.Tasks[0].ComputerName = 'PC1'

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams

        Should -Invoke Invoke-Command -Times 1 -Exactly -ParameterFilter {
            (& $testJobArguments) -and
            ($ComputerName -eq 'PC1')
        }
    }
    It 'with direct invocation when Tasks.ComputerName is the localhost' {
        $testNewInputFile.Tasks[0].ComputerName = $env:COMPUTERNAME

        $testNewInputFile | ConvertTo-Json -Depth 7 |
        Out-File @testOutParams

        .$testScript @testParams

        Should -Not -Invoke Invoke-Command
    }
}
Describe 'create an Excel file' {
    BeforeAll {
        $testNewInputFile = Copy-ObjectHC $testInputFile
        $testNewInputFile.Tasks[0].OlderThan.Quantity = 3
        $testNewInputFile.Tasks[0].OlderThan.Unit = 'Day'

        $testNewInputFile | ConvertTo-Json -Depth 5 |
        Out-File @testOutParams

        . $testScript @testParams

        $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'
    }
    It 'in the log folder' {
        $testExcelLogFile | Should -Not -BeNullOrEmpty
    }
    Context "with sheet 'Overview'" {
        BeforeAll {
            $testExportedExcelRows = @(
                @{
                    ComputerName      = $testInputFile.Tasks[0].ComputerName
                    OlderThan         = '3 Days'
                    SourceFolder      = $testInputFile.Tasks[0].SourceFolder
                    DestinationFolder = $testData[0].DestinationFolderPath
                    FileName          = $testData[0].FileName
                    FileCreationTime  = $testData[0].FileCreationTime
                    Action            = $testData[0].Action
                    Error             = $testData[0].Error
                }
                @{
                    ComputerName      = $testInputFile.Tasks[0].ComputerName
                    OlderThan         = '3 Days'
                    SourceFolder      = $testInputFile.Tasks[0].SourceFolder
                    DestinationFolder = $testData[1].DestinationFolderPath
                    FileName          = $testData[1].FileName
                    FileCreationTime  = $testData[1].FileCreationTime
                    Action            = $testData[1].Action
                    Error             = $testData[1].Error
                }
            )

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Overview'
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $testExportedExcelRows) {
                $actualRow = $actual | Where-Object {
                    $_.FileName -eq $testRow.FileName
                }
                $actualRow.ComputerName | Should -Be $testRow.ComputerName
                $actualRow.SourceFolder | Should -Be $testRow.SourceFolder
                $actualRow.DestinationFolder | Should -Be $testRow.DestinationFolder
                $actualRow.FileCreationTime.ToString('yyyyMMdd HHmmss') |
                Should -Be $testRow.FileCreationTime.ToString('yyyyMMdd HHmmss')
                $actualRow.OlderThan | Should -Be $testRow.OlderThan
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.Action | Should -Be $testRow.Action
            }
        }
    }
    Context "with sheet 'Errors'" {
        BeforeAll {
            Remove-Item -Path $testParams.LogFolder -Recurse -Force

            Mock Invoke-Command {
                throw 'Oops'
            } -ParameterFilter {
                $FilePath -eq $testParams.MoveScript
            }

            . $testScript @testParams

            $testExportedExcelRows = @(
                @{
                    ComputerName = $testInputFile.Tasks[0].ComputerName
                    SourceFolder = $testInputFile.Tasks[0].SourceFolder
                    Error        = 'Oops'
                }
            )

            $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '*.xlsx'

            $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Errors'
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $testExportedExcelRows.Count
        }
        It 'with the correct data in the rows' {
            $actualRow.ComputerName | Should -Be $testRow.ComputerName
            $actualRow.SourceFolder | Should -Be $testRow.SourceFolder
            $actualRow.Error | Should -Be $testRow.Error
        }
    }
}
Describe 'SendMail.When' {
    BeforeAll {
        $testParamFilter = @{
            ParameterFilter = { $To -eq $testNewInputFile.SendMail.To }
        }
    }
    BeforeEach {
        $error.Clear()
    }
    Context 'send no e-mail to the user' {
        BeforeAll {
            Mock Invoke-Command {
            } -ParameterFilter {
                $FilePath -eq $testParams.MoveScript
            }
        }
        It "'Never'" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'Never'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Not -Invoke Send-MailHC @testParamFilter
        }
        It "'OnlyOnError' and no errors are found" {
            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnError'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Not -Invoke Send-MailHC
        }
        It "'OnlyOnErrorOrAction' and there are no errors and no actions" {
            Mock Invoke-Command {
            } -ParameterFilter {
                $FilePath -eq $testParams.MoveScript
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnErrorOrAction'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Not -Invoke Send-MailHC
        }
    }
    Context 'send an e-mail to the user' {
        It "'OnlyOnError' and there are errors" {
            Mock Invoke-Command {
                $testData[1]
            } -ParameterFilter {
                $FilePath -eq $testParams.MoveScript
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnError'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC @testParamFilter
        }
        It "'OnlyOnErrorOrAction' and there are actions but no errors" {
            Mock Invoke-Command {
                $testData[0]
            } -ParameterFilter {
                $FilePath -eq $testParams.MoveScript
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnErrorOrAction'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC @testParamFilter
        }
        It "'OnlyOnErrorOrAction' and there are errors but no actions" {
            Mock Invoke-Command {
                $testData[1]
            } -ParameterFilter {
                $FilePath -eq $testParams.MoveScript
            }

            $testNewInputFile = Copy-ObjectHC $testInputFile
            $testNewInputFile.SendMail.When = 'OnlyOnErrorOrAction'

            $testNewInputFile | ConvertTo-Json -Depth 7 |
            Out-File @testOutParams

            .$testScript @testParams

            Should -Invoke Send-MailHC @testParamFilter
        }
    }
}
Describe 'send an e-mail' {
    BeforeAll {
        $error.Clear()
        $testNewInputFile = Copy-ObjectHC $testInputFile

        $testNewInputFile | ConvertTo-Json -Depth 5 |
        Out-File @testOutParams

        . $testScript @testParams
    }
    It 'to the user' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq $testNewInputFile.SendMail.To) -and
            ($Bcc -eq $testParams.ScriptAdmin) -and
            ($Priority -eq 'High') -and
            ($Subject -eq '1 file moved, 1 error') -and
            ($Attachments -like '*log.xlsx') -and
            ($Message -like (
                "*From: <a href=`"{0}`">{0}</a><br>To: <a href=`"{1}`">{1}</a><br>Move files older than 1 month<br>Moved: 1, <b style=`"color:red;`">errors: 1*" -f $(
                    "\\$($testNewInputFile.Tasks[0].ComputerName)\C$\$($testFolder.Source.Substring(3))"
                ),
                $(
                    "\\$($testNewInputFile.Tasks[0].ComputerName)\C$\$($testFolder.Destination.Substring(3))"
                )
            ))
        }
    }
}