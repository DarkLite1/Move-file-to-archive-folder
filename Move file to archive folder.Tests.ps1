#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        ImportFile = $testOutParams.FilePath
        LogFolder  = New-Item 'TestDrive:/log' -ItemType Directory
    }

    $testFolder = @{
        Source      = (New-Item 'TestDrive:/Source' -ItemType Directory).FullName
        Destination = (New-Item 'TestDrive:/Destination' -ItemType Directory).FullName
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
            ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and 
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
                @{
                    # MailTo       = @('bob@contoso.com')
                    Tasks = @()
                } | ConvertTo-Json | Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'MailTo' addresses found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'Tasks is missing' {
                @{
                    MailTo = @('bob@contoso.com')
                } | ConvertTo-Json | Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'Tasks' found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'SourceFolderPath is missing' {
                @{
                    MailTo = @('bob@contoso.com')
                    Tasks  = @(
                        @{
                            # SourceFolderPath           = '\\contoso\folderA'
                            DestinationFolderPath      = '\\contoso\folderB'
                            DestinationFolderStructure = 'Year\Month'
                            OlderThanUnit              = 'Month'
                            OlderThanQuantity          = 1
                        }
                    )
                } | ConvertTo-Json | Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'SourceFolderPath' found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'DestinationFolderPath is missing' {
                @{
                    MailTo = @('bob@contoso.com')
                    Tasks  = @(
                        @{
                            SourceFolderPath           = '\\contoso\folderA'
                            # DestinationFolderPath      = '\\contoso\folderB'
                            DestinationFolderStructure = 'Year\Month'
                            OlderThanUnit              = 'Month'
                            OlderThanQuantity          = 1
                        }
                    )
                } | ConvertTo-Json | Out-File @testOutParams
                
                .$testScript @testParams
                
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'DestinationFolderPath' found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            Context 'ComputerName' {
                It 'is used with DestinationFolderPath UNC path' {
                    @{
                        MailTo = @('bob@contoso.com')
                        Tasks  = @(
                            @{
                                ComputerName               = $env:COMPUTERNAME
                                SourceFolderPath           = '\\contoso\folderA'
                                DestinationFolderPath      = 'C:\folderB'
                                DestinationFolderStructure = 'Year\Month'
                                OlderThanUnit              = 'Month'
                                OlderThanQuantity          = 1
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                
                    .$testScript @testParams
                
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*When ComputerName is used only local paths are allowed*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'is used with SourceFolderPath UNC path' {
                    @{
                        MailTo = @('bob@contoso.com')
                        Tasks  = @(
                            @{
                                ComputerName               = $env:COMPUTERNAME
                                SourceFolderPath           = 'C:\folderA'
                                DestinationFolderPath      = '\\contoso\folderB'
                                DestinationFolderStructure = 'Year\Month'
                                OlderThanUnit              = 'Month'
                                OlderThanQuantity          = 1
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                
                    .$testScript @testParams
                
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*When ComputerName is used only local paths are allowed*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'is not used and DestinationFolderPath is a local path' {
                    @{
                        MailTo = @('bob@contoso.com')
                        Tasks  = @(
                            @{
                                # ComputerName               = $env:COMPUTERNAME
                                SourceFolderPath           = '\\contoso\folderA'
                                DestinationFolderPath      = 'C:\folderB'
                                DestinationFolderStructure = 'Year\Month'
                                OlderThanUnit              = 'Month'
                                OlderThanQuantity          = 1
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                
                    .$testScript @testParams
                
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*When local paths are used the ComputerName is mandatory*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'is not used and SourceFolderPath is a local path' {
                    @{
                        MailTo = @('bob@contoso.com')
                        Tasks  = @(
                            @{
                                # ComputerName               = $env:COMPUTERNAME
                                SourceFolderPath           = 'C:\folderA'
                                DestinationFolderPath      = '\\contoso\folderB'
                                DestinationFolderStructure = 'Year\Month'
                                OlderThanUnit              = 'Month'
                                OlderThanQuantity          = 1
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                
                    .$testScript @testParams
                
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                        (&$MailAdminParams) -and ($Message -like "*$ImportFile*When local paths are used the ComputerName is mandatory*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context 'DestinationFolderStructure' {
                It 'is missing' {
                    @{
                        MailTo = @('bob@contoso.com')
                        Tasks  = @(
                            @{
                                SourceFolderPath      = '\\contoso\folderA'
                                DestinationFolderPath = '\\contoso\folderB'
                                # DestinationFolderStructure = "Year\\Month"
                                OlderThanUnit         = 'Month'
                                OlderThanQuantity     = 1
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                
                    .$testScript @testParams
                
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'DestinationFolderStructure' found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'is not supported' {
                    @{
                        MailTo = @('bob@contoso.com')
                        Tasks  = @(
                            @{
                                SourceFolderPath           = '\\contoso\folderA'
                                DestinationFolderPath      = '\\contoso\folderB'
                                DestinationFolderStructure = "wrong"
                                OlderThanUnit              = 'Month'
                                OlderThanQuantity          = 1
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
                
                    .$testScript @testParams
                
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and ($Message -like "*$ImportFile*Value 'wrong' is not supported by 'DestinationFolderStructure'. Valid options are 'Year-Month', 'Year\Month', 'Year' or 'YYYYMM'.*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context 'OlderThanUnit' {
                It 'is missing' {
                    @{
                        MailTo = @('bob@contoso.com')
                        Tasks  = @(
                            @{
                                SourceFolderPath           = '\\contoso\folderA'
                                DestinationFolderPath      = '\\contoso\folderB'
                                DestinationFolderStructure = 'Year\Month'
                                # OlderThanUnit              = 'Month'
                                OlderThanQuantity          = 1
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
            
                    .$testScript @testParams
                            
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*$ImportFile*No 'OlderThanUnit' found*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'is not supported' {
                    @{
                        MailTo = @('bob@contoso.com')
                        Tasks  = @(
                            @{
                                SourceFolderPath           = '\\contoso\folderA'
                                DestinationFolderPath      = '\\contoso\folderB'
                                DestinationFolderStructure = 'Year\Month'
                                OlderThanUnit              = "notSupported"
                                OlderThanQuantity          = 1
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
            
                    .$testScript @testParams
                            
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*$ImportFile*Value 'notSupported' is not supported by 'OlderThanUnit'. Valid options are 'Day', 'Month' or 'Year'*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
            Context 'OlderThanQuantity' {
                It 'is missing' {
                    @{
                        MailTo = @('bob@contoso.com')
                        Tasks  = @(
                            @{
                                SourceFolderPath           = '\\contoso\folderA'
                                DestinationFolderPath      = '\\contoso\folderB'
                                DestinationFolderStructure = 'Year\Month'
                                OlderThanUnit              = 'Month'
                                # OlderThanQuantity          = 1
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams
            
                    .$testScript @testParams
                            
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'OlderThanQuantity' not found. Use value number '0' to move all files*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
                It 'is not a number' {
                    @{
                        MailTo = @('bob@contoso.com')
                        Tasks  = @(
                            @{
                                SourceFolderPath           = '\\contoso\folderA'
                                DestinationFolderPath      = '\\contoso\folderB'
                                DestinationFolderStructure = 'Year\Month'
                                OlderThanUnit              = 'Month'
                                OlderThanQuantity          = 'a'
                            }
                        )
                    } | ConvertTo-Json | Out-File @testOutParams

                    .$testScript @testParams
            
                    Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and ($Message -like "*$ImportFile*Property 'OlderThanQuantity' needs to be a number, the value 'a' is not supported*")
                    }
                    Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                        $EntryType -eq 'Error'
                    }
                }
            }
        }
    }
}
Describe 'a file in the source folder' {
    Context 'is not moved when it is created more recently than' {
        BeforeAll {
            $testFile = (New-Item -Path "$($testFolder.source)\file.txt" -ItemType File).FullName
        }
        It 'Day' {
            @{
                MailTo = @('bob@contoso.com')
                Tasks  = @(
                    @{
                        ComputerName               = $env:COMPUTERNAME
                        SourceFolderPath           = $testFolder.Source
                        DestinationFolderPath      = $testFolder.Destination
                        DestinationFolderStructure = 'Year\Month'
                        OlderThanUnit              = 'Day'
                        OlderThanQuantity          = 3
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddDays(-2)
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 1
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 0
        }
        It 'Month' {
            @{
                MailTo = @('bob@contoso.com')
                Tasks  = @(
                    @{
                        ComputerName               = $env:COMPUTERNAME
                        SourceFolderPath           = $testFolder.Source
                        DestinationFolderPath      = $testFolder.Destination
                        DestinationFolderStructure = 'Year\Month'
                        OlderThanUnit              = 'Month'
                        OlderThanQuantity          = 3
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddMonths(-2)
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 1
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 0
        }
        It 'Year' {
            @{
                MailTo = @('bob@contoso.com')
                Tasks  = @(
                    @{
                        ComputerName               = $env:COMPUTERNAME
                        SourceFolderPath           = $testFolder.Source
                        DestinationFolderPath      = $testFolder.Destination
                        DestinationFolderStructure = 'Year\Month'
                        OlderThanUnit              = 'Year'
                        OlderThanQuantity          = 3
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddYears(-2)
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 1
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 0
        }
    }
    Context 'is moved when it is older than' {
        BeforeEach {
            @($testFolder.Source, $testFolder.Destination) | ForEach-Object {
                Remove-Item "$_\*" -Recurse -Force
            }
            $testFile = (New-Item -Path "$($testFolder.source)\file.txt" -ItemType File).FullName
        }
        It 'Day' {
            @{
                MailTo = @('bob@contoso.com')
                Tasks  = @(
                    @{
                        ComputerName               = $env:COMPUTERNAME
                        SourceFolderPath           = $testFolder.Source
                        DestinationFolderPath      = $testFolder.Destination
                        DestinationFolderStructure = 'Year\Month'
                        OlderThanUnit              = 'Day'
                        OlderThanQuantity          = 3
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddDays(-4)
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 1
        }
        It 'Month' {
            @{
                MailTo = @('bob@contoso.com')
                Tasks  = @(
                    @{
                        ComputerName               = $env:COMPUTERNAME
                        SourceFolderPath           = $testFolder.Source
                        DestinationFolderPath      = $testFolder.Destination
                        DestinationFolderStructure = 'Year\Month'
                        OlderThanUnit              = 'Month'
                        OlderThanQuantity          = 3
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddMonths(-4)
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 1
        }
        It 'Year' {
            @{
                MailTo = @('bob@contoso.com')
                Tasks  = @(
                    @{
                        ComputerName               = $env:COMPUTERNAME
                        SourceFolderPath           = $testFolder.Source
                        DestinationFolderPath      = $testFolder.Destination
                        DestinationFolderStructure = 'Year\Month'
                        OlderThanUnit              = 'Year'
                        OlderThanQuantity          = 3
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddYears(-4)
            }

            . $testScript @testParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 1
        }
    }
    Context 'is moved to a folder with structure' {
        BeforeEach {
            @($testFolder.Source, $testFolder.Destination) | ForEach-Object {
                Remove-Item "$_\*" -Recurse -Force
            }
            $testFile = (New-Item -Path "$($testFolder.source)\file.txt" -ItemType File).FullName
        }
        It 'Year' {
            @{
                MailTo = @('bob@contoso.com')
                Tasks  = @(
                    @{
                        ComputerName               = $env:COMPUTERNAME
                        SourceFolderPath           = $testFolder.Source
                        DestinationFolderPath      = $testFolder.Destination
                        DestinationFolderStructure = 'Year'
                        OlderThanUnit              = 'Day'
                        OlderThanQuantity          = 3
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

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
            @{
                MailTo = @('bob@contoso.com')
                Tasks  = @(
                    @{
                        ComputerName               = $env:COMPUTERNAME
                        SourceFolderPath           = $testFolder.Source
                        DestinationFolderPath      = $testFolder.Destination
                        DestinationFolderStructure = 'Year-Month'
                        OlderThanUnit              = 'Day'
                        OlderThanQuantity          = 3
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

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
            @{
                MailTo = @('bob@contoso.com')
                Tasks  = @(
                    @{
                        ComputerName               = $env:COMPUTERNAME
                        SourceFolderPath           = $testFolder.Source
                        DestinationFolderPath      = $testFolder.Destination
                        DestinationFolderStructure = 'Year\Month'
                        OlderThanUnit              = 'Day'
                        OlderThanQuantity          = 3
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

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
            @{
                MailTo = @('bob@contoso.com')
                Tasks  = @(
                    @{
                        ComputerName               = $env:COMPUTERNAME
                        SourceFolderPath           = $testFolder.Source
                        DestinationFolderPath      = $testFolder.Destination
                        DestinationFolderStructure = 'YYYYMM'
                        OlderThanUnit              = 'Day'
                        OlderThanQuantity          = 3
                    }
                )
            } | ConvertTo-Json | Out-File @testOutParams

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
        @{
            MailTo = @('bob@contoso.com')
            Tasks  = @(
                @{
                    ComputerName               = $env:COMPUTERNAME
                    SourceFolderPath           = $testFolder.Source
                    DestinationFolderPath      = $testFolder.Destination
                    DestinationFolderStructure = 'Year'
                    OlderThanUnit              = 'Day'
                    OlderThanQuantity          = 3
                }
            )
        } | ConvertTo-Json | Out-File @testOutParams

        $testFileCreationDate = (Get-Date).AddDays(-4)

        $testFile = (New-Item -Path "$($testFolder.source)\file.txt" -ItemType File).FullName

        Get-Item -Path $testFile | ForEach-Object {
            $_.CreationTime = $testFileCreationDate
        }

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
    } -Tag test
    It 'send a summary mail to the user' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
            ($To -eq 'bob@contoso.com') -and
            ($Bcc -eq $ScriptAdmin) -and
            ($Priority -eq $testMail.Priority) -and
            ($Subject -eq $testMail.Subject) -and
            ($Attachments -like '*log.xlsx') -and
            ($Message -like $testMail.Message)
        }
    }
}
