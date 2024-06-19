#Requires -Modules Pester
#Requires -Version 7

BeforeAll {
    $testFolder = @{
        Source      = (New-Item 'TestDrive:/Source' -ItemType Directory).FullName
        Destination = (New-Item 'TestDrive:/Destination' -ItemType Directory).FullName
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        SourceFolder           = $testFolder.Source
        DestinationFolder      = $testFolder.Destination
        DestinationChildFolder = 'Year\Month'
        OlderThanUnit          = 'Month'
        OlderThanQuantity      = 1
        DuplicateFile          = $null
    }
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @(
        'SourceFolder', 'DestinationFolder', 'DestinationChildFolder',
        'OlderThanUnit', 'OlderThanQuantity'
    ) {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
        Should -BeTrue
    }
}
Describe 'a file in the source folder' {
    Context 'is not moved when it is created more recently than' {
        BeforeAll {
            $testNewParams = Copy-ObjectHC $testParams
            $testNewParams.OlderThanQuantity = 3

            $testFile = (New-Item -Path "$($testFolder.source)\file.txt" -ItemType File).FullName
        }
        It 'Day' {
            $testNewParams.OlderThanUnit = 'Day'

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddDays(-2)
            }

            . $testScript @testNewParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 1
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 0
        }
        It 'Month' {
            $testNewParams.OlderThanUnit = 'Month'

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddMonths(-2)
            }

            . $testScript @testNewParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 1
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 0
        }
        It 'Year' {
            $testNewParams.OlderThanUnit = 'Year'

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddYears(-2)
            }

            . $testScript @testNewParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 1
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 0
        }
    }
    Context 'is moved when it is OlderThan' {
        BeforeAll {
            $testNewParams = Copy-ObjectHC $testParams
            $testNewParams.OlderThanQuantity = 3
        }
        BeforeEach {
            @($testFolder.Source, $testFolder.Destination) | ForEach-Object {
                Remove-Item "$_\*" -Recurse -Force
            }
            $testFile = (New-Item -Path "$($testFolder.source)\file.txt" -ItemType File).FullName
        }
        It 'Day' {
            $testNewParams.OlderThanUnit = 'Day'

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddDays(-4)
            }

            . $testScript @testNewParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 1
        }
        It 'Month' {
            $testNewParams.OlderThanUnit = 'Month'

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddMonths(-4)
            }

            . $testScript @testNewParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 1
        }
        It 'Year' {
            $testNewParams.OlderThanUnit = 'Year'

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = (Get-Date).AddYears(-4)
            }

            . $testScript @testNewParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path $testFolder.Destination | Should -HaveCount 1
        }
    }
    Context 'is moved to the DestinationChildFolder' {
        BeforeAll {
            $testNewParams = Copy-ObjectHC $testParams
            $testNewParams.OlderThanQuantity = 3
            $testNewParams.OlderThanUnit = 'Day'
        }
        BeforeEach {
            @($testFolder.Source, $testFolder.Destination) | ForEach-Object {
                Remove-Item "$_\*" -Recurse -Force
            }
            $testFile = (New-Item -Path "$($testFolder.source)\file.txt" -ItemType File).FullName
        }
        It 'Year' {
            $testNewParams.DestinationChildFolder = 'Year'

            $testFileCreationDate = (Get-Date).AddDays(-4)

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = $testFileCreationDate
            }

            . $testScript @testNewParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path (
                $testFolder.Destination + '\' +
                $testFileCreationDate.ToString('yyyy')
            ) | Should -HaveCount 1
        }
        It 'Year-Month' {
            $testNewParams.DestinationChildFolder = 'Year-Month'

            $testFileCreationDate = (Get-Date).AddDays(-4)

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = $testFileCreationDate
            }

            . $testScript @testNewParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path (
                $testFolder.Destination + '\' +
                $testFileCreationDate.ToString('yyyy') + '-' +
                $testFileCreationDate.ToString('MM')
            ) | Should -HaveCount 1
        }
        It 'Year\Month' {
            $testNewParams.DestinationChildFolder = 'Year\Month'

            $testFileCreationDate = (Get-Date).AddDays(-4)

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = $testFileCreationDate
            }

            . $testScript @testNewParams

            Get-ChildItem -Path $testFolder.Source | Should -HaveCount 0
            Get-ChildItem -Path (
                $testFolder.Destination + '\' +
                $testFileCreationDate.ToString('yyyy') + '\' +
                $testFileCreationDate.ToString('MM')
            ) | Should -HaveCount 1
        }
        It 'YYYYMM' {
            $testNewParams.DestinationChildFolder = 'YYYYMM'

            $testFileCreationDate = (Get-Date).AddDays(-4)

            Get-Item -Path $testFile | ForEach-Object {
                $_.CreationTime = $testFileCreationDate
            }

            . $testScript @testNewParams

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
        $testNewParams = Copy-ObjectHC $testParams
        $testNewParams.DestinationChildFolder = 'Year'
        $testNewParams.OlderThanQuantity = 3
        $testNewParams.OlderThanUnit = 'Day'

        $testFileCreationDate = (Get-Date).AddDays(-4)

        $testFile = New-Item -Path "$($testFolder.source)\file.txt" -ItemType File

        $testFile | Get-Item | ForEach-Object {
            $_.CreationTime = $testFileCreationDate
        }

        $Error.Clear()
        $actual = . $testScript @testNewParams
    }
    Context 'return a result object' {
        BeforeAll {
            $expected = @(
                @{
                    Action                = 'File moved'
                    FileCreationTime      = $testFileCreationDate
                    FileName              = $testFile.Name
                    DestinationFolderPath = "$($testFolder.Destination)\{0}" -f $testFileCreationDate.ToString('yyyy')
                    Error                 = $null
                }
            )
        }
        It 'with the correct total rows' {
            $actual | Should -HaveCount $expected.Count
        }
        It 'with the correct data in the rows' {
            foreach ($testRow in $expected) {
                $actualRow = $actual | Where-Object {
                    $_.FileName -eq $testRow.FileName
                }
                $actualRow.FileCreationTime.ToString('yyyyMMdd HHmmss') |
                Should -Be $testRow.FileCreationTime.ToString('yyyyMMdd HHmmss')
                $actualRow.DestinationFolderPath | Should -Be $testRow.DestinationFolderPath
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.Action | Should -Be $testRow.Action
            }
        }
    }
}