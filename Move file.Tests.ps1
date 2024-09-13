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
        Recurse                = $false
    }
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach @(
        'SourceFolder', 'DestinationFolder', 'DestinationChildFolder',
        'OlderThanUnit', 'OlderThanQuantity', 'Recurse'
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
Describe 'when a file is in a sub folder and Recurse is' {
    BeforeAll {
        $testSourceFolder = (New-Item -Path "$($testFolder.source)\Folder" -ItemType Directory).FullName

        $testSourceFile = New-Item -Path "$testSourceFolder\1.txt" -ItemType File

        $testFileCreationDate = (Get-Date).AddDays(-4)

        Get-Item -Path $testSourceFile | ForEach-Object {
            $_.CreationTime = $testFileCreationDate
        }

        $testNewParams = Copy-ObjectHC $testParams
        $testNewParams.OlderThanQuantity = 0
        $testNewParams.OlderThanUnit = 'Day'
        $testNewParams.DestinationChildFolder = 'Year'
    }
    Context 'false' {
        It 'do not move the file' {
            $testNewParams.Recurse = $false

            . $testScript @testNewParams

            $testSourceFile | Should -Exist

            "$($testFolder.Destination)\$($testFileCreationDate.ToString('yyyy'))\$($testSourceFile.Name)" |
            Should -Not -Exist
        }
    }
    Context 'true' {
        It 'move the file to the destination folder' {
            $testNewParams.Recurse = $true

            . $testScript @testNewParams

            $testSourceFile | Should -Not -Exist

            "$($testFolder.Destination)\$($testFileCreationDate.ToString('yyyy'))\$($testSourceFile.Name)" | Should -Exist
        }
    }
}
Describe 'when a file already exists in the destination folder' {
    BeforeAll {
        $testNewParams = Copy-ObjectHC $testParams
        $testNewParams.OlderThanQuantity = 0
        $testNewParams.OlderThanUnit = 'Day'
        $testNewParams.DestinationChildFolder = 'Year'

        $currentYear = (Get-Date).ToString('yyyy')
    }

    Context 'and DuplicateFile is blank' {
        BeforeAll {
            $Error.Clear()

            $testNewParams.DuplicateFile = $null

            @($testFolder.Source, $testFolder.Destination) | ForEach-Object {
                Remove-Item "$_\*" -Recurse -Force
            }

            $null = New-Item -Path "$($testFolder.Destination)\$currentYear" -ItemType Directory

            $testFile = @{
                Source      = New-Item -Path "$($testFolder.Source)\a.txt" -ItemType File
                Destination = New-Item -Path "$($testFolder.Destination)\$currentYear\a.txt" -ItemType File
            }

            $actual = . $testScript @testNewParams
        }
        It 'the file is not moved' {
            $testFile.Source | Should -Exist
            $testFile.Destination | Should -Exist
        }
        It 'no action is taken' {
            $actual.Action | Should -BeNullOrEmpty
        }
        It 'an error is thrown' {
            $actual.Error |
            Should -BeExactly "Duplicate file name in destination folder. (See 'Option.DuplicateFile: OverwriteFile or RenameFile')"

            $error | Should -HaveCount 0
        }
    }
    Context "and DuplicateFile is 'OverwriteFile'" {
        BeforeAll {
            $Error.Clear()

            $testNewParams.DuplicateFile = 'OverwriteFile'

            @($testFolder.Source, $testFolder.Destination) | ForEach-Object {
                Remove-Item "$_\*" -Recurse -Force
            }

            $null = New-Item -Path "$($testFolder.Destination)\$currentYear" -ItemType Directory

            $testFile = @{
                Source      = New-Item -Path "$($testFolder.Source)\a.txt" -ItemType File
                Destination = New-Item -Path "$($testFolder.Destination)\$currentYear\a.txt" -ItemType File
            }

            $actual = . $testScript @testNewParams
        }
        It 'the file is moved' {
            $testFile.Source | Should -Not -Exist
            $testFile.Destination | Should -Exist
        }
        It 'action is taken' {
            $actual.Action | Should -BeExactly 'File moved and overwritten'
        }
        It 'no error is thrown' {
            $actual.Error | Should -BeNullOrEmpty

            $error | Should -HaveCount 0
        }
    }
    Context "and DuplicateFile is 'RenameFile'" {
        BeforeAll {
            $Error.Clear()

            $testNewParams.DuplicateFile = 'RenameFile'

            @($testFolder.Source, $testFolder.Destination) | ForEach-Object {
                Remove-Item "$_\*" -Recurse -Force
            }

            $null = New-Item -Path "$($testFolder.Destination)\$currentYear" -ItemType Directory

            $testFile = @{
                Source      = New-Item -Path "$($testFolder.Source)\a.txt" -ItemType File
                Destination = New-Item -Path "$($testFolder.Destination)\$currentYear\a.txt" -ItemType File
            }

            $actual = . $testScript @testNewParams
        }
        It 'the file is moved' {
            $testFile.Source | Should -Not -Exist
            $testFile.Destination | Should -Exist
        }
        It 'action is taken' {
            $actual.Action | Should -BeLike "*File moved with new name 'a*.txt' due to duplicate file name*"
        }
        It 'no error is thrown' {
            $actual.Error | Should -BeNullOrEmpty

            $error | Should -HaveCount 0
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
                $actualRow.DateTime.ToString('yyyyMMdd') |
                Should -Be (Get-Date).ToString('yyyyMMdd')
                $actualRow.FileCreationTime.ToString('yyyyMMdd HHmmss') |
                Should -Be $testRow.FileCreationTime.ToString('yyyyMMdd HHmmss')
                $actualRow.DestinationFolderPath | Should -Be $testRow.DestinationFolderPath
                $actualRow.Error | Should -Be $testRow.Error
                $actualRow.Action | Should -Be $testRow.Action
            }
        }
    }
}