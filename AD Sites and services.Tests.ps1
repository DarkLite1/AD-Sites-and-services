#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName  = 'Test'
        OU          = 'OU=XXX,OU=EU,DC=contoso,DC=net'
        CountryCode = 'XXX'
        MailTo      = 'bob@contoso.com'
        LogFolder   = "TestDrive:/log" 
        ScriptAdmin = 'admin@contoso.com'
    }

    Mock Get-ADReplicationSite
    Mock Get-ADReplicationSubnet
    Mock Get-ADUserHC
    Mock Get-PrintersInstalledHC
    Mock Get-ServersHC
    Mock Send-MailHC
    Mock Write-EventLog
}

Describe 'logging' {
    It 'create log folder' {
        .$testScript @testParams
        $testParams.LogFolder | Should -Exist
    } 
}
Describe 'send an email to admin when' {
    It "the file ComputersNotInOU is not found" {
        .$testScript @testParams -ComputersNotInOU '.NonExistingFile.txt'

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            ($To -eq $ScriptAdmin) -and 
            ($Priority -eq 'High') -and 
            ($Subject -eq 'FAILURE') -and 
            ($Message -like "*File * not found*")
        }
    } 
}
Describe 'create an Excel file when' {
    BeforeEach {
        Remove-Item "$($testParams.LogFolder)\*" -Recurse -Force -EA Ignore
    }
    It "AD Replication sites are found" {
        Mock Get-ADReplicationSite {
            [PSCustomObject]@{
                Name = 'XXX-My Site-1'
            }
        }

        .$testScript @testParams

        Get-ChildItem -Path $testParams.LogFolder | 
        Where-Object { $_.Name -like '*Test AD Sites and subnets.xlsx' } | 
        Should -HaveCount 1
    } 
    It "AD Replication subnets are found" {
        Mock Get-ADReplicationSubnet {
            [PSCustomObject]@{
                Name = '10.10.10.00/2'
            }
        }

        .$testScript @testParams

        Get-ChildItem -Path $testParams.LogFolder | 
        Where-Object { $_.Name -like '*Test AD Sites and subnets.xlsx' } | 
        Should -HaveCount 1
    } 
    It "AD Users are found with an office that does not exist in one of the subnets" {
        Mock Get-ADUserHC {
            [PSCustomObject]@{
                'Logon name' = 'Bob'
                Office       = 'Brussels'
            }
        }
        Mock Get-ADReplicationSubnet {
            [PSCustomObject]@{
                Name     = '10.10.10.00/2'
                Location = 'Leuven'
            }
        }

        .$testScript @testParams

        Get-ChildItem -Path $testParams.LogFolder | 
        Where-Object { $_.Name -like '*AD Users.xlsx' } | 
        Should -HaveCount 1
    } 
    It "Installed printers are found with a Location that does not exist in one of the subnets" {
        Mock Get-PrintersInstalledHC {
            [PSCustomObject]@{
                ServerName = 'S1'
                Printers   = [PSCustomObject]@{
                    ComputerName = 'S1'
                    Name         = 'Bob'
                    Location     = 'Brussels'
                }
            }
        }
        Mock Get-ADReplicationSubnet {
            [PSCustomObject]@{
                Name     = '10.10.10.00/2'
                Location = 'Leuven'
            }
        }

        .$testScript @testParams

        Get-ChildItem -Path $testParams.LogFolder | 
        Where-Object { $_.Name -like '*Printers installed.xlsx' } | 
        Should -HaveCount 1
    } 
}
Describe 'do not create an Excel file when' {
    BeforeEach {
        Remove-Item "$($testParams.LogFolder)\*" -Recurse -Force -EA Ignore
    }
    It "AD Users are found with an office that does exist in one of the subnets" {
        Mock Get-ADUserHC {
            [PSCustomObject]@{
                'Logon name' = 'Bob'
                Office       = 'Brussels'
            }
        }
        Mock Get-ADReplicationSubnet {
            [PSCustomObject]@{
                Name     = '10.10.10.00/2'
                Location = 'Brussels'
            }
        }

        .$testScript @testParams

        Get-ChildItem -Path $testParams.LogFolder | 
        Where-Object { $_.Name -like '*AD Users.xlsx' } | 
        Should -BeNullOrEmpty
    } 
    It "Installed printers are found with a Location that does exist in one of the subnets" {
        Mock Get-PrintersInstalledHC {
            [PSCustomObject]@{
                ServerName = 'S1'
                Printers   = [PSCustomObject]@{
                    ComputerName = 'S1'
                    Name         = 'Bob'
                    Location     = 'Leuven'
                }
            }
        }
        Mock Get-ADReplicationSubnet {
            [PSCustomObject]@{
                Name     = '10.10.10.00/2'
                Location = 'Leuven'
            }
        }

        .$testScript @testParams

        Get-ChildItem -Path $testParams.LogFolder | 
        Where-Object { $_.Name -like '*Printers installed.xlsx' } | 
        Should -BeNullOrEmpty
    } 
}

