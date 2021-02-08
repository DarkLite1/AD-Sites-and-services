<#
    .SYNOPSIS
        Send a summary of all sites and subnets in an Excel file to the user.

    .DESCRIPTION
        Send a summary of all sites and subnets in an Excel file to the user. For all users that have
        an office that is unknown in the subnets list, an Excel file will be created. Same goes for
        all installed printers on servers that have an office location that is unknown.

        This script is intended to run as a scheduled task and will not change anything. It will only
        report on the found anomalies.
 #>

Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String[]]$OU,
    [Parameter(Mandatory)]
    [String[]]$CountryCode,
    [Parameter(Mandatory)]
    [String[]]$MailTo,
    [String]$ComputersNotInOU,
    [String]$LogFolder = $env:POWERSHELL_LOG_FOLDER,
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Try {
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        #region Logging
        $LogParams = @{
            LogFolder    = New-FolderHC -Path $LogFolder -ChildPath "AD Reports\AD Sites and services\$ScriptName"
            Name         = $ScriptName
            Date         = 'ScriptStartTime'
            NoFormatting = $true
        }
        $LogFile = New-LogFileNameHC @LogParams

        $ExcelParams = @{
            AutoSize     = $true
            FreezeTopRow = $true
        }

        $MailAttachments = @()
        #endregion

        if ($ComputersNotInOU -and (-not (Test-Path -LiteralPath $ComputersNotInOU -PathType Leaf))) {
            throw "File '$ComputersNotInOU' not found."
        }
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        #region Create location filter
        $SearchString = $(foreach ($C in $CountryCode) {
                "(Location -like '$C*')"
            }) -join ' -or '
        Write-EventLog @EventVerboseParams -Message "CountryCode search string is '$SearchString'"

        $Filter = [ScriptBlock]::Create($SearchString)
        #endregion

        #region Sites
        $ADReplicationSite = Get-ADReplicationSite -Filter $Filter -Properties Location, WhenCreated, 
        WhenChanged, Subnets, ObjectClass | 
        Select-Object Name, Description, Location, @{N = 'SubnetCount'; E = { $_.Subnets.Count } }, 
        ObjectClass, DistinguishedName, WhenCreated, WhenChanged

        if ($ADReplicationSite) {
            Write-EventLog @EventOutParams -Message "$($ADReplicationSite.Count) AD Replication sites found"

            $ExcelParams.Path = "$LogFile AD Sites and subnets.xlsx"
            $MailAttachments += $ExcelParams.Path

            $ADReplicationSite | Sort-Object Name | 
            Export-Excel @ExcelParams -TableName 'Sites' -WorkSheetName 'Sites'
        }
        #endregion

        #region Subnets
        $ADReplicationSubnet = Get-ADReplicationSubnet -Filter $Filter -Properties Description, WhenCreated, WhenChanged, ObjectClass | Select-Object Name, Description, Location, 
        @{Name = 'SiteName'; E = { $null = $_.Site -match '(?<=CN=)(.*?)(?=,CN=)'; $Matches[0] } }, 
        ObjectClass, DistinguishedName, WhenCreated, WhenChanged

        if ($ADReplicationSubnet) {
            Write-EventLog @EventOutParams -Message "$($ADReplicationSubnet.Count) AD Replication subnets found"

            $ExcelParams.Path = "$LogFile AD Sites and subnets.xlsx"
            $MailAttachments += $ExcelParams.Path

            $ADReplicationSubnet | Sort-Object Name | Export-Excel @ExcelParams -TableName 'Subnets' -WorkSheetName 'Subnets'
        }
        #endregion

        #region Users
        $Users = Get-ADUserHC -OU $OU | Where-Object { $ADReplicationSubnet.Location -notcontains $_.Office }

        if ($Users) {
            Write-EventLog @EventOutParams -Message "$($Users.Count) AD users found with offices that don't exist in the subnet collection"

            $ExcelParams.Path = "$LogFile AD Users.xlsx"
            $MailAttachments += $ExcelParams.Path

            $Users | Group-Object Office | Sort-Object Name | Select-Object @{Name = 'Office'; Expression = { $_.Name } }, Count | Export-Excel @ExcelParams -TableName 'Summary' -WorkSheetName 'Summary'

            $Users | Sort-Object Office | Select-Object 'Logon name', 'Display name', Office, OU | Export-Excel @ExcelParams -TableName 'Users' -WorkSheetName 'Users'
        }
        #endregion

        #region Printers
        $ServerParams = @{
            OU   = $OU
            Path = $ComputersNotInOU
        }
        Remove-EmptyParamsHC $ServerParams
        $ComputerName = Get-ServersHC @ServerParams

        $Printers = (Get-PrintersInstalledHC $ComputerName).Printers |
        Where-Object { $ADReplicationSubnet.Location -notcontains $_.Location }

        if ($Printers) {
            Write-EventLog @EventOutParams -Message "$($Printers.Count) installed printers found with locations that don't exist in the subnet collection"

            $ExcelParams.Path = "$LogFile Printers installed.xlsx"
            $MailAttachments += $ExcelParams.Path

            $Printers | Group-Object Location | Sort-Object Name |
            Select-Object @{Name = 'Location'; Expression = { $_.Name } }, Count |
            Export-Excel @ExcelParams -TableName 'Summary' -WorkSheetName 'Summary'

            $Printers | Sort-Object ComputerName, Name |
            Select-Object @{Name = 'ServerName'; Expression = { $_.ComputerName } },
            @{Name = 'PrinterName'; Expression = { $_.Name } }, Location |
            Export-Excel @ExcelParams -TableName 'Printers' -WorkSheetName 'Printers'
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    Try {
        Get-ScriptRuntimeHC -Stop

        $HtmlOu = ConvertTo-OuNameHC -OU $OU | Sort-Object | ConvertTo-HtmlListHC -Header 'Organizational units:'

        $HTML = @"
        <p></p>

        <table width="100%">
            <tr><th>Quantity</th><th>Description</th></tr>
            <tr>
                <td align="center" width="10">$(($ADReplicationSite | Measure-Object).Count)</td>
                <td width="400">Sites found in active directory where the location starts with one of the following: $($CountryCode -join ', ')</td>
                <td></td>
            </tr>
            <tr>
                <td align="center" width="10">$(($ADReplicationSubnet | Measure-Object).Count)</td>
                <td width="400">Subnets found in active directory where the location starts with one of the following: $($CountryCode -join ', ')</td>
                <td></td>
            </tr>
            <tr>
                <td align="center" width="10">$(($Users | Measure-Object).Count)</td>
                <td width="400">User accounts in the active directory where the office is not matching any of the known subnet locations</td>
                <td></td>
            </tr>
            <tr>
                <td align="center" width="10">$(($Printers | Measure-Object).Count)</td>
                <td width="400">Printers installed on servers where the location is not matching any of the known subnet locations</td>
                <td></td>
            </tr>
        </table>
"@

        $MailParams = @{
            To          = $MailTo
            Bcc         = $ScriptAdmin
            Subject     = "$(($ADReplicationSite | Measure-Object).Count) sites and $(($ADReplicationSubnet | Measure-Object).Count) subnets"
            Message     = $HTML, $HtmlOu
            Attachments = $MailAttachments | Select-Object -Unique
            LogFolder   = $LogParams.LogFolder
            Header      = $ScriptName
            Save        = $LogFile + ' - Mail.html'
        }
        Remove-EmptyParamsHC $MailParams
        Send-MailHC @MailParams
        Write-EventLog @EventOutParams -Message ($env:USERNAME + ' - ' + 'Mail sent')
        Write-EventLog @EventEndParams
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"; Exit 1
    }
}