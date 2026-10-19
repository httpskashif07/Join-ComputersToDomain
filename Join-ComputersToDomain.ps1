<#
.SYNOPSIS
    Remotely joins multiple computers, listed in a CSV file, to an Active Directory domain
    and generates a report of the results.

.DESCRIPTION
    This script reads a list of computer names from a CSV file and, for each computer:
        1. Connects to it remotely (WMI) to confirm it is reachable and check its current domain.
        2. Skips it if it is already a member of the target domain (no restart required).
        3. Joins it to the domain with Add-Computer, optionally placing it in a specific OU.
        4. Restarts it (a restart is required for a domain join to take effect) and waits
           for it to come back online.
        5. Reconnects to verify that the computer is now a member of the domain.

    When all computers have been processed, the script displays a summary and exports a
    CSV report showing which computers were successfully joined.

    Credentials are never stored in the script. They are requested securely with
    Get-Credential when not supplied as parameters.

    Requirements:
        - Windows PowerShell 5.1 (Add-Computer is not available in PowerShell 7).
        - Run as Administrator on a computer that can reach the targets (e.g. the DC).
        - DNS on each target computer must point to the domain controller.
        - Each target must allow remote WMI ("Windows Management Instrumentation (WMI)"
          firewall rules enabled) and have a local administrator account.

.PARAMETER CsvPath
    Path to the CSV file containing the computers to join. The file must have a column
    named "ComputerName".

.PARAMETER DomainName
    Fully qualified name of the domain to join. Default: Kashif.powershell.com

.PARAMETER OUPath
    Optional distinguished name of the OU for the new computer accounts, for example
    "OU=Workstations,DC=Kashif,DC=powershell,DC=com". If omitted, computers are placed
    in the default Computers container.

.PARAMETER DomainCredential
    Domain account with permission to join computers to the domain.
    If omitted, you are prompted for it.

.PARAMETER LocalCredential
    Local administrator account that exists on the target computers.
    If omitted, you are prompted for it.

.PARAMETER ReportPath
    Path of the CSV report to create. Default: .\DomainJoinReport_<date-time>.csv

.PARAMETER RestartTimeoutSeconds
    How long to wait for each computer to come back online after restarting. Default: 600

.EXAMPLE
    Join all computers listed in Computers.csv:
    PS> .\Join-ComputersToDomain.ps1 -CsvPath .\Computers.csv

.EXAMPLE
    Join computers and place them in the Workstations OU:
    PS> .\Join-ComputersToDomain.ps1 -CsvPath .\Computers.csv -OUPath "OU=Workstations,DC=Kashif,DC=powershell,DC=com"

.EXAMPLE
    Preview what the script would do without changing anything:
    PS> .\Join-ComputersToDomain.ps1 -CsvPath .\Computers.csv -WhatIf

.NOTES
    Author: Kashif Sharma
    Created: 2026-10-19
    Version: 1.0
    Last Modified: 2026-10-19
    Change Log:
        1.0 - Initial release

.LINK
    https://learn.microsoft.com/powershell/module/microsoft.powershell.management/add-computer
#>

#Requires -Version 5.1
#Requires -PSEdition Desktop
#Requires -RunAsAdministrator

#Parameters
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$CsvPath,

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$DomainName = 'Kashif.powershell.com',

    [Parameter()]
    [string]$OUPath,

    [Parameter()]
    [System.Management.Automation.PSCredential]$DomainCredential,

    [Parameter()]
    [System.Management.Automation.PSCredential]$LocalCredential,

    [Parameter()]
    [string]$ReportPath = (Join-Path -Path (Get-Location) -ChildPath ("DomainJoinReport_{0:yyyyMMdd_HHmmss}.csv" -f (Get-Date))),

    [Parameter()]
    [ValidateRange(60, 3600)]
    [int]$RestartTimeoutSeconds = 600
)

#Functions

function Import-ComputerList {
    <#
    .SYNOPSIS
        Reads computer names from the CSV file and removes blanks and duplicates.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $rows = Import-Csv -Path $Path

    if (-not $rows) {
        throw "The CSV file '$Path' is empty."
    }
    if ($rows[0].PSObject.Properties.Name -notcontains 'ComputerName') {
        throw "The CSV file '$Path' must contain a column named 'ComputerName'."
    }

    $rows |
        ForEach-Object { "$($_.ComputerName)".Trim() } |
        Where-Object { $_ } |
        Sort-Object -Unique
}

function Get-RemoteDomainStatus {
    <#
    .SYNOPSIS
        Connects to a remote computer with WMI and returns its name and domain membership.
        Throws an error if the computer cannot be reached.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Computer,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]$Credential
    )

    Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer -Credential $Credential -ErrorAction Stop |
        Select-Object -Property Name, Domain, PartOfDomain
}

function New-JoinResult {
    <#
    .SYNOPSIS
        Creates one row for the final report.
    #>
    param(
        [string]$Computer,
        [string]$PreviousDomain = '',
        [string]$Status,
        [bool]$Restarted = $false,
        [bool]$JoinedToDomain = $false,
        [string]$Message = ''
    )

    [pscustomobject]@{
        ComputerName   = $Computer
        PreviousDomain = $PreviousDomain
        Status         = $Status
        Restarted      = $Restarted
        JoinedToDomain = $JoinedToDomain
        Message        = $Message
        Timestamp      = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }
}

function Join-RemoteComputer {
    <#
    .SYNOPSIS
        Joins one remote computer to the domain, restarts it, verifies the join and
        returns a report row.
    #>
    param(
        [Parameter(Mandatory = $true)] [string]$Computer,
        [Parameter(Mandatory = $true)] [string]$Domain,
        [Parameter(Mandatory = $true)] [System.Management.Automation.PSCredential]$DomainCred,
        [Parameter(Mandatory = $true)] [System.Management.Automation.PSCredential]$LocalCred,
        [string]$OU,
        [int]$TimeoutSeconds = 600
    )

    # Step 1: Connect remotely and check the current domain membership.
    try {
        $before = Get-RemoteDomainStatus -Computer $Computer -Credential $LocalCred
    }
    catch {
        return New-JoinResult -Computer $Computer -Status 'Unreachable' `
            -Message "Could not connect: $($_.Exception.Message)"
    }

    $previousDomain = if ($before.PartOfDomain) { $before.Domain } else { "Workgroup: $($before.Domain)" }

    # Step 2: Skip computers that are already members of the target domain.
    if ($before.PartOfDomain -and $before.Domain -eq $Domain) {
        return New-JoinResult -Computer $Computer -PreviousDomain $previousDomain -Status 'AlreadyJoined' `
            -JoinedToDomain $true -Message 'Already a member of the domain. No restart required.'
    }

    # Step 3: Join the computer to the domain.
    $joinParams = @{
        ComputerName    = $Computer
        LocalCredential = $LocalCred
        DomainName      = $Domain
        Credential      = $DomainCred
        Force           = $true
        PassThru        = $true
        ErrorAction     = 'Stop'
    }
    if ($OU) { $joinParams.OUPath = $OU }

    try {
        $joinResult = Add-Computer @joinParams
        if (-not $joinResult.HasSucceeded) {
            return New-JoinResult -Computer $Computer -PreviousDomain $previousDomain -Status 'JoinFailed' `
                -Message 'Add-Computer reported that the join did not succeed.'
        }
    }
    catch {
        return New-JoinResult -Computer $Computer -PreviousDomain $previousDomain -Status 'JoinFailed' `
            -Message $_.Exception.Message
    }

    # Step 4: Restart the computer (required for the join to take effect) and wait for it.
    try {
        Restart-Computer -ComputerName $Computer -Credential $LocalCred -Force `
            -Wait -For Wmi -Timeout $TimeoutSeconds -Delay 5 -ErrorAction Stop
    }
    catch {
        return New-JoinResult -Computer $Computer -PreviousDomain $previousDomain -Status 'RestartFailed' `
            -Message "Joined, but the restart could not be confirmed: $($_.Exception.Message)"
    }

    # Step 5: Reconnect and verify the computer is now a member of the domain.
    try {
        $after = Get-RemoteDomainStatus -Computer $Computer -Credential $LocalCred
        if ($after.PartOfDomain -and $after.Domain -eq $Domain) {
            return New-JoinResult -Computer $Computer -PreviousDomain $previousDomain -Status 'Joined' `
                -Restarted $true -JoinedToDomain $true -Message "Successfully joined $Domain."
        }
        return New-JoinResult -Computer $Computer -PreviousDomain $previousDomain -Status 'VerificationFailed' `
            -Restarted $true -Message "After restart the computer reports domain '$($after.Domain)'."
    }
    catch {
        return New-JoinResult -Computer $Computer -PreviousDomain $previousDomain -Status 'VerificationFailed' `
            -Restarted $true -Message "Restarted, but could not reconnect to verify: $($_.Exception.Message)"
    }
}

#Execution Code

# 1. Read the list of computers from the CSV file.
try {
    $computers = @(Import-ComputerList -Path $CsvPath)
}
catch {
    Write-Error $_.Exception.Message
    return
}
Write-Host "Found $($computers.Count) computer(s) in '$CsvPath'." -ForegroundColor Cyan

# 2. Confirm the domain can be resolved in DNS before going further.
try {
    Resolve-DnsName -Name $DomainName -ErrorAction Stop | Out-Null
}
catch {
    Write-Error "Cannot resolve '$DomainName'. Check that this computer's DNS points to the domain controller."
    return
}

# 3. Collect credentials securely (no passwords are stored in the script).
if (-not $DomainCredential) {
    $DomainCredential = Get-Credential -UserName "$DomainName\Administrator" `
        -Message "Enter a domain account that can join computers to $DomainName"
    if (-not $DomainCredential) { Write-Error 'Domain credentials are required.'; return }
}
if (-not $LocalCredential) {
    $LocalCredential = Get-Credential -UserName 'Administrator' `
        -Message 'Enter a local administrator account that exists on the target computers'
    if (-not $LocalCredential) { Write-Error 'Local administrator credentials are required.'; return }
}

# 4. Process each computer.
$results = New-Object System.Collections.Generic.List[object]
$counter = 0

foreach ($computer in $computers) {
    $counter++
    Write-Progress -Activity "Joining computers to $DomainName" -Status "$computer ($counter of $($computers.Count))" `
        -PercentComplete (($counter / $computers.Count) * 100)

    if ($PSCmdlet.ShouldProcess($computer, "Join to domain '$DomainName' and restart")) {
        Write-Host "[$counter/$($computers.Count)] Processing $computer..." -ForegroundColor Yellow
        $result = Join-RemoteComputer -Computer $computer -Domain $DomainName -DomainCred $DomainCredential `
            -LocalCred $LocalCredential -OU $OUPath -TimeoutSeconds $RestartTimeoutSeconds
        $results.Add($result)

        $colour = if ($result.JoinedToDomain) { 'Green' } else { 'Red' }
        Write-Host "    $($result.Status): $($result.Message)" -ForegroundColor $colour
    }
}
Write-Progress -Activity "Joining computers to $DomainName" -Completed

# 5. Generate the report.
if ($results.Count -gt 0) {
    $results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

    $joined = @($results | Where-Object { $_.JoinedToDomain }).Count
    $failed = $results.Count - $joined

    Write-Host ''
    Write-Host '===== Domain Join Report =====' -ForegroundColor Cyan
    $results | Format-Table -Property ComputerName, Status, Restarted, JoinedToDomain, Message -AutoSize -Wrap
    Write-Host "Domain          : $DomainName"
    Write-Host "Total computers : $($results.Count)"
    Write-Host "In domain       : $joined" -ForegroundColor Green
    Write-Host "Failed          : $failed" -ForegroundColor $(if ($failed) { 'Red' } else { 'Green' })
    Write-Host "Report saved to : $ReportPath" -ForegroundColor Cyan
}
