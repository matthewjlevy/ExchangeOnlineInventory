<#
.SYNOPSIS
    Inventories INACTIVE Exchange Online mailboxes (mailboxes of deleted users that
    are preserved by a hold) and their archives, and exports size, item count,
    message date range, the reason each mailbox is held, and the date it became
    inactive to a CSV file.

.DESCRIPTION
    An "inactive mailbox" is the mailbox of a user whose Microsoft 365 account has
    been deleted but whose mailbox is preserved by a Litigation Hold, an eDiscovery
    or In-Place Hold, a Microsoft Purview retention policy/label, or a delay hold.
    These mailboxes are NOT returned by the standard mailbox inventory script; this
    script covers them specifically.

    For each inactive mailbox the script collects:
        - Mailbox               : Primary SMTP address the mailbox had before deletion
        - DisplayName           : Display name of the former user
        - MailboxType           : "Primary" or "Archive"
        - ExchangeGuid          : Unique mailbox identifier (used for all lookups)
        - DistinguishedName     : Unique directory identifier
        - HoldReasons           : Why the mailbox is held (Litigation / eDiscovery /
                                  In-Place / Retention Policy / Retention Label /
                                  Delay Hold), derived from the mailbox hold flags
        - InPlaceHolds          : Raw In-Place/retention hold identifiers (for audit)
        - LitigationHoldEnabled : True/False
        - BecameInactive        : Date the mailbox became inactive (WhenSoftDeleted)
        - Messages              : Total item count (ItemCount)
        - Size (bytes)          : Total mailbox size in bytes
        - Size (GB)             : Total mailbox size in gigabytes (rounded to 2 dp)
        - Oldest Message        : Received date of the oldest item across all folders
        - Newest Message        : Received date of the newest item across all folders

    Each mailbox produces one "Primary" row. If the inactive mailbox has an archive,
    a second "Archive" row is produced with the same identifiers. The CSV is suitable
    for comparing against a third-party archive of the former
    user's mailbox.

    IMPORTANT - SMTP ADDRESS COLLISIONS
    -----------------------------------
    An inactive mailbox can share its old SMTP address with a NEW active mailbox
    that has since reused that address. The "Mailbox" column deliberately shows the
    old SMTP address for comparison, but that address is NOT unique. Use the
    ExchangeGuid (or DistinguishedName) column to uniquely identify the inactive
    mailbox. All statistics in this report are looked up by ExchangeGuid, so the
    figures always refer to the correct inactive mailbox regardless of any collision.

    ================================================================================
    INSTRUCTIONS FOR THE PERSON RUNNING THIS SCRIPT
    ================================================================================

    PREREQUISITES
    -------------
    1. A Windows machine with PowerShell 5.1 or PowerShell 7+.

    2. An Exchange Online / Microsoft 365 account that has at least the
       "View-Only Recipients" role (the "Global Reader", "Compliance
       Administrator", "Exchange Administrator", or "Global Administrator" roles
       also work). This account is only used to READ mailbox information; the
       script makes no changes.

    3. The Exchange Online Management PowerShell module. To install it, open
       PowerShell and run the following command once:

           Install-Module ExchangeOnlineManagement -Scope CurrentUser

       If prompted to trust the PSGallery repository, answer "Yes" (Y).
       (If the machine has an older version, update it with:
           Update-Module ExchangeOnlineManagement )

    HOW TO RUN
    ----------
    1. Save this script somewhere convenient, e.g.
       C:\Temp\InactiveMailboxInventory.ps1

    2. Open PowerShell and change to that folder, e.g.:

           cd C:\Temp

    3. If script execution is blocked, allow it for this session only by running:

           Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

    4. Run the script. You will be prompted to sign in to Microsoft 365 in a
       browser window (this supports multi-factor authentication):

           .\InactiveMailboxInventory.ps1

       To choose where the CSV is written, supply the -OutputPath parameter:

           .\InactiveMailboxInventory.ps1 -OutputPath "C:\Temp\Inactive.csv"

    5. When it finishes, the script prints the full path to the generated CSV
       file.

    NOTES ON RUNTIME
    ----------------
    - Statistics for every inactive mailbox and archive are read individually, so
      the script may take a while. A progress bar shows how far it has got.
    - Exchange Online throttles heavy reporting workloads. The script detects
      transient throttling/timeout/connection errors and automatically retries
      each affected call with an increasing (exponential) back-off delay, and
      re-establishes the Exchange Online session if it drops mid-run.
    - If a mailbox still cannot be read after all retries, the script logs a
      warning, records it in a companion "*_failures.log" file next to the CSV,
      and continues. Always check that log so you know the report is complete.
    - On tenants with many inactive mailboxes you can pace the script with
      -ThrottleDelayMs to reduce the chance of being throttled.

.PARAMETER OutputPath
    Full path (including file name) for the CSV output. If omitted, the file is
    written to the current directory as:
        InactiveMailboxInventory_<OrganisationName>_<yyyyMMdd_HHmmss>.csv

.PARAMETER MaxRetries
    Maximum number of automatic retries per failed lookup when a transient
    (throttling/timeout/connection) error occurs. Default is 5. Set to 0 to
    disable retries.

.PARAMETER ThrottleDelayMs
    Optional pause, in milliseconds, inserted after each mailbox is processed.
    Use this on large tenants to pace the script and avoid triggering dynamic
    throttling. Default is 0 (no pause).

.EXAMPLE
    .\InactiveMailboxInventory.ps1

    Connects interactively, inventories all inactive mailboxes, and writes a
    timestamped CSV to the current directory.

.EXAMPLE
    .\InactiveMailboxInventory.ps1 -OutputPath "C:\Reports\Inactive.csv" -ThrottleDelayMs 200

    Connects interactively, paces the run, and writes the CSV to the given path.

.NOTES
    Author  : Matthew Levy (MVP)
    Purpose : Inventory of inactive (held, deleted-user) Exchange Online mailboxes,
              including hold reason and inactivation date, for comparison against a
              third-party archive.
    Version : 1.0

    This script only READS data from Exchange Online. It does not modify, move,
    or delete any mailbox, message, or hold.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 20)]
    [int]$MaxRetries = 5,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 60000)]
    [int]$ThrottleDelayMs = 0
)

# Tracks whether THIS script opened the EXO connection, so we only disconnect
# a session we created and leave any pre-existing session intact.
$script:ConnectionOpenedByScript = $false

function Test-TransientError {
    # Returns $true for errors that are worth retrying: throttling, timeouts,
    # service-busy, and dropped/expired connections.
    param($ErrorRecord)

    $msg = "$($ErrorRecord.Exception.Message)"
    return ($msg -match '(?i)(429|503|throttl|ServerBusy|TooManyRequests|timed out|timeout|operation has timed|service is unavailable|connection was closed|session .*expired|token .*expired|unable to connect)')
}

function Get-ServerBackoffSeconds {
    # Best-effort extraction of a server-suggested back-off from the error text.
    # Returns 0 when none is found.
    param($ErrorRecord)

    $msg = "$($ErrorRecord.Exception.Message)"
    if ($msg -match 'BackOffMilliseconds\D+(\d+)') {
        return [int][math]::Ceiling([int]$matches[1] / 1000)
    }
    if ($msg -match 'retry after\D+(\d+)\s*second') {
        return [int]$matches[1]
    }
    return 0
}

function Restore-EXOConnection {
    # Re-establishes the Exchange Online session if it has dropped. Uses cached
    # tokens when possible; may prompt for sign-in if the token has fully expired.
    try {
        $conn = Get-ConnectionInformation -ErrorAction SilentlyContinue
        if (-not $conn) {
            Write-Warning 'Exchange Online session lost; attempting to reconnect...'
            Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "Reconnect attempt failed: $($_.Exception.Message)"
    }
}

function Invoke-WithRetry {
    # Runs a script block, retrying transient failures with exponential back-off
    # (capped at 60s) and reconnecting the session between attempts if needed.
    # Non-transient errors, or errors past MaxRetries, are re-thrown.
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [string]$OperationName = 'operation',
        [int]$MaxRetries = 5
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return & $ScriptBlock
        }
        catch {
            $err = $_
            if (-not (Test-TransientError -ErrorRecord $err) -or $attempt -gt $MaxRetries) {
                throw
            }

            $backoff = [int][math]::Min([math]::Pow(2, $attempt), 60)
            $serverBackoff = Get-ServerBackoffSeconds -ErrorRecord $err
            if ($serverBackoff -gt $backoff) { $backoff = $serverBackoff }

            Write-Warning ("{0}: transient error (attempt {1} of {2}); retrying in {3}s -> {4}" -f `
                    $OperationName, $attempt, $MaxRetries, $backoff, $err.Exception.Message)

            Restore-EXOConnection
            Start-Sleep -Seconds $backoff
        }
    }
}

function Convert-SizeToBytes {
    # Parses the "12.34 GB (13,247,905,792 bytes)" string returned by the
    # statistics cmdlets into a plain [long] byte count. Returns 0 when empty.
    param($TotalItemSize)

    if ($null -eq $TotalItemSize) { return [long]0 }

    $text = $TotalItemSize.ToString()
    if ([string]::IsNullOrWhiteSpace($text)) { return [long]0 }

    if ($text -match '\(([\d,]+)\s*bytes\)') {
        return [long]($matches[1] -replace ',', '')
    }

    return [long]0
}

function Get-MailboxDateRange {
    # Returns a PSCustomObject with Oldest and Newest received dates across all
    # folders for an inactive mailbox (or its archive). Nulls are ignored.
    param(
        [string]$Identity,
        [switch]$Archive
    )

    $oldest = $null
    $newest = $null

    $folderParams = @{
        Identity                    = $Identity
        IncludeOldestAndNewestItems = $true
        ErrorAction                 = 'Stop'
    }
    if ($Archive) { $folderParams['Archive'] = $true }

    $folders = Invoke-WithRetry -OperationName "Get-EXOMailboxFolderStatistics ($Identity)" -MaxRetries $script:MaxRetries -ScriptBlock {
        Get-EXOMailboxFolderStatistics @folderParams
    }

    foreach ($folder in $folders) {
        if ($folder.OldestItemReceivedDate) {
            if ($null -eq $oldest -or $folder.OldestItemReceivedDate -lt $oldest) {
                $oldest = $folder.OldestItemReceivedDate
            }
        }
        if ($folder.NewestItemReceivedDate) {
            if ($null -eq $newest -or $folder.NewestItemReceivedDate -gt $newest) {
                $newest = $folder.NewestItemReceivedDate
            }
        }
    }

    return [PSCustomObject]@{
        Oldest = $oldest
        Newest = $newest
    }
}

function Get-HoldReasons {
    # Derives a human-readable list of hold reasons from the mailbox hold flags.
    # Prefix meanings (InPlaceHolds): UniH = eDiscovery hold; mbx/skp/grp =
    # Microsoft Purview retention policy; cld or bare GUID = In-Place Hold;
    # leading '-' = EXCLUDED from an org-wide policy (not a hold).
    param($Mailbox)

    $reasons = [System.Collections.Generic.List[string]]::new()

    if ($Mailbox.LitigationHoldEnabled) { $reasons.Add('Litigation Hold') }
    if ($Mailbox.ComplianceTagHoldApplied) { $reasons.Add('Retention Label Hold') }
    if ($Mailbox.DelayHoldApplied) { $reasons.Add('Delay Hold (Outlook data)') }
    
    foreach ($hold in @($Mailbox.InPlaceHolds)) {
        if ([string]::IsNullOrWhiteSpace($hold)) { continue }
        if ($hold.StartsWith('-')) { continue }               # exclusion, not a hold
        if ($hold.StartsWith('UniH')) { $reasons.Add('eDiscovery Hold'); continue }
        if ($hold.StartsWith('mbx') -or $hold.StartsWith('skp') -or $hold.StartsWith('grp')) {
            $reasons.Add('Retention Policy Hold'); continue
        }
        $reasons.Add('In-Place Hold')                          # 'cld' prefix or bare GUID
    }

    if ($reasons.Count -eq 0) {
        return 'None detected (possibly an org-wide retention policy; check Get-OrganizationConfig)'
    }
    return (($reasons | Select-Object -Unique) -join '; ')
}

function New-InactiveInventoryRow {
    # Builds a single CSV row (primary or archive) for an inactive mailbox.
    # All statistics are looked up by ExchangeGuid so SMTP-address collisions
    # with reused active addresses cannot point at the wrong mailbox.
    param(
        $Mailbox,
        [ValidateSet('Primary', 'Archive')]
        [string]$MailboxType
    )

    $identity = $Mailbox.ExchangeGuid.ToString()

    $statParams = @{
        Identity    = $identity
        Properties  = 'ItemCount', 'TotalItemSize'
        ErrorAction = 'Stop'
    }
    if ($MailboxType -eq 'Archive') { $statParams['Archive'] = $true }

    $stats = Invoke-WithRetry -OperationName "Get-EXOMailboxStatistics ($MailboxType $identity)" -MaxRetries $script:MaxRetries -ScriptBlock {
        Get-EXOMailboxStatistics @statParams
    }

    $bytes = Convert-SizeToBytes -TotalItemSize $stats.TotalItemSize
    $dates = Get-MailboxDateRange -Identity $identity -Archive:($MailboxType -eq 'Archive')

    return [PSCustomObject]@{
        'Mailbox'               = $Mailbox.PrimarySmtpAddress
        'DisplayName'           = $Mailbox.DisplayName
        'MailboxType'           = $MailboxType
        'ExchangeGuid'          = $identity
        'DistinguishedName'     = $Mailbox.DistinguishedName
        'HoldReasons'           = Get-HoldReasons -Mailbox $Mailbox
        'InPlaceHolds'          = (@($Mailbox.InPlaceHolds) -join '; ')
        'LitigationHoldEnabled' = [bool]$Mailbox.LitigationHoldEnabled
        'BecameInactive'        = $Mailbox.WhenSoftDeleted
        'Messages'              = [int]$stats.ItemCount
        'Size (bytes)'          = $bytes
        'Size (GB)'             = [math]::Round($bytes / 1GB, 2)
        'Oldest Message'        = $dates.Oldest
        'Newest Message'        = $dates.Newest
    }
}

# --- Ensure the Exchange Online Management module is available ----------------
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Error @"
The 'ExchangeOnlineManagement' module is not installed.
Install it by running the following command, then re-run this script:

    Install-Module ExchangeOnlineManagement -Scope CurrentUser
"@
    return
}

Import-Module ExchangeOnlineManagement -ErrorAction Stop

# --- Connect to Exchange Online (reuse an existing session if present) --------
try {
    $existingConnection = Get-ConnectionInformation -ErrorAction SilentlyContinue

    if (-not $existingConnection) {
        Write-Host 'Connecting to Exchange Online. A sign-in window will open...' -ForegroundColor Cyan
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        $script:ConnectionOpenedByScript = $true
    }
    else {
        Write-Host 'Using existing Exchange Online connection.' -ForegroundColor Cyan
    }
}
catch {
    Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
    return
}

# --- Build default output path using the organisation name --------------------
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $orgName = 'UnknownOrg'
    try {
        $orgName = (Get-OrganizationConfig -ErrorAction Stop).Name
    }
    catch {
        Write-Warning "Could not determine organisation name; using '$orgName'."
    }

    # Strip characters that are not valid in Windows file names.
    $safeOrg = ($orgName -replace '[\\/:*?"<>|]', '_')
    $fileName = 'InactiveMailboxInventory_{0}_{1}.csv' -f $safeOrg, (Get-Date -Format 'yyyyMMdd_HHmmss')
    $OutputPath = Join-Path -Path (Get-Location) -ChildPath $fileName
}

# Companion failure log lives next to the CSV.
$outDir = Split-Path -Parent $OutputPath
if ([string]::IsNullOrEmpty($outDir)) { $outDir = (Get-Location).Path }
$failureLogPath = Join-Path -Path $outDir -ChildPath (([System.IO.Path]::GetFileNameWithoutExtension($OutputPath)) + '_failures.log')

# --- Inventory ----------------------------------------------------------------
$results = [System.Collections.Generic.List[object]]::new()
$failures = [System.Collections.Generic.List[object]]::new()

# Properties that carry hold information are not in the default set, so request
# them explicitly.
$mailboxProps = @(
    'DisplayName', 'PrimarySmtpAddress', 'DistinguishedName', 'ExchangeGuid',
    'ArchiveStatus', 'ArchiveGuid', 'WhenSoftDeleted',
    'LitigationHoldEnabled', 'InPlaceHolds', 'ComplianceTagHoldApplied',
    'DelayHoldApplied'
)

try {
    Write-Host 'Retrieving inactive mailbox list...' -ForegroundColor Cyan
    $mailboxes = Get-EXOMailbox -InactiveMailboxOnly -ResultSize Unlimited -Properties $mailboxProps -ErrorAction Stop

    $total = @($mailboxes).Count
    Write-Host "Found $total inactive mailbox(es). Collecting statistics..." -ForegroundColor Cyan

    if ($total -eq 0) {
        Write-Host 'No inactive mailboxes found in this tenant.' -ForegroundColor Yellow
    }

    $index = 0
    foreach ($mbx in $mailboxes) {
        $index++
        $label = "$($mbx.PrimarySmtpAddress) [$($mbx.ExchangeGuid)]"

        Write-Progress -Activity 'Inventorying inactive mailboxes' `
            -Status "$index of $total : $label" `
            -PercentComplete (($index / [math]::Max($total, 1)) * 100)

        # Primary mailbox
        try {
            $results.Add((New-InactiveInventoryRow -Mailbox $mbx -MailboxType 'Primary'))
        }
        catch {
            Write-Warning "Failed to read primary statistics for '$label' after retries: $($_.Exception.Message)"
            $failures.Add([PSCustomObject]@{
                    Timestamp    = (Get-Date)
                    Mailbox      = $mbx.PrimarySmtpAddress
                    ExchangeGuid = $mbx.ExchangeGuid
                    MailboxType  = 'Primary'
                    Error        = $_.Exception.Message
                })
        }

        # Archive mailbox (only if an archive exists)
        $hasArchive = ($mbx.ArchiveStatus -eq 'Active') -or
                      ($mbx.ArchiveGuid -and $mbx.ArchiveGuid -ne [Guid]::Empty)
        if ($hasArchive) {
            try {
                $results.Add((New-InactiveInventoryRow -Mailbox $mbx -MailboxType 'Archive'))
            }
            catch {
                Write-Warning "Failed to read archive statistics for '$label' after retries: $($_.Exception.Message)"
                $failures.Add([PSCustomObject]@{
                        Timestamp    = (Get-Date)
                        Mailbox      = $mbx.PrimarySmtpAddress
                        ExchangeGuid = $mbx.ExchangeGuid
                        MailboxType  = 'Archive'
                        Error        = $_.Exception.Message
                    })
            }
        }

        if ($ThrottleDelayMs -gt 0) { Start-Sleep -Milliseconds $ThrottleDelayMs }
    }

    Write-Progress -Activity 'Inventorying inactive mailboxes' -Completed
}
catch {
    Write-Error "Failed while retrieving inactive mailboxes: $($_.Exception.Message)"
}

# --- Export -------------------------------------------------------------------
if ($results.Count -gt 0) {
    try {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host ("Done. {0} row(s) written to: {1}" -f $results.Count, (Resolve-Path -Path $OutputPath)) -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to write CSV to '$OutputPath': $($_.Exception.Message)"
    }
}
else {
    Write-Warning 'No inactive mailbox data was collected; CSV was not created.'
}

# --- Write failure log (if any mailboxes could not be read) -------------------
if ($failures.Count -gt 0) {
    try {
        $logLines = $failures | ForEach-Object {
            '{0:u}  [{1}]  {2}  ({3})  ::  {4}' -f $_.Timestamp, $_.MailboxType, $_.Mailbox, $_.ExchangeGuid, $_.Error
        }
        $header = @(
            'Inactive Mailbox Inventory - failure log',
            "Generated : $(Get-Date -Format 'u')",
            "Failures  : $($failures.Count)",
            ('-' * 60)
        )
        Set-Content -Path $failureLogPath -Value ($header + $logLines) -Encoding UTF8
        Write-Warning ("{0} mailbox/archive read(s) failed after retries. See: {1}" -f $failures.Count, $failureLogPath)
    }
    catch {
        Write-Warning "Failed to write failure log to '$failureLogPath': $($_.Exception.Message)"
    }
}
else {
    Write-Host 'No mailbox read failures.' -ForegroundColor Green
}

# --- Disconnect (only the session this script created) ------------------------
if ($script:ConnectionOpenedByScript) {
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host 'Disconnected from Exchange Online.' -ForegroundColor Cyan
    }
    catch {
        Write-Warning "Failed to disconnect cleanly: $($_.Exception.Message)"
    }
}
