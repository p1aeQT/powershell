# ============================================================
#  USB Device History Viewer
#  Reads past events from Windows Event Logs
#  - USB Connections
#  - USB Disconnections
#  - pnputil /remove-device removals
# ============================================================

param(
    [int]$DaysBack = 30   # How many days back to search (default: 30)
)

# ── Helpers ───────────────────────────────────────────────────────────────────

function Write-Header {
    Clear-Host
    Write-Host ""
    Write-Host "  ╔══════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "  ║       USB DEVICE HISTORY VIEWER          ║" -ForegroundColor Yellow
    Write-Host "  ║  Connect · Disconnect · pnputil Removal  ║" -ForegroundColor Yellow
    Write-Host "  ╚══════════════════════════════════════════╝" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  Searching last $DaysBack day(s) of event logs..." -ForegroundColor DarkGray
    Write-Host "  ─────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
}

function Write-EventEntry {
    param(
        [string]$Tag,
        [ConsoleColor]$TagColor,
        [datetime]$Time,
        [string]$DeviceName,
        [string]$InstanceId,
        [string]$Extra
    )

    $ts = $Time.ToString("yyyy-MM-dd  HH:mm:ss")

    Write-Host "  [" -NoNewline -ForegroundColor DarkGray
    Write-Host $ts -NoNewline -ForegroundColor Cyan
    Write-Host "]  " -NoNewline -ForegroundColor DarkGray
    Write-Host $Tag -NoNewline -ForegroundColor $TagColor
    Write-Host ""

    if ($DeviceName) {
        Write-Host "               Device   : " -NoNewline -ForegroundColor DarkGray
        Write-Host $DeviceName -ForegroundColor White
    }
    if ($InstanceId) {
        Write-Host "               ID       : " -NoNewline -ForegroundColor DarkGray
        Write-Host $InstanceId -ForegroundColor Gray
    }
    if ($Extra) {
        Write-Host "               Info     : " -NoNewline -ForegroundColor DarkGray
        Write-Host $Extra -ForegroundColor Gray
    }
    Write-Host ""
}

function Write-SectionTitle {
    param([string]$Title, [ConsoleColor]$Color)
    Write-Host "  ┌─ $Title " -ForegroundColor $Color
    Write-Host ""
}

function Write-NoResults {
    Write-Host "  (no events found in the last $DaysBack day(s))" -ForegroundColor DarkGray
    Write-Host ""
}

# ── Start ─────────────────────────────────────────────────────────────────────

Write-Header

$since = (Get-Date).AddDays(-$DaysBack)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — USB Connections
# Source : Microsoft-Windows-DriverFrameworks-UserMode/Operational
# Event  : 2003  (device instance initialized / arrived)
# ══════════════════════════════════════════════════════════════════════════════

Write-SectionTitle "USB CONNECTIONS" Green

try {
    $connectEvents = Get-WinEvent -FilterHashtable @{
        LogName   = 'Microsoft-Windows-DriverFrameworks-UserMode/Operational'
        Id        = 2003
        StartTime = $since
    } -ErrorAction SilentlyContinue

    if ($connectEvents) {
        foreach ($ev in ($connectEvents | Sort-Object TimeCreated)) {
            $xml    = [xml]$ev.ToXml()
            $ns     = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNamespace("e", "http://schemas.microsoft.com/win/2004/08/events/event")
            $data   = $xml.SelectNodes("//e:Data", $ns)
            $instanceId = ($data | Where-Object { $_.Name -eq "InstanceId" }).'#text'
            $devDesc    = ($data | Where-Object { $_.Name -eq "DeviceDescription" }).'#text'

            Write-EventEntry `
                -Tag        " CONNECTED   " `
                -TagColor   Green `
                -Time       $ev.TimeCreated `
                -DeviceName ($devDesc -ne $null ? $devDesc : $ev.Message.Split("`n")[0].Trim()) `
                -InstanceId $instanceId
        }
    } else {
        Write-NoResults
    }
}
catch {
    Write-Host "  [WARN] Could not read connection log: $_" -ForegroundColor DarkYellow
    Write-Host ""
}

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — USB Disconnections
# Source : Microsoft-Windows-DriverFrameworks-UserMode/Operational
# Events : 2100 (surprise removal), 2102 (device removed)
# ══════════════════════════════════════════════════════════════════════════════

Write-SectionTitle "USB DISCONNECTIONS" Red

try {
    $disconnectEvents = Get-WinEvent -FilterHashtable @{
        LogName   = 'Microsoft-Windows-DriverFrameworks-UserMode/Operational'
        Id        = 2100, 2102
        StartTime = $since
    } -ErrorAction SilentlyContinue

    if ($disconnectEvents) {
        foreach ($ev in ($disconnectEvents | Sort-Object TimeCreated)) {
            $xml  = [xml]$ev.ToXml()
            $ns   = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNamespace("e", "http://schemas.microsoft.com/win/2004/08/events/event")
            $data = $xml.SelectNodes("//e:Data", $ns)

            $instanceId = ($data | Where-Object { $_.Name -eq "InstanceId" }).'#text'
            $devDesc    = ($data | Where-Object { $_.Name -eq "DeviceDescription" }).'#text'
            $reason     = if ($ev.Id -eq 2100) { "Surprise / forced removal" } else { "Normal removal" }

            Write-EventEntry `
                -Tag        " DISCONNECTED" `
                -TagColor   Red `
                -Time       $ev.TimeCreated `
                -DeviceName ($devDesc -ne $null ? $devDesc : "") `
                -InstanceId $instanceId `
                -Extra      $reason
        }
    } else {
        Write-NoResults
    }
}
catch {
    Write-Host "  [WARN] Could not read disconnection log: $_" -ForegroundColor DarkYellow
    Write-Host ""
}

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — pnputil /remove-device
# Source : Security log — Process Creation (Event 4688)
#          Fallback: System log keyword search
# ══════════════════════════════════════════════════════════════════════════════

Write-SectionTitle "pnputil /remove-device REMOVALS" Magenta

$pnpFound = $false

# --- Method A: Security log (requires Process Creation Auditing to be on) ----
try {
    $secEvents = Get-WinEvent -FilterHashtable @{
        LogName   = 'Security'
        Id        = 4688
        StartTime = $since
    } -ErrorAction SilentlyContinue |
        Where-Object { $_.Message -match 'pnputil' }

    foreach ($ev in ($secEvents | Sort-Object TimeCreated)) {
        $pnpFound = $true

        # Pull the command line from the message
        $cmdLine = ""
        if ($ev.Message -match 'Process Command Line:\s+(.+)') {
            $cmdLine = $Matches[1].Trim()
        }

        # Only show /remove-device calls
        if ($cmdLine -notmatch '/remove-device|/remove') { continue }

        $instanceId = ""
        if ($cmdLine -match '"([^"]+)"')            { $instanceId = $Matches[1] }
        elseif ($cmdLine -match '/remove-device\s+(\S+)') { $instanceId = $Matches[1] }

        Write-EventEntry `
            -Tag        " FORCE REMOVE" `
            -TagColor   Magenta `
            -Time       $ev.TimeCreated `
            -DeviceName "pnputil.exe" `
            -InstanceId $instanceId `
            -Extra      $cmdLine
    }
}
catch { <# Security log unavailable or no audit policy — fall through #> }

# --- Method B: System log — PnP device removal entries --------------------
try {
    $sysEvents = Get-WinEvent -FilterHashtable @{
        LogName      = 'System'
        ProviderName = 'Microsoft-Windows-Kernel-PnP'
        Id           = 430, 431   # Device removed / device problem
        StartTime    = $since
    } -ErrorAction SilentlyContinue

    foreach ($ev in ($sysEvents | Sort-Object TimeCreated)) {
        $xml  = [xml]$ev.ToXml()
        $ns   = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
        $ns.AddNamespace("e", "http://schemas.microsoft.com/win/2004/08/events/event")
        $data = $xml.SelectNodes("//e:Data", $ns)

        $instanceId = ($data | Where-Object { $_.Name -eq "DeviceInstanceId" }).'#text'

        # Only flag USB devices
        if ($instanceId -notmatch 'USB') { continue }

        $pnpFound = $true

        Write-EventEntry `
            -Tag        " FORCE REMOVE" `
            -TagColor   Magenta `
            -Time       $ev.TimeCreated `
            -DeviceName "Device removed (Kernel-PnP)" `
            -InstanceId $instanceId `
            -Extra      "Event ID $($ev.Id)"
    }
}
catch { <# ignore #> }

# --- Method C: Microsoft-Windows-PnPUserMode/Operational ------------------
try {
    $pnpModeEvents = Get-WinEvent -FilterHashtable @{
        LogName   = 'Microsoft-Windows-PnPUserMode/Operational'
        StartTime = $since
    } -ErrorAction SilentlyContinue |
        Where-Object { $_.Message -match 'remove|delete|uninstall' -and $_.Message -match 'USB' }

    foreach ($ev in ($pnpModeEvents | Sort-Object TimeCreated)) {
        $pnpFound = $true
        Write-EventEntry `
            -Tag        " FORCE REMOVE" `
            -TagColor   Magenta `
            -Time       $ev.TimeCreated `
            -DeviceName "PnPUserMode event" `
            -Extra      ($ev.Message.Split("`n")[0].Trim())
    }
}
catch { <# log may not exist #> }

if (-not $pnpFound) {
    Write-NoResults
    Write-Host "  [TIP] To capture pnputil removals in the future, enable Process" -ForegroundColor DarkGray
    Write-Host "        Creation Auditing in Local Security Policy:" -ForegroundColor DarkGray
    Write-Host "        secpol.msc → Advanced Audit Policy → Detailed Tracking" -ForegroundColor DarkGray
    Write-Host "        → Audit Process Creation → enable Success" -ForegroundColor DarkGray
    Write-Host ""
}

# ── Footer ────────────────────────────────────────────────────────────────────

Write-Host "  ─────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Search range : $(($since).ToString('yyyy-MM-dd HH:mm')) → $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor DarkGray
Write-Host "  Run as Admin : $( ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator) )" -ForegroundColor DarkGray
Write-Host ""
