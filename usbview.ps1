# ============================================================
#  USB Device History Viewer  (PowerShell 5.1 compatible)
#  Reads past events from Windows Event Logs
#  - USB Connections
#  - USB Disconnections
#  - pnputil /remove-device removals
# ============================================================

param(
    [int]$DaysBack = 30
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
        [System.ConsoleColor]$TagColor,
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
    param([string]$Title, [System.ConsoleColor]$Color)
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
# SECTION 1 - USB Connections
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
            $xml  = [xml]$ev.ToXml()
            $ns   = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNamespace("e", "http://schemas.microsoft.com/win/2004/08/events/event")
            $data = $xml.SelectNodes("//e:Data", $ns)

            $instanceId = ($data | Where-Object { $_.Name -eq "InstanceId" }).'#text'
            $devDesc    = ($data | Where-Object { $_.Name -eq "DeviceDescription" }).'#text'

            if ($devDesc) {
                $displayName = $devDesc
            } else {
                $displayName = $ev.Message.Split("`n")[0].Trim()
            }

            Write-EventEntry `
                -Tag        " CONNECTED   " `
                -TagColor   Green `
                -Time       $ev.TimeCreated `
                -DeviceName $displayName `
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
# SECTION 2 - USB Disconnections
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

            if ($ev.Id -eq 2100) {
                $reason = "Surprise / forced removal"
            } else {
                $reason = "Normal removal"
            }

            if ($devDesc) {
                $displayName = $devDesc
            } else {
                $displayName = ""
            }

            Write-EventEntry `
                -Tag        " DISCONNECTED" `
                -TagColor   Red `
                -Time       $ev.TimeCreated `
                -DeviceName $displayName `
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
# SECTION 3 - pnputil /remove-device
# ══════════════════════════════════════════════════════════════════════════════

Write-SectionTitle "pnputil /remove-device REMOVALS" Magenta

$pnpFound = $false

# --- Method A: Security log (Event 4688 - Process Creation) ------------------
try {
    $secEvents = Get-WinEvent -FilterHashtable @{
        LogName   = 'Security'
        Id        = 4688
        StartTime = $since
    } -ErrorAction SilentlyContinue | Where-Object { $_.Message -match 'pnputil' }

    if ($secEvents) {
        foreach ($ev in ($secEvents | Sort-Object TimeCreated)) {
            $cmdLine = ""
            if ($ev.Message -match 'Process Command Line:\s+(.+)') {
                $cmdLine = $Matches[1].Trim()
            }

            if ($cmdLine -notmatch '/remove-device|/remove') { continue }

            $instanceId = ""
            if ($cmdLine -match '"([^"]+)"') {
                $instanceId = $Matches[1]
            } elseif ($cmdLine -match '/remove-device\s+(\S+)') {
                $instanceId = $Matches[1]
            }

            $pnpFound = $true

            Write-EventEntry `
                -Tag        " FORCE REMOVE" `
                -TagColor   Magenta `
                -Time       $ev.TimeCreated `
                -DeviceName "pnputil.exe" `
                -InstanceId $instanceId `
                -Extra      $cmdLine
        }
    }
}
catch { }

# --- Method B: System log - Kernel-PnP device removal -----------------------
try {
    $sysEvents = Get-WinEvent -FilterHashtable @{
        LogName      = 'System'
        ProviderName = 'Microsoft-Windows-Kernel-PnP'
        Id           = 430, 431
        StartTime    = $since
    } -ErrorAction SilentlyContinue

    if ($sysEvents) {
        foreach ($ev in ($sysEvents | Sort-Object TimeCreated)) {
            $xml  = [xml]$ev.ToXml()
            $ns   = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $ns.AddNamespace("e", "http://schemas.microsoft.com/win/2004/08/events/event")
            $data = $xml.SelectNodes("//e:Data", $ns)

            $instanceId = ($data | Where-Object { $_.Name -eq "DeviceInstanceId" }).'#text'

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
}
catch { }

# --- Method C: PnPUserMode/Operational log -----------------------------------
try {
    $pnpModeEvents = Get-WinEvent -FilterHashtable @{
        LogName   = 'Microsoft-Windows-PnPUserMode/Operational'
        StartTime = $since
    } -ErrorAction SilentlyContinue | Where-Object {
        $_.Message -match 'remove|delete|uninstall' -and $_.Message -match 'USB'
    }

    if ($pnpModeEvents) {
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
}
catch { }

if (-not $pnpFound) {
    Write-NoResults
    Write-Host "  [TIP] To capture pnputil removals in the future, enable Process" -ForegroundColor DarkGray
    Write-Host "        Creation Auditing in Local Security Policy:" -ForegroundColor DarkGray
    Write-Host "        secpol.msc -> Advanced Audit Policy -> Detailed Tracking" -ForegroundColor DarkGray
    Write-Host "        -> Audit Process Creation -> enable Success" -ForegroundColor DarkGray
    Write-Host ""
}

# ── Footer ────────────────────────────────────────────────────────────────────

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

Write-Host "  ─────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Search range : $($since.ToString('yyyy-MM-dd HH:mm')) -> $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -ForegroundColor DarkGray
Write-Host "  Run as Admin : $isAdmin" -ForegroundColor DarkGray
Write-Host ""
