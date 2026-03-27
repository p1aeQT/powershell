# ================================================================
#  USB DEVICE HISTORY VIEWER
#  Registry-based | PowerShell 5.1 compatible
# ================================================================

# ── Convert registry FILETIME bytes to DateTime ──────────────────
function Convert-FileTime {
    param([byte[]]$Bytes)
    if (-not $Bytes -or $Bytes.Count -lt 8) { return $null }
    try {
        $ft = [BitConverter]::ToInt64($Bytes, 0)
        if ($ft -le 0) { return $null }
        return [DateTime]::FromFileTimeUtc($ft).ToLocalTime()
    } catch { return $null }
}

# ── Read first install / last connect / last disconnect ──────────
function Get-DeviceTimestamps {
    param([string]$DevKeyPath)
    $propBase   = "$DevKeyPath\Properties\{83da6326-97a6-4088-9453-a1923f573b29}"
    $first      = $null
    $lastPlug   = $null
    $lastUnplug = $null
    foreach ($sub in @("0064","0065","0066")) {
        $subPath = "$propBase\00$sub"
        try {
            $val   = Get-ItemProperty -Path "Registry::$subPath" -ErrorAction SilentlyContinue
            $bytes = $null
            if ($val) {
                $raw = $val.'(default)'
                if (-not $raw) {
                    $props = $val.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' }
                    if ($props) { $raw = ($props | Select-Object -First 1).Value }
                }
                if ($raw -is [byte[]]) { $bytes = $raw }
            }
            $dt = Convert-FileTime $bytes
            switch ($sub) {
                "0064" { $first      = $dt }
                "0065" { $lastPlug   = $dt }
                "0066" { $lastUnplug = $dt }
            }
        } catch { }
    }
    return [PSCustomObject]@{
        FirstInstall     = $first
        LastConnected    = $lastPlug
        LastDisconnected = $lastUnplug
    }
}

# ── Loading bar ──────────────────────────────────────────────────
function Show-LoadingBar {
    param(
        [int]$Percent,
        [string]$Label
    )
    $width    = 40
    $filled   = [math]::Floor($width * $Percent / 100)
    $empty    = $width - $filled
    $bar      = "█" * $filled + "░" * $empty
    $padLabel = $Label.PadRight(32)

    Write-Host "`r  $padLabel [" -NoNewline -ForegroundColor DarkGray
    Write-Host $bar -NoNewline -ForegroundColor Cyan
    Write-Host "] " -NoNewline -ForegroundColor DarkGray
    Write-Host "$($Percent.ToString().PadLeft(3))%" -NoNewline -ForegroundColor Yellow
}

# ── Splash screen ────────────────────────────────────────────────
function Show-Splash {
    Clear-Host
    Write-Host ""
    Write-Host "                                                        " -ForegroundColor DarkCyan
    Write-Host "    ██╗   ██╗███████╗██████╗     ██╗  ██╗██╗███████╗████████╗" -ForegroundColor Cyan
    Write-Host "    ██║   ██║██╔════╝██╔══██╗    ██║  ██║██║██╔════╝╚══██╔══╝" -ForegroundColor Cyan
    Write-Host "    ██║   ██║███████╗██████╔╝    ███████║██║███████╗   ██║   " -ForegroundColor Cyan
    Write-Host "    ██║   ██║╚════██║██╔══██╗    ██╔══██║██║╚════██║   ██║   " -ForegroundColor Cyan
    Write-Host "    ╚██████╔╝███████║██████╔╝    ██║  ██║██║███████║   ██║   " -ForegroundColor Cyan
    Write-Host "     ╚═════╝ ╚══════╝╚═════╝     ╚═╝  ╚═╝╚═╝╚══════╝   ╚═╝   " -ForegroundColor DarkCyan
    Write-Host ""
    Write-Host "    ──────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "              R E G I S T R Y   B A S E D   S C A N          " -ForegroundColor DarkGray
    Write-Host "    ──────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $adminTag = if ($isAdmin) { "YES" } else { "NO  [!] Run as Admin for full timestamps" }
    $adminCol = if ($isAdmin) { "Green" } else { "DarkYellow" }
    Write-Host "    Host      : $($env:COMPUTERNAME)   |   Admin: " -NoNewline -ForegroundColor DarkGray
    Write-Host $adminTag -ForegroundColor $adminCol
    Write-Host ""
}

# ── Section divider ──────────────────────────────────────────────
function Write-Divider {
    param([string]$Title, [System.ConsoleColor]$Color)
    $line = "═" * 60
    Write-Host ""
    Write-Host "  ╔$line╗" -ForegroundColor $Color
    $pad  = [math]::Floor((60 - $Title.Length) / 2)
    $rpad = 60 - $Title.Length - $pad
    Write-Host "  ║$(' ' * $pad)$Title$(' ' * $rpad)║" -ForegroundColor $Color
    Write-Host "  ╚$line╝" -ForegroundColor $Color
    Write-Host ""
}

# ── Device card ──────────────────────────────────────────────────
function Write-DeviceCard {
    param(
        [string]$FriendlyName,
        [string]$DeviceType,
        [string]$InstanceId,
        [string]$Vid,
        [string]$Pid,
        [string]$SerialNumber,
        $FirstInstall,
        $LastConnected,
        $LastDisconnected,
        [int]$Index
    )

    Write-Host "  ┌─[ DEVICE #$Index ]──────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  │" -ForegroundColor DarkGray

    Write-Host "  │   " -NoNewline -ForegroundColor DarkGray
    Write-Host "  $FriendlyName" -ForegroundColor White

    Write-Host "  │" -ForegroundColor DarkGray

    if ($DeviceType) {
        Write-Host "  │   Type          " -NoNewline -ForegroundColor DarkGray
        Write-Host "»  $DeviceType" -ForegroundColor Gray
    }
    if ($Vid -or $Pid) {
        Write-Host "  │   VID / PID     " -NoNewline -ForegroundColor DarkGray
        Write-Host "»  $Vid / $Pid" -ForegroundColor Gray
    }
    if ($SerialNumber) {
        Write-Host "  │   Serial        " -NoNewline -ForegroundColor DarkGray
        Write-Host "»  $SerialNumber" -ForegroundColor Gray
    }
    if ($InstanceId) {
        Write-Host "  │   Instance ID   " -NoNewline -ForegroundColor DarkGray
        Write-Host "»  $InstanceId" -ForegroundColor DarkGray
    }

    Write-Host "  │" -ForegroundColor DarkGray

    # First install
    Write-Host "  │   " -NoNewline -ForegroundColor DarkGray
    Write-Host "⊕ First Install  " -NoNewline -ForegroundColor DarkGray
    if ($FirstInstall) {
        Write-Host "»  $($FirstInstall.ToString('yyyy-MM-dd  HH:mm:ss'))" -ForegroundColor Yellow
    } else {
        Write-Host "»  unknown" -ForegroundColor DarkGray
    }

    # Last connected
    Write-Host "  │   " -NoNewline -ForegroundColor DarkGray
    Write-Host "▲ Last Plugged   " -NoNewline -ForegroundColor DarkGray
    if ($LastConnected) {
        Write-Host "»  $($LastConnected.ToString('yyyy-MM-dd  HH:mm:ss'))" -ForegroundColor Green
    } else {
        Write-Host "»  unknown" -ForegroundColor DarkGray
    }

    # Last disconnected
    Write-Host "  │   " -NoNewline -ForegroundColor DarkGray
    Write-Host "▼ Last Unplugged " -NoNewline -ForegroundColor DarkGray
    if ($LastDisconnected) {
        Write-Host "»  $($LastDisconnected.ToString('yyyy-MM-dd  HH:mm:ss'))" -ForegroundColor Red
    } else {
        Write-Host "»  unknown" -ForegroundColor DarkGray
    }

    Write-Host "  │" -ForegroundColor DarkGray
    Write-Host "  └────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
}

# ════════════════════════════════════════════════════════════════════
#  MAIN — collect everything silently first, then display
# ════════════════════════════════════════════════════════════════════

Show-Splash

Write-Host "  Scanning your machine for USB history..." -ForegroundColor DarkGray
Write-Host ""

$allDevices  = @()
$pnpResults  = @()

# ── Step 1 of 3: USBSTOR ────────────────────────────────────────
Show-LoadingBar -Percent 0 -Label "[ 1/3 ] Scanning USBSTOR..."
Start-Sleep -Milliseconds 200

$usbstorRoot = "HKLM:\SYSTEM\CurrentControlSet\Enum\USBSTOR"
if (Test-Path $usbstorRoot) {
    $deviceClasses = Get-ChildItem -Path $usbstorRoot -ErrorAction SilentlyContinue
    $classCount    = @($deviceClasses).Count
    $ci            = 0

    foreach ($class in $deviceClasses) {
        $ci++
        $pct = [math]::Floor(($ci / [math]::Max($classCount,1)) * 30)
        Show-LoadingBar -Percent $pct -Label "[ 1/3 ] Scanning USBSTOR..."

        $classNameRaw = $class.PSChildName
        $devType = ""
        $vendor  = ""
        $product = ""

        if ($classNameRaw -match 'Disk')   { $devType = "Mass Storage  (Disk)" }
        if ($classNameRaw -match 'CdRom')  { $devType = "CD / DVD-ROM" }
        if ($classNameRaw -match 'Tape')   { $devType = "Tape Drive" }
        if ($classNameRaw -match 'Other')  { $devType = "Other Storage" }
        if ($classNameRaw -match 'Ven_([^&]+)')  { $vendor  = $Matches[1] }
        if ($classNameRaw -match 'Prod_([^&]+)') { $product = $Matches[1] }

        $instances = Get-ChildItem -Path $class.PSPath -ErrorAction SilentlyContinue
        foreach ($inst in $instances) {
            $serial      = $inst.PSChildName
            $cleanSerial = $serial -replace '&\d+$', ''
            $props       = Get-ItemProperty -Path $inst.PSPath -ErrorAction SilentlyContinue
            $fname       = $props.FriendlyName
            if (-not $fname) {
                if ($vendor -and $product) { $fname = "$vendor $product" }
                else { $fname = $classNameRaw }
            }
            $instanceId = "USBSTOR\$classNameRaw\$serial"
            $times      = Get-DeviceTimestamps -DevKeyPath $inst.PSPath
            $allDevices += [PSCustomObject]@{
                FriendlyName     = $fname
                DeviceType       = $devType
                Bus              = "USBSTOR"
                VID              = ""
                PID              = ""
                SerialNumber     = $cleanSerial
                InstanceId       = $instanceId
                FirstInstall     = $times.FirstInstall
                LastConnected    = $times.LastConnected
                LastDisconnected = $times.LastDisconnected
                SortTime         = if ($times.LastConnected) { $times.LastConnected } elseif ($times.FirstInstall) { $times.FirstInstall } else { [datetime]::MinValue }
            }
        }
    }
}

# ── Step 2 of 3: USB ────────────────────────────────────────────
Show-LoadingBar -Percent 33 -Label "[ 2/3 ] Scanning USB devices..."
Start-Sleep -Milliseconds 200

$usbRoot = "HKLM:\SYSTEM\CurrentControlSet\Enum\USB"
if (Test-Path $usbRoot) {
    $vidPidKeys = Get-ChildItem -Path $usbRoot -ErrorAction SilentlyContinue
    $vpCount    = @($vidPidKeys).Count
    $vi         = 0

    foreach ($vidPidKey in $vidPidKeys) {
        $vi++
        $pct = 33 + [math]::Floor(($vi / [math]::Max($vpCount,1)) * 33)
        Show-LoadingBar -Percent $pct -Label "[ 2/3 ] Scanning USB devices..."

        $vidPidName = $vidPidKey.PSChildName
        $vid = ""
        $usbPid = ""
        if ($vidPidName -match 'VID_([0-9A-Fa-f]+)') { $vid = $Matches[1] }
        if ($vidPidName -match 'PID_([0-9A-Fa-f]+)') { $usbPid = $Matches[1] }

        $instances = Get-ChildItem -Path $vidPidKey.PSPath -ErrorAction SilentlyContinue
        foreach ($inst in $instances) {
            $serial  = $inst.PSChildName
            $props   = Get-ItemProperty -Path $inst.PSPath -ErrorAction SilentlyContinue
            $fname   = $props.FriendlyName
            $devClass = $props.Class
            $service  = $props.Service
            if (-not $fname)     { $fname   = $vidPidName }
            if ($devClass)       { $devType = $devClass   } else { $devType = "USB Device" }
            if ($service)        { $devType = "$devType ($service)".Trim(" ()") }
            $instanceId = "USB\$vidPidName\$serial"
            $times      = Get-DeviceTimestamps -DevKeyPath $inst.PSPath
            $allDevices += [PSCustomObject]@{
                FriendlyName     = $fname
                DeviceType       = $devType
                Bus              = "USB"
                VID              = $vid
                PID              = $usbPid
                SerialNumber     = $serial
                InstanceId       = $instanceId
                FirstInstall     = $times.FirstInstall
                LastConnected    = $times.LastConnected
                LastDisconnected = $times.LastDisconnected
                SortTime         = if ($times.LastConnected) { $times.LastConnected } elseif ($times.FirstInstall) { $times.FirstInstall } else { [datetime]::MinValue }
            }
        }
    }
}

# ── Step 3 of 3: pnputil Security log ───────────────────────────
Show-LoadingBar -Percent 66 -Label "[ 3/3 ] Checking pnputil removals..."
Start-Sleep -Milliseconds 200

$since = (Get-Date).AddDays(-90)
try {
    $secEvents = Get-WinEvent -FilterHashtable @{
        LogName   = 'Security'
        Id        = 4688
        StartTime = $since
    } -ErrorAction SilentlyContinue | Where-Object { $_.Message -match 'pnputil' }

    if ($secEvents) {
        foreach ($ev in ($secEvents | Sort-Object TimeCreated -Descending)) {
            $cmdLine = ""
            if ($ev.Message -match 'Process Command Line:\s+(.+)') { $cmdLine = $Matches[1].Trim() }
            if ($cmdLine -notmatch '/remove-device|/remove') { continue }
            $instanceId = ""
            if ($cmdLine -match '"([^"]+)"')                       { $instanceId = $Matches[1] }
            elseif ($cmdLine -match '/remove-device\s+(\S+)')      { $instanceId = $Matches[1] }
            $pnpResults += [PSCustomObject]@{
                Time       = $ev.TimeCreated
                InstanceId = $instanceId
                Command    = $cmdLine
            }
        }
    }
} catch { }

Show-LoadingBar -Percent 100 -Label "[ ✓  ] Scan complete!          "
Start-Sleep -Milliseconds 400

# ════════════════════════════════════════════════════════════════════
#  NOW DISPLAY EVERYTHING
# ════════════════════════════════════════════════════════════════════

Clear-Host
Show-Splash

# ── Legend ───────────────────────────────────────────────────────
Write-Host "  LEGEND   " -NoNewline -ForegroundColor DarkGray
Write-Host " ⊕ First Install" -NoNewline -ForegroundColor Yellow
Write-Host "   ▲ Last Plugged" -NoNewline -ForegroundColor Green
Write-Host "   ▼ Last Unplugged" -ForegroundColor Red
Write-Host ""

# ── Storage Devices (USBSTOR) ────────────────────────────────────
$storDevices = $allDevices | Where-Object { $_.Bus -eq "USBSTOR" } | Sort-Object SortTime -Descending

Write-Divider -Title "  STORAGE DEVICES  ( $(@($storDevices).Count) found )  " -Color Cyan

if (@($storDevices).Count -eq 0) {
    Write-Host "  No USB storage devices found." -ForegroundColor DarkGray
    Write-Host ""
} else {
    $idx = 1
    foreach ($d in $storDevices) {
        Write-DeviceCard `
            -FriendlyName     $d.FriendlyName `
            -DeviceType       $d.DeviceType `
            -InstanceId       $d.InstanceId `
            -Vid              $d.VID `
            -Pid              $d.PID `
            -SerialNumber     $d.SerialNumber `
            -FirstInstall     $d.FirstInstall `
            -LastConnected    $d.LastConnected `
            -LastDisconnected $d.LastDisconnected `
            -Index            $idx
        $idx++
    }
}

# ── Other USB Devices ────────────────────────────────────────────
$otherDevices = $allDevices | Where-Object { $_.Bus -eq "USB" } | Sort-Object SortTime -Descending

Write-Divider -Title "  OTHER USB DEVICES  ( $(@($otherDevices).Count) found )  " -Color DarkCyan

if (@($otherDevices).Count -eq 0) {
    Write-Host "  No other USB devices found." -ForegroundColor DarkGray
    Write-Host ""
} else {
    $idx = 1
    foreach ($d in $otherDevices) {
        Write-DeviceCard `
            -FriendlyName     $d.FriendlyName `
            -DeviceType       $d.DeviceType `
            -InstanceId       $d.InstanceId `
            -Vid              $d.VID `
            -Pid              $d.PID `
            -SerialNumber     $d.SerialNumber `
            -FirstInstall     $d.FirstInstall `
            -LastConnected    $d.LastConnected `
            -LastDisconnected $d.LastDisconnected `
            -Index            $idx
        $idx++
    }
}

# ── pnputil Removals ─────────────────────────────────────────────
Write-Divider -Title "  FORCED REMOVALS  ( pnputil )  " -Color Magenta

if ($pnpResults.Count -eq 0) {
    Write-Host "  No forced removals found in Security log (last 90 days)." -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  TIP: To capture these in the future, enable Process Creation" -ForegroundColor DarkGray
    Write-Host "       Auditing via:  secpol.msc  »  Advanced Audit Policy  " -ForegroundColor DarkGray
    Write-Host "       »  Detailed Tracking  »  Audit Process Creation  »  Success" -ForegroundColor DarkGray
    Write-Host ""
} else {
    $ri = 1
    foreach ($r in $pnpResults) {
        Write-Host "  ┌─[ REMOVAL #$ri ]──────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host "  │" -ForegroundColor DarkGray
        Write-Host "  │   " -NoNewline -ForegroundColor DarkGray
        Write-Host "  Time      »  $($r.Time.ToString('yyyy-MM-dd  HH:mm:ss'))" -ForegroundColor Cyan
        Write-Host "  │   " -NoNewline -ForegroundColor DarkGray
        Write-Host "  Instance  »  $($r.InstanceId)" -ForegroundColor White
        Write-Host "  │   " -NoNewline -ForegroundColor DarkGray
        Write-Host "  Command   »  $($r.Command)" -ForegroundColor Gray
        Write-Host "  │" -ForegroundColor DarkGray
        Write-Host "  └────────────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host ""
        $ri++
    }
}

# ── Summary footer ───────────────────────────────────────────────
$total = $allDevices.Count
Write-Host "  ╔════════════════════════════════════════════════════════════╗" -ForegroundColor DarkGray
Write-Host "  ║  " -NoNewline -ForegroundColor DarkGray
Write-Host "  SCAN COMPLETE" -NoNewline -ForegroundColor Green
Write-Host "  ·  " -NoNewline -ForegroundColor DarkGray
Write-Host "$total total device(s) found" -NoNewline -ForegroundColor White
Write-Host "  ·  " -NoNewline -ForegroundColor DarkGray
Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -NoNewline -ForegroundColor Cyan
$padR = 62 - 16 - "$total total device(s) found".Length - (Get-Date -Format 'yyyy-MM-dd HH:mm:ss').Length
Write-Host (" " * [math]::Max($padR,1)) -NoNewline
Write-Host "║" -ForegroundColor DarkGray
Write-Host "  ╚════════════════════════════════════════════════════════════╝" -ForegroundColor DarkGray
Write-Host ""
