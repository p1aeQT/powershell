# ================================================================
#  USB DEVICE HISTORY VIEWER  |  Registry-based  |  PS 5.1
#  Timestamps enhanced to match USBDeview (NirSoft)
# ================================================================

function Convert-FileTime {
    param([byte[]]$Bytes)
    if (-not $Bytes -or $Bytes.Count -lt 8) { return $null }
    try {
        $ft = [BitConverter]::ToInt64($Bytes, 0)
        if ($ft -le 0) { return $null }
        return [DateTime]::FromFileTimeUtc($ft).ToLocalTime()
    } catch { return $null }
}

function Get-DeviceTimestamps {
    param([string]$DevKeyPath)
    $propBase = "$DevKeyPath\Properties\{83da6326-97a6-4088-9453-a1923f573b29}"
    
    $installTime      = $null   # Registry key LastWriteTime (Registry Time)
    $firstInstallTime = $null   # 0064
    $connectTime      = $null   # 0065
    $disconnectTime   = $null   # 0066

    # Get Install Time from registry key LastWriteTime
    try {
        $installTime = (Get-Item "Registry::$DevKeyPath" -ErrorAction SilentlyContinue).LastWriteTime
    } catch { }

    # Get the four Properties subkeys
    foreach ($sub in @("0064","0065","0066")) {
        try {
            $val = Get-ItemProperty -Path "Registry::$propBase\00$sub" -ErrorAction SilentlyContinue
            $raw = $null
            if ($val) {
                $raw = $val.'(default)'
                if (-not $raw) {
                    $p = $val.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' }
                    if ($p) { $raw = ($p | Select-Object -First 1).Value }
                }
            }
            $dt = if ($raw -is [byte[]]) { Convert-FileTime $raw } else { $null }
            switch ($sub) {
                "0064" { $firstInstallTime = $dt }
                "0065" { $connectTime      = $dt }
                "0066" { $disconnectTime   = $dt }
            }
        } catch { }
    }

    return [PSCustomObject]@{
        InstallTime      = $installTime
        FirstInstallTime = $firstInstallTime
        ConnectTime      = $connectTime
        DisconnectTime   = $disconnectTime
    }
}

function Get-ConnectionStatus {
    param($LastConnected, $LastDisconnected, [string]$InstanceId, [string]$RegKeyPath)

    if ($LastConnected -or $LastDisconnected) {
        if ($LastConnected -and $LastDisconnected) {
            if ($LastDisconnected -gt $LastConnected) { return "DISCONNECTED" }
            else                                      { return "CONNECTED"    }
        }
        if ($LastDisconnected -and -not $LastConnected) { return "DISCONNECTED" }
        if ($LastConnected    -and -not $LastDisconnected) { return "CONNECTED" }
    }

    if ($InstanceId -and $script:WmiPnpIds.Count -gt 0) {
        if ($script:WmiPnpIds -contains $InstanceId) { return "CONNECTED" }
        return "DISCONNECTED"
    }

    if ($RegKeyPath) {
        try {
            $cf = (Get-ItemProperty -Path "Registry::$RegKeyPath" -Name ConfigFlags -ErrorAction SilentlyContinue).ConfigFlags
            if ($null -ne $cf) {
                if ($cf -band 0x04) { return "DISCONNECTED" }
                else                { return "CONNECTED"    }
            }
        } catch { }
    }

    return "UNKNOWN"
}

function Show-LoadingBar {
    param([int]$Percent, [string]$Label)
    $filled = [math]::Floor(38 * $Percent / 100)
    $bar    = ("█" * $filled) + ("░" * (38 - $filled))
    Write-Host "`r  $($Label.PadRight(30)) [" -NoNewline -ForegroundColor DarkGray
    Write-Host $bar -NoNewline -ForegroundColor Cyan
    Write-Host "] " -NoNewline -ForegroundColor DarkGray
    Write-Host "$($Percent.ToString().PadLeft(3))%" -NoNewline -ForegroundColor Yellow
}

function Show-Splash {
    Clear-Host
    Write-Host ""
    Write-Host "    ██╗   ██╗███████╗██████╗     ██╗  ██╗██╗███████╗████████╗" -ForegroundColor Cyan
    Write-Host "    ██║   ██║██╔════╝██╔══██╗    ██║  ██║██║██╔════╝╚══██╔══╝" -ForegroundColor Cyan
    Write-Host "    ██║   ██║███████╗██████╔╝    ███████║██║███████╗   ██║   " -ForegroundColor Cyan
    Write-Host "    ██║   ██║╚════██║██╔══██╗    ██╔══██║██║╚════██║   ██║   " -ForegroundColor Cyan
    Write-Host "    ╚██████╔╝███████║██████╔╝    ██║  ██║██║███████║   ██║   " -ForegroundColor Cyan
    Write-Host "     ╚═════╝ ╚══════╝╚═════╝     ╚═╝  ╚═╝╚═╝╚══════╝   ╚═╝  " -ForegroundColor DarkCyan
    Write-Host "    ── REGISTRY BASED SCAN ─────────────────────────────────" -ForegroundColor DarkGray
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $aTag = if ($isAdmin) { "YES" } else { "NO (!)" }
    $aCol = if ($isAdmin) { "Green" } else { "DarkYellow" }
    Write-Host "    Host: $($env:COMPUTERNAME)  |  Admin: " -NoNewline -ForegroundColor DarkGray
    Write-Host $aTag -ForegroundColor $aCol
    Write-Host ""
}

function Write-SectionHeader {
    param([string]$Title, [System.ConsoleColor]$Color)
    $line = "─" * 58
    Write-Host "  ┌$line┐" -ForegroundColor $Color
    $pad  = [math]::Floor((58 - $Title.Length) / 2)
    $rpad = 58 - $Title.Length - $pad
    Write-Host "  │$(' ' * $pad)$Title$(' ' * $rpad)│" -ForegroundColor $Color
    Write-Host "  └$line┘" -ForegroundColor $Color
}

function Write-DeviceRow {
    param(
        [string]$FriendlyName,
        [string]$DeviceType,
        [string]$InstanceId,
        [string]$RegKeyPath,
        [string]$Vid,
        [string]$UsbPid,
        [string]$SerialNumber,
        $InstallTime,
        $FirstInstallTime,
        $ConnectTime,
        $DisconnectTime,
        [int]$Index
    )

    $status = Get-ConnectionStatus -LastConnected $ConnectTime -LastDisconnected $DisconnectTime -InstanceId $InstanceId -RegKeyPath $RegKeyPath

    $it = if ($InstallTime)      { $InstallTime.ToString('yyyy-MM-dd HH:mm')      } else { "─────────────────" }
    $fi = if ($FirstInstallTime) { $FirstInstallTime.ToString('yyyy-MM-dd HH:mm') } else { "─────────────────" }
    $ct = if ($ConnectTime)      { $ConnectTime.ToString('yyyy-MM-dd HH:mm')      } else { "─────────────────" }
    $dt = if ($DisconnectTime)   { $DisconnectTime.ToString('yyyy-MM-dd HH:mm')   } else { "─────────────────" }

    # Status badge
    $badgeText  = ""
    $badgeColor = "DarkGray"
    switch ($status) {
        "CONNECTED"    { $badgeText = " ● CONNECTED    "; $badgeColor = "Green"      }
        "DISCONNECTED" { $badgeText = " ○ DISCONNECTED "; $badgeColor = "DarkYellow" }
        "UNKNOWN"      { $badgeText = " ? UNKNOWN      "; $badgeColor = "DarkGray"   }
    }

    # Header row
    Write-Host "  " -NoNewline
    Write-Host "#$($Index.ToString().PadLeft(2))" -NoNewline -ForegroundColor DarkGray
    Write-Host " │ " -NoNewline -ForegroundColor DarkGray
    Write-Host $FriendlyName.PadRight(32) -NoNewline -ForegroundColor White
    Write-Host " │ " -NoNewline -ForegroundColor DarkGray
    Write-Host $badgeText -NoNewline -BackgroundColor DarkGray -ForegroundColor $badgeColor
    if ($DeviceType) {
        Write-Host "  $DeviceType" -ForegroundColor Gray
    } else {
        Write-Host ""
    }

    # Disconnected notice
    if ($status -eq "DISCONNECTED") {
        $sinceMsg = ""
        if ($DisconnectTime) {
            $span = (Get-Date) - $DisconnectTime
            if    ($span.TotalDays  -ge 1) { $sinceMsg = "  ($([math]::Floor($span.TotalDays))d ago)" }
            elseif($span.TotalHours -ge 1) { $sinceMsg = "  ($([math]::Floor($span.TotalHours))h ago)" }
            else                           { $sinceMsg = "  ($([math]::Floor($span.TotalMinutes))m ago)" }
        }
        Write-Host "       " -NoNewline
        Write-Host "⚠  Device is currently unplugged — last seen $dt$sinceMsg" -ForegroundColor DarkYellow
    }

    # Timestamps row — now matches USBDeview style
    Write-Host "       " -NoNewline
    Write-Host "Install: " -NoNewline -ForegroundColor DarkGray
    Write-Host $it -NoNewline -ForegroundColor Gray
    Write-Host "  ⊕ First: " -NoNewline -ForegroundColor DarkYellow
    Write-Host $fi -NoNewline -ForegroundColor Yellow
    Write-Host "  ▲ Connect: " -NoNewline -ForegroundColor DarkGreen
    Write-Host $ct -NoNewline -ForegroundColor Green
    Write-Host "  ▼ Disconnect: " -NoNewline -ForegroundColor DarkRed
    Write-Host $dt -ForegroundColor Red

    # Detail row
    $details = @()
    if ($Vid -or $UsbPid) { $details += "VID/PID: $Vid/$UsbPid" }
    if ($SerialNumber)    { $details += "S/N: $SerialNumber" }
    if ($InstanceId)      { $details += $InstanceId }
    if ($details.Count -gt 0) {
        Write-Host "       " -NoNewline
        Write-Host ($details -join "  ·  ") -ForegroundColor DarkGray
    }

    Write-Host "  " -NoNewline -ForegroundColor DarkGray
    Write-Host ("·" * 60) -ForegroundColor DarkGray
}

# ════════════════════════════════════════════════════════════════
#  MAIN SCAN
# ════════════════════════════════════════════════════════════════

Show-Splash
Write-Host "  Scanning..." -ForegroundColor DarkGray
Write-Host ""

$allDevices = @()

# Pre-load WMI for live connection status
$script:WmiPnpIds = @()
try {
    $script:WmiPnpIds = @(
        Get-WmiObject Win32_PnPEntity -ErrorAction SilentlyContinue |
        Where-Object { $_.DeviceID -match '^USB\\|^USBSTOR\\' } |
        Select-Object -ExpandProperty DeviceID
    )
} catch { }

# Step 1 — USBSTOR (Storage devices)
Show-LoadingBar -Percent 0 -Label "[ 1/3 ] USBSTOR"
$usbstorRoot = "HKLM:\SYSTEM\CurrentControlSet\Enum\USBSTOR"
if (Test-Path $usbstorRoot) {
    $classes = Get-ChildItem $usbstorRoot -ErrorAction SilentlyContinue
    $ci = 0; $ct = @($classes).Count
    foreach ($class in $classes) {
        $ci++
        Show-LoadingBar -Percent ([math]::Floor($ci/$([math]::Max($ct,1))*30)) -Label "[ 1/3 ] USBSTOR"
        $cn = $class.PSChildName
        $devType = "Storage"
        if ($cn -match 'CdRom') { $devType = "CD/DVD-ROM" }
        if ($cn -match 'Tape')  { $devType = "Tape Drive" }
        $vendor = ""; $product = ""
        if ($cn -match 'Ven_([^&]+)')  { $vendor  = $Matches[1] }
        if ($cn -match 'Prod_([^&]+)') { $product = $Matches[1] }
        foreach ($inst in (Get-ChildItem $class.PSPath -ErrorAction SilentlyContinue)) {
            $serial = $inst.PSChildName -replace '&\d+$',''
            $props  = Get-ItemProperty $inst.PSPath -ErrorAction SilentlyContinue
            $fname  = $props.FriendlyName
            if (-not $fname) { $fname = if ($vendor -and $product) { "$vendor $product" } else { $cn } }
            $times  = Get-DeviceTimestamps -DevKeyPath $inst.PSPath
            $iid    = "USBSTOR\$cn\$($inst.PSChildName)"
            $allDevices += [PSCustomObject]@{
                FriendlyName=$fname; DeviceType=$devType; Bus="USBSTOR"
                VID=""; PID=""; SerialNumber=$serial
                InstanceId=$iid; RegKeyPath=$inst.PSPath
                InstallTime=$times.InstallTime
                FirstInstallTime=$times.FirstInstallTime
                ConnectTime=$times.ConnectTime
                DisconnectTime=$times.DisconnectTime
                SortTime=if($times.ConnectTime){$times.ConnectTime}elseif($times.FirstInstallTime){$times.FirstInstallTime}else{[datetime]::MinValue}
            }
        }
    }
}

# Step 2 — USB (other devices)
Show-LoadingBar -Percent 33 -Label "[ 2/3 ] USB Devices"
$usbRoot = "HKLM:\SYSTEM\CurrentControlSet\Enum\USB"
if (Test-Path $usbRoot) {
    $vpKeys = Get-ChildItem $usbRoot -ErrorAction SilentlyContinue
    $vi = 0; $vt = @($vpKeys).Count
    foreach ($vpKey in $vpKeys) {
        $vi++
        Show-LoadingBar -Percent (33+[math]::Floor($vi/$([math]::Max($vt,1))*33)) -Label "[ 2/3 ] USB Devices"
        $vpn    = $vpKey.PSChildName
        $vid    = if ($vpn -match 'VID_([0-9A-Fa-f]+)') { $Matches[1] } else { "" }
        $usbPid = if ($vpn -match 'PID_([0-9A-Fa-f]+)') { $Matches[1] } else { "" }
        foreach ($inst in (Get-ChildItem $vpKey.PSPath -ErrorAction SilentlyContinue)) {
            $props    = Get-ItemProperty $inst.PSPath -ErrorAction SilentlyContinue
            $fname    = $props.FriendlyName; if (-not $fname) { $fname = $vpn }
            $devClass = $props.Class;  $service = $props.Service
            $devType  = if ($devClass) { $devClass } else { "USB Device" }
            if ($service) { $devType = "$devType ($service)".Trim(" ()") }
            $times = Get-DeviceTimestamps -DevKeyPath $inst.PSPath
            $iid   = "USB\$vpn\$($inst.PSChildName)"
            $allDevices += [PSCustomObject]@{
                FriendlyName=$fname; DeviceType=$devType; Bus="USB"
                VID=$vid; PID=$usbPid; SerialNumber=$inst.PSChildName
                InstanceId=$iid; RegKeyPath=$inst.PSPath
                InstallTime=$times.InstallTime
                FirstInstallTime=$times.FirstInstallTime
                ConnectTime=$times.ConnectTime
                DisconnectTime=$times.DisconnectTime
                SortTime=if($times.ConnectTime){$times.ConnectTime}elseif($times.FirstInstallTime){$times.FirstInstallTime}else{[datetime]::MinValue}
            }
        }
    }
}

# Step 3 — pnputil forced removals (last 90 days)
Show-LoadingBar -Percent 66 -Label "[ 3/3 ] pnputil log"
$pnpResults = @()
try {
    $evts = Get-WinEvent -FilterHashtable @{ LogName='Security'; Id=4688; StartTime=(Get-Date).AddDays(-90) } -ErrorAction SilentlyContinue |
            Where-Object { $_.Message -match 'pnputil' }
    if ($evts) {
        foreach ($ev in ($evts | Sort-Object TimeCreated -Descending)) {
            $cmd = ""; if ($ev.Message -match 'Process Command Line:\s+(.+)') { $cmd = $Matches[1].Trim() }
            if ($cmd -notmatch '/remove-device|/remove') { continue }
            $iid = ""
            if ($cmd -match '"([^"]+)"') { $iid = $Matches[1] }
            elseif ($cmd -match '/remove-device\s+(\S+)') { $iid = $Matches[1] }
            $pnpResults += [PSCustomObject]@{ Time=$ev.TimeCreated; InstanceId=$iid; Command=$cmd }
        }
    }
} catch { }

Show-LoadingBar -Percent 100 -Label "[ ✓  ] Done"
Start-Sleep -Milliseconds 300

# ════════════════════════════════════════════════════════════════
#  DISPLAY RESULTS
# ════════════════════════════════════════════════════════════════

Clear-Host
Show-Splash

# Summary
$totalConnected = 0; $totalDisconnected = 0
foreach ($d in $allDevices) {
    $s = Get-ConnectionStatus -LastConnected $d.ConnectTime -LastDisconnected $d.DisconnectTime -InstanceId $d.InstanceId -RegKeyPath $d.RegKeyPath
    if ($s -eq "CONNECTED")    { $totalConnected++ }
    if ($s -eq "DISCONNECTED") { $totalDisconnected++ }
}

Write-Host "  STATUS   " -NoNewline -ForegroundColor DarkGray
Write-Host "● $totalConnected connected" -NoNewline -ForegroundColor Green
Write-Host "   ○ $totalDisconnected unplugged" -ForegroundColor DarkYellow
Write-Host ""
Write-Host "  LEGEND   " -NoNewline -ForegroundColor DarkGray
Write-Host "Install / ⊕ First Install  ▲ Connect  ▼ Disconnect" -ForegroundColor Gray
Write-Host ""

# Storage Devices
$stor = $allDevices | Where-Object { $_.Bus -eq "USBSTOR" } | Sort-Object SortTime -Descending
Write-SectionHeader -Title " STORAGE DEVICES  ($(@($stor).Count) found) " -Color Cyan
if (@($stor).Count -eq 0) {
    Write-Host "  (none)" -ForegroundColor DarkGray
} else {
    $i = 1
    foreach ($d in $stor) {
        Write-DeviceRow -FriendlyName $d.FriendlyName -DeviceType $d.DeviceType `
            -InstanceId $d.InstanceId -RegKeyPath $d.RegKeyPath `
            -Vid $d.VID -UsbPid $d.PID -SerialNumber $d.SerialNumber `
            -InstallTime $d.InstallTime -FirstInstallTime $d.FirstInstallTime `
            -ConnectTime $d.ConnectTime -DisconnectTime $d.DisconnectTime -Index $i
        $i++
    }
}

Write-Host ""

# Other USB Devices
$other = $allDevices | Where-Object { $_.Bus -eq "USB" } | Sort-Object SortTime -Descending
Write-SectionHeader -Title " OTHER USB DEVICES  ($(@($other).Count) found) " -Color DarkCyan
if (@($other).Count -eq 0) {
    Write-Host "  (none)" -ForegroundColor DarkGray
} else {
    $i = 1
    foreach ($d in $other) {
        Write-DeviceRow -FriendlyName $d.FriendlyName -DeviceType $d.DeviceType `
            -InstanceId $d.InstanceId -RegKeyPath $d.RegKeyPath `
            -Vid $d.VID -UsbPid $d.PID -SerialNumber $d.SerialNumber `
            -InstallTime $d.InstallTime -FirstInstallTime $d.FirstInstallTime `
            -ConnectTime $d.ConnectTime -DisconnectTime $d.DisconnectTime -Index $i
        $i++
    }
}

Write-Host ""

# Forced Removals
Write-SectionHeader -Title " FORCED REMOVALS  (pnputil / last 90 days) " -Color Magenta
if ($pnpResults.Count -eq 0) {
    Write-Host "  (none found — enable Audit Process Creation in secpol.msc)" -ForegroundColor DarkGray
} else {
    foreach ($r in $pnpResults) {
        Write-Host "  " -NoNewline
        Write-Host $r.Time.ToString('yyyy-MM-dd HH:mm:ss') -NoNewline -ForegroundColor Cyan
        Write-Host "  " -NoNewline
        Write-Host $r.InstanceId -NoNewline -ForegroundColor White
        Write-Host "  ·  " -NoNewline -ForegroundColor DarkGray
        Write-Host $r.Command -ForegroundColor Gray
    }
}

Write-Host ""
$total = $allDevices.Count
Write-Host "  ── " -NoNewline -ForegroundColor DarkGray
Write-Host "SCAN COMPLETE" -NoNewline -ForegroundColor Green
Write-Host "  ·  " -NoNewline -ForegroundColor DarkGray
Write-Host "$total device(s)" -NoNewline -ForegroundColor White
Write-Host "  ·  " -NoNewline -ForegroundColor DarkGray
Write-Host (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') -NoNewline -ForegroundColor Cyan
Write-Host "  ──" -ForegroundColor DarkGray
Write-Host ""