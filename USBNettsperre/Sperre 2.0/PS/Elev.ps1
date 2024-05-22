$ProgressPreference = 'SilentlyContinue'

Function Get-ConnectedWifi {
    $wifiInfo = netsh wlan show interfaces
    $connectedWifi = ($wifiInfo -match '^\s*SSID\s*:')[0] -replace '^\s*SSID\s*:\s*', ''
    return $connectedWifi
}

function Get-LastRebootTime {
    $os = Get-WmiObject -Class Win32_OperatingSystem
    $lastBootTime = $os.ConvertToDateTime($os.LastBootUpTime)
    $elapsedTime = [DateTime]::Now - $lastBootTime
    $days = [math]::floor($elapsedTime.TotalDays)
    $hours = [math]::floor($elapsedTime.TotalHours) % 24
    $minutes = [math]::floor($elapsedTime.TotalMinutes) % 60
    return " $days dager $hours timer og $minutes minutter siden siste omstart."
}

$connectedWifi = Get-ConnectedWifi

if ($connectedWifi -ne "VTFK") {
    Write-Output "Du er ikke koblet til VTFK-nettverket."
}

$msgTitle = "Wi-Fi-nettverksstatus"
$msgBody = "Koblet til Wi-Fi-nettverk: $connectedWifi`n"
$lastRebootTime = Get-LastRebootTime
$msgBody += "Sist startet på nytt: $lastRebootTime`n"

$ipAddress = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq 'IPv4' -and $_.InterfaceAlias -eq 'Wi-Fi' }).IPAddress
$allowedIPRanges = @("10.148.96")
$allowedIP = $false

foreach ($range in $allowedIPRanges) {
    if ($ipAddress.StartsWith($range)) {
        $allowedIP = $true
        break
    }
}

if (-not $allowedIP) {
    $advarselMsg = "ADVARSEL: SJEKK IP-ADRESSE!"
    $msgBody += "`n$advarselMsg`n"
}

$url_eksamen = "https://eksamen.vtfk.no/elev"
$url_vg = "http://vg.no/"

# Sett sikkerhetsprotokollen til Tls12
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

try {
    $response_eksamen = Invoke-WebRequest -Uri $url_eksamen -UseBasicParsing -TimeoutSec 5
    $eksamen_status = $true
}
catch {
    $eksamen_status = $false
    $errorMsg = "Feil ved tilgang til eksamen.vtfk.no: $($_.Exception.Message)"
    $msgBody += $errorMsg + "`n"
}

try {
    $response_vg = Invoke-WebRequest -Uri $url_vg -UseBasicParsing -TimeoutSec 2
    $vg_status = $true
}
catch {
    $vg_status = $false
    $errorMsg = "Feil ved tilgang til vg.no: $($_.Exception.Message)"
    $msgBody += $errorMsg + "`n"
}

$logName = "Microsoft-Windows-WLAN-AutoConfig/Operational"
$eventID = 8001
$startTime = (Get-Date).AddDays(-1)
$filter = @{
    LogName = $logName
    ID = $eventID
    StartTime = $startTime
}
$wifiEvents = Get-WinEvent -FilterHashtable $filter

$vlanInfo = "VLAN Oversikt`n`n"
$vlanInfo += "---------------------------------`n"

foreach ($event in $wifiEvents) {
    $eventXml = [xml]$event.ToXml()
    $ssid = $eventXml.Event.EventData.Data | Where-Object { $_.Name -eq "SSID" } | Select-Object -ExpandProperty "#text"
    $connectionDate = $event.TimeCreated
    $vlanInfo += "SSID: $ssid`nTilkoblingsdato: $connectionDate`n"
    $vlanInfo += "---------------------------------`n"
}

$msgBody += $vlanInfo

if ($eksamen_status -and -not $vg_status) {
    $msgBody += "Eleven er sperret."
}
elseif ($eksamen_status -and $vg_status) {
    $msgBody += "Eleven er IKKE Sperret og har nettverkstilgang."
}
else {
    $msgBody += "Ukjent feil, sjekk logg."
}

# Legg til IP-adressen i meldingen
$msgBody += "`nIP-adresse: $ipAddress"

$null = [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Windows.Forms.MessageBox]::Show($msgBody, $msgTitle)

$logFolder = "log"
if (-not (Test-Path $logFolder)) {
    New-Item -ItemType Directory -Path $logFolder | Out-Null
}

$serialNumber = (Get-WmiObject -Class Win32_BIOS).SerialNumber
$userName = $env:USERNAME
$fileName = "$logFolder\$userName-$serialNumber-SSID-oversikt.txt"
$logContent = "IP-adresse: $ipAddress`n`n" + $vlanInfo
$logContent | Out-File -FilePath $fileName

Write-Output "SSID-oversikten er eksportert til $fileName"
