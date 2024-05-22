$ProgressPreference = 'SilentlyContinue'

Function Get-ConnectedWifi {
    $wifiInfo = netsh wlan show interfaces
    $connectedWifi = ($wifiInfo -match '^\s*SSID\s*:')[0] -replace '^\s*SSID\s*:\s*', ''
    return $connectedWifi
}

function Get-WifiDetails {
    $wifiInfo = netsh wlan show interfaces
    $details = @{
        SSID = ($wifiInfo -match '^\s*SSID\s*:')[0] -replace '^\s*SSID\s*:\s*', ''
        Signal = ($wifiInfo -match '^\s*Signal\s*:')[0] -replace '^\s*Signal\s*:\s*', ''
        Speed = ($wifiInfo -match '^\s*Receive rate\s*:')[0] -replace '^\s*Receive rate\s*:\s*', ''
    }
    return $details
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

$wifiDetails = Get-WifiDetails

if ($wifiDetails.SSID -ne "VTFK") {
    Write-Output "Du er ikke koblet til VTFK-nettverket."
}

$msgTitle = "Wi-Fi-nettverksstatus"
$msgBody = "==== Wi-Fi Nettverksstatus ====`n"
$msgBody += "SSID: $($wifiDetails.SSID)`n"
$msgBody += "Signalstyrke: $($wifiDetails.Signal)`n"
$msgBody += "Hastighet: $($wifiDetails.Speed)`n"
$msgBody += "`n==== Systeminformasjon ====`n"
$lastRebootTime = Get-LastRebootTime
$msgBody += "Sist startet p  nytt: $lastRebootTime`n"
$msgBody += "Rapport generert: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')`n"

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

$vlanInfo = "`n==== VLAN Oversikt ====`n"
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
    $msgBody += "`n==== Status ====`nEleven er sperret."
}
elseif ($eksamen_status -and $vg_status) {
    $msgBody += "`n==== Status ====`nEleven er IKKE Sperret og har nettverkstilgang."
}
else {
    $msgBody += "`n==== Status ====`nUkjent feil, sjekk logg."
}

# VLAN Mapping
$vlanMapping = @{
    96  = "Eksamen"
    188 = "IT VLAN"
    120 = "VLAN For nettsperre med Teams"
    164 = "Ansatt Wi-Fi"
    112 = "Eksameninett (Nett med sperringer)"
}

# Extract the third octet of the IP address to determine the VLAN
$thirdOctet = [int]($ipAddress.Split('.')[2])

# Check if the third octet matches any known VLANs
if ($vlanMapping.ContainsKey($thirdOctet)) {
    $vlanName = $vlanMapping[$thirdOctet]
    $msgBody += "`nBrukeren har f lgende VLAN: $vlanName"
} else {
    $msgBody += "`nVLAN er ukjent."
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
$fileName = "$logFolder\\$userName-$serialNumber-SSID-oversikt.txt"
$logContent = "IP-adresse: $ipAddress`n`n" + $vlanInfo
$logContent | Out-File -FilePath $fileName

Write-Output "SSID-oversikten er eksportert til $fileName"
