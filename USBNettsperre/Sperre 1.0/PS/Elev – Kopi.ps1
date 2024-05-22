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
    $response_vg = Invoke-WebRequest -Uri $url_vg -UseBasicParsing -TimeoutSec 5
    $vg_status = $true
}
catch {
    $vg_status = $false
}

$GetEventLogs = Get-EventLog -LogName System -Newest 500 | Where-Object { $_.EventID -eq 10001 }
$vlan_oversikt = @()
foreach ($Event in $GetEventLogs) {
    if ($Event.Message -match 'SSID: ([^\r\n]+)') {
        $SSID = $matches[1]
        $date = $Event.TimeGenerated
        $vlan_oversikt += "Dato: $date | SSID: $SSID"
    }
}

if ($eksamen_status) {
    $msgBody += "Du har tilgang til eksamen.vtfk.no!"
} else {
    $msgBody += "Du har ikke tilgang til eksamen.vtfk.no!"
}

$elev_status = if ($eksamen_status) { "OK" } else { "SPERRET" }
$msgBody += "`nDin status: $elev_status"
$msgBody += "`nIP-adresse: $ipAddress"

[System.Windows.Forms.MessageBox]::Show($msgBody, $msgTitle, 0, 48)

$date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$logEntry = "$date,$env:USERNAME,$env:COMPUTERNAME,$ipAddress,$vlan_oversikt"

$logFilePath = "C:\temp\nettverkslogg.csv"

if (-not (Test-Path $logFilePath)) {
    $header = "Dato,Brukernavn,Maskinnavn,IP-adresse,VLAN-oversikt"
    Add-Content -Path $logFilePath -Value $header
}

Add-Content -Path $logFilePath -Value $logEntry
