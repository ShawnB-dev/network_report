function New-NetworkReport {
    [CmdletBinding()]
    param(
        [string]$OutputPath = (Join-Path $PSScriptRoot "..\reports"),
        [switch]$IncludeArp,
        [switch]$IncludeDns,
        [switch]$IncludeLatency,
        [switch]$IncludeRouting,
        [switch]$IncludeFirewall,
        [switch]$IncludeSnmp,
        [string[]]$DnsServers = @("8.8.8.8","1.1.1.1"),
        [string]$LatencyTarget = "8.8.8.8"
    )

    # Ensure reports folder exists
    if (-not (Test-Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory | Out-Null
    }

    $timestamp = Get-Date
    $htmlPath  = Join-Path $OutputPath "NetworkReport.html"
    $xlsxPath  = Join-Path $OutputPath "NetworkReport.xlsx"

    # --- Collect data ---

    $systemInfo = [pscustomobject]@{
        ComputerName = $env:COMPUTERNAME
        UserName     = $env:USERNAME
        OSVersion    = (Get-CimInstance Win32_OperatingSystem).Caption
        Timestamp    = $timestamp
    }

    $networkInfo = Get-NetIPConfiguration | Select-Object InterfaceAlias,IPv4Address,IPv4DefaultGateway,DnsServer

    $dnsResult = $null
    if ($IncludeDns) {
        try {
            $dnsResult = Invoke-DnsHealthCheck -DnsServers $DnsServers -ErrorAction Stop
        } catch {
            $dnsResult = @([pscustomobject]@{ Server="Error"; Status=$_.Exception.Message })
        }
    }

    $latencyResult = $null
    if ($IncludeLatency) {
        try {
            $latencyResult = Invoke-LatencyJitterMonitor -Target $LatencyTarget -ErrorAction Stop
        } catch {
            $latencyResult = @([pscustomobject]@{ Target=$LatencyTarget; Status="Error"; Note=$_.Exception.Message })
        }
    }

    $arpResult = $null
    if ($IncludeArp) {
        try {
            $arpResult = Invoke-ArpAnalysis -VendorSet Enterprise -ErrorAction Stop
        } catch {
            $arpResult = @([pscustomobject]@{ IPAddress=""; MACAddress=""; Vendor="Error"; Note=$_.Exception.Message })
        }
    }

    $routingResult = $null
    if ($IncludeRouting) {
        try {
            $routingResult = Invoke-RoutingTableAnalysis -ErrorAction Stop
        } catch {
            $routingResult = @([pscustomobject]@{ Destination="Error"; Note=$_.Exception.Message })
        }
    }

    $firewallResult = $null
    if ($IncludeFirewall) {
        try {
            $firewallResult = Invoke-FirewallRuleAudit -ErrorAction Stop
        } catch {
            $firewallResult = @([pscustomobject]@{ Name="Error"; Note=$_.Exception.Message })
        }
    }

    $snmpResult = $null
    if ($IncludeSnmp) {
        try {
            $snmpResult = Invoke-SnmpSwitchMapper -ErrorAction Stop
        } catch {
            $snmpResult = @([pscustomobject]@{ Switch="Error"; Note=$_.Exception.Message })
        }
    }

    # --- Build simple status summary ---

    $status = [pscustomobject]@{
        DnsStatus      = if ($dnsResult)      { "OK" } else { "Skipped" }
        LatencyStatus  = if ($latencyResult)  { "OK" } else { "Skipped" }
        ArpStatus      = if ($arpResult)      { "OK" } else { "Skipped" }
        RoutingStatus  = if ($routingResult)  { "OK" } else { "Skipped" }
        FirewallStatus = if ($firewallResult) { "OK" } else { "Skipped" }
        SnmpStatus     = if ($snmpResult)     { "OK" } else { "Skipped" }
    }

    # --- HTML report ---

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Network Report - $($systemInfo.ComputerName)</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; }
h1, h2 { color: #333; }
table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
th, td { border: 1px solid #ccc; padding: 6px 8px; font-size: 12px; }
th { background-color: #f2f2f2; text-align: left; }
.status-ok { color: #0a0; font-weight: bold; }
.status-skip { color: #999; }
.footer { font-size: 11px; color: #666; margin-top: 20px; }
</style>
</head>
<body>
<h1>Network Health Report</h1>
<p>Generated: $($systemInfo.Timestamp)</p>
<h2>System Info</h2>
<table>
<tr><th>Computer Name</th><td>$($systemInfo.ComputerName)</td></tr>
<tr><th>User</th><td>$($systemInfo.UserName)</td></tr>
<tr><th>OS</th><td>$($systemInfo.OSVersion)</td></tr>
</table>

<h2>Status Summary</h2>
<table>
<tr><th>Check</th><th>Status</th></tr>
<tr><td>DNS</td><td class="status-$(if($status.DnsStatus -eq 'OK'){'ok'}else{'skip'})">$($status.DnsStatus)</td></tr>
<tr><td>Latency</td><td class="status-$(if($status.LatencyStatus -eq 'OK'){'ok'}else{'skip'})">$($status.LatencyStatus)</td></tr>
<tr><td>ARP</td><td class="status-$(if($status.ArpStatus -eq 'OK'){'ok'}else{'skip'})">$($status.ArpStatus)</td></tr>
<tr><td>Routing</td><td class="status-$(if($status.RoutingStatus -eq 'OK'){'ok'}else{'skip'})">$($status.RoutingStatus)</td></tr>
<tr><td>Firewall</td><td class="status-$(if($status.FirewallStatus -eq 'OK'){'ok'}else{'skip'})">$($status.FirewallStatus)</td></tr>
<tr><td>SNMP</td><td class="status-$(if($status.SnmpStatus -eq 'OK'){'ok'}else{'skip'})">$($status.SnmpStatus)</td></tr>
</table>

<h2>Network Interfaces</h2>
$($networkInfo | ConvertTo-Html -Fragment)

@if($IncludeDns){
<h2>DNS Health</h2>
$($dnsResult | ConvertTo-Html -Fragment)
}

@if($IncludeLatency){
<h2>Latency / Jitter</h2>
$($latencyResult | ConvertTo-Html -Fragment)
}

@if($IncludeArp){
<h2>ARP Table (with Vendor)</h2>
$($arpResult | ConvertTo-Html -Fragment)
}

@if($IncludeRouting){
<h2>Routing Table</h2>
$($routingResult | ConvertTo-Html -Fragment)
}

@if($IncludeFirewall){
<h2>Firewall Rules</h2>
$($firewallResult | ConvertTo-Html -Fragment)
}

@if($IncludeSnmp){
<h2>SNMP Switch Mapping</h2>
$($snmpResult | ConvertTo-Html -Fragment)
}

<div class="footer">
Generated by NetworkTools :: New-NetworkReport<br/>
© 2026 Shawn — MIT License
</div>
</body>
</html>
"@

    # Replace @if blocks manually (simple approach)
    $html = $html -replace '@if\(\$IncludeDns\)\{', ''
    $html = $html -replace '@if\(\$IncludeLatency\)\{', ''
    $html = $html -replace '@if\(\$IncludeArp\)\{', ''
    $html = $html -replace '@if\(\$IncludeRouting\)\{', ''
    $html = $html -replace '@if\(\$IncludeFirewall\)\{', ''
    $html = $html -replace '@if\(\$IncludeSnmp\)\{', ''
    $html = $html -replace '\}', ''

    $html | Set-Content -Path $htmlPath -Encoding UTF8

    # --- Excel export (requires ImportExcel module) ---

    try {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            Write-Verbose "ImportExcel not found. Skipping Excel export."
        }
        else {
            Remove-Item -Path $xlsxPath -ErrorAction SilentlyContinue

            $systemInfo    | Export-Excel -Path $xlsxPath -WorksheetName 'SystemInfo' -AutoSize
            $networkInfo   | Export-Excel -Path $xlsxPath -WorksheetName 'NetworkInfo' -AutoSize -Append

            if ($dnsResult)      { $dnsResult      | Export-Excel -Path $xlsxPath -WorksheetName 'DNS'       -AutoSize -Append }
            if ($latencyResult)  { $latencyResult  | Export-Excel -Path $xlsxPath -WorksheetName 'Latency'   -AutoSize -Append }
            if ($arpResult)      { $arpResult      | Export-Excel -Path $xlsxPath -WorksheetName 'ARP'       -AutoSize -Append }
            if ($routingResult)  { $routingResult  | Export-Excel -Path $xlsxPath -WorksheetName 'Routing'   -AutoSize -Append }
            if ($firewallResult) { $firewallResult | Export-Excel -Path $xlsxPath -WorksheetName 'Firewall'  -AutoSize -Append }
            if ($snmpResult)     { $snmpResult     | Export-Excel -Path $xlsxPath -WorksheetName 'SwitchPorts' -AutoSize -Append }
        }
    }
    catch {
        Write-Warning "Excel export failed: $($_.Exception.Message)"
    }

    [pscustomobject]@{
        HtmlReportPath = $htmlPath
        ExcelReportPath = if (Test-Path $xlsxPath) { $xlsxPath } else { $null }
        Timestamp = $timestamp
        ComputerName = $systemInfo.ComputerName
    }
}