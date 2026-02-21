# NetworkTools PowerShell Module

A professional, modular PowerShell toolkit designed for network diagnostics, auditing, and inventory tasks.  
This module provides fast, scriptable, and automation‑friendly functions for real‑world network administration.

![PowerShell](https://img.shields.io/badge/PowerShell-Module-blue)
![License: MIT](https://img.shields.io/badge/License-MIT-green)
![Status](https://img.shields.io/badge/Version-0.1.0-orange)

---

## Features

- Subnet scanning with optional hostname resolution  
- DNS server health checks  
- DHCP lease export (CSV/JSON)  
- Latency and jitter monitoring  
- Routing table analysis  
- Windows Firewall rule auditing  
- ARP table analysis with optional vendor lookup  
- SNMP‑based switch port and MAC mapping  
- Network change logging for documentation and compliance  
- **Network Health Dashboard (HTML + Excel)**  
- **Excel‑friendly reporting for Japanese enterprise workflows**

All functions include:

- Comment‑based help  
- Parameter validation  
- Consistent output objects  
- Clean, predictable behavior  

---

## Installation

Clone the repository:

```powershell
git clone https://github.com/<your-username>/network_report.git

Import the module:

```powershell
Import-Module .\NetworkTools -Force

Verify available functions:

```powershell
Get-Command -Module NetworkTools

## Usage Examples

Scan a Subnet

```powershell
Invoke-SubnetScan -Subnet "192.168.1"

Check DNS Server Health

```powershell
Invoke-DnsHealthCheck -DnsServers "8.8.8.8", "1.1.1.1"

Monitor Latency and Jitter

```powershell
Invoke-LatencyJitterMonitor -Target "8.8.8.8"

Audit firewall rules

```powershell
Invoke-FirewallRuleAudit

Log a Network Change

```powershell
New-NetworkChangeLong -ChangeDescription "Updated VLAN assisngments"

Network Health Dashboard and Excel Report

The module includes a reporting function that generates both an HTML Dashboard and and Excel workbook that summarizes key network diagnostics.

Generate a full network report

```powershell
New-NetworkReport -IncludeArp -IncludeDns -IncludeLatency -IncludeRouting -IncludeFirewall -IncludeSnmp

This creates:

- **reports\NetworkReport.html** - Visual dashboard suitable for screenshots and documentation.

- **reports\NetworkReport.xlsx** = Multi-sheet Excel workbook, including system info, network info, DNS, latency, ARP, routing, firewall and SNMP

