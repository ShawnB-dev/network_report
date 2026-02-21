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

- reports\NetworkReport.html - Visual dashboard suitable for screenshots and documentation.

- reports\NetworkReport.xlsx = Multi-sheet Excel workbook, including system info, network info, DNS, latency, ARP, routing, firewall and SNMP

Minimal Report

```powershell
New-NetworkReport

By default, the report includes system and network information. Additional checks can be enabled via switches.

NOTE: Excel export uses the ImportExcel PowerShell module. 

if ImportExcel is not installed, the HTML report is still generated.

## OUI Vendor Lookup

The module includes multiple OUI datasets that allow Invoke-ArpAnalysis to identify device manufacturers based on MAC address prefixes. 

Available vendor sets:

- Enterprise - Cisco, Juniper, HPE, Aruba, Dell, Fortinet, Palo Alto, etc.

- Japan - NEC, Fujitsu, Sharp, Toshiba, Sony, Panasonic, Buffalo, Yamaha, etc.

- Wireless - Ubiquiti, TP-Link, ASUS, Apple, Samsung, Google, Amazon, etc.

Analyze ARP Table using the Enterprise vendor set

```powershell
Invoke-ArpAnalysis -VendorSet Enterprise

Analyze ARP Table using the Japan-focused vendor set

```powershell
Invoke-ArpAnalysis 

Analyze ARP Table using the Wireless vendor set

```powershell
Invoke-ArpAnalyzis -VendorSet Wireless

Resolve hostnames during ARP Analysis

```powershell
Invoke-ArpAnalysis -VendorSet Enterprise -ResolveHostnames

## OUI Data Files

The OUI datasets are stored under:

NetworkTools/data/oui/
│   oui_enterprise.csv
│   oui_japan.csv
│   oui_wireless.csv

Each file contains a Prefix and Vendor pairing.

## Module Structure

NetworkTools/
│   NetworkTools.psd1
│   NetworkTools.psm1
│
├── functions/
│     ExportDhcpLeases.ps1
│     Invoke-SubnetScan.ps1
│     Invoke-DnsHealthCheck.ps1
│     Invoke-LatencyJitterMonitor.ps1
│     Invoke-RoutingTableAnalysis.ps1
│     Invoke-FirewallRuleAudit.ps1
│     Invoke-ArpAnalysis.ps1
│     Invoke-SnmpSwitchMapper.ps1
│     New-NetworkChangeLog.ps1
│     New-NetworkReport.ps1
│
├── data/
│     └── oui/
│          oui_enterprise.csv
│          oui_japan.csv
│          oui_wireless.csv
│
└── docs/
      README_EN.md
      README_JP.md
      LICENSE

## Purpose

This module was created to demonstrate practical PowerShell skills used in real network administration:

* Modular function design
* Clean, maintainable code
* Automation-friendly output
* Documentation aligned with industry standards
* Realistic network diagnostic workflows
* Excel-friendly reporting for Japanese enterprise environments

## License

This project is licensed under the MIT License. 
See the LICENSE file for full details. 