# 🛠️ QuickFIX

> *Because nobody has time to click through 15 Windows menus just to check a MAC address or see if a RAM stick died.*

Built in the IT trenches of a busy clinic, **QuickFIX** is a no-nonsense, native PowerShell panel designed to diagnose and troubleshoot daily Windows headaches. 

No bloated third-party software. No unnecessary GUIs. Just pure, safe `Get-CimInstance` queries and a terminal to get the job done fast.

## 📦 What's in the box (so far)?

* **Deep System Info (`Show-SystemInfo`):** Spits out the actual hardware specs without the Windows fluff. CPU cache, exact RAM module locators, Motherboard/BIOS details, storage partitions, and GPU status. 
* **Network Diagnostics (`Show-NetworkInfo`):** Grabs your active Ethernet adapters, pulls the IPs, Subnets, and DNS, and automatically pings the default gateway so you don't have to open a second terminal window.
* **The "Hold Up" Screen (`Rollback-Welcome`):** A mandatory prompt reminding you to create a System Restore point before messing with the machine. *(Note: The actual creation logic is currently a mockup, but the guilt-trip is real).*
* **Built-in Receipts (`Write-Log`):** Logs all queries and any WMI/CIM errors silently to `QuickFix.log` in the same directory. If something breaks, you'll know why.

## ⚙️ Requirements

* Windows OS (Tested on modern builds).
* PowerShell 5.1+
* **Run as Administrator:** Seriously, just run it as Admin. Half the hardware queries will give you blank stares if you don't.

## 🚀 How to fire it up

## 🚀 Quick Execution (Support Mode)

If you are on a machine behind a corporate proxy (like in our clinic), run this command in an **Administrator PowerShell** to launch QuickFIX directly from the cloud:

```powershell
$wc = New-Object Net.WebClient; $wc.Proxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials; IEX $wc.DownloadString('[https://raw.githubusercontent.com/emanoeI/QuickFIX/refs/heads/main/QuickFIX.ps1](https://raw.githubusercontent.com/emanoeI/QuickFIX/refs/heads/main/QuickFIX.ps1)')
