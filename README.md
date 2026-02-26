<h1 align="center">🛠 QuickFix</h1>
<h3 align="center">Corporate Support & Diagnostic Script</h3>
<p align="center"><strong>coded by emanoel peres :)</strong></p>

<p align="center">
  <img src="https://img.shields.io/badge/PowerShell-5.0+-5391FE?logo=powershell&logoColor=white" />
  <img src="https://img.shields.io/badge/Windows-10%2B-0078D6?logo=windows&logoColor=white" />
  <img src="https://img.shields.io/badge/Status-Production-success" />
  <img src="https://img.shields.io/badge/Scope-Internal-blue" />
  <img src="https://img.shields.io/badge/License-Internal-red" />
</p>

<hr/>

<h2>📌 Purpose</h2>

<p>
<strong>QuickFix</strong> is a corporate PowerShell-based tool designed to standardize IT support procedures,
reduce ticket resolution time, and ensure full traceability of technical interventions performed within Organization.
</p>

<p>
The script runs natively on Windows systems and requires no installation.
</p>

<h3>Core Principles</h3>
<ul>
  <li>Operational standardization</li>
  <li>Minimal system intervention</li>
  <li>Full action traceability</li>
  <li>Modular and isolated execution</li>
</ul>

<p>No action with potential impact executes without explicit technician confirmation.</p>

<hr/>

<h2> Technical Architecture</h2>
<p>
QuickFix consists of <strong>6 independent modules</strong>, accessible via a numeric menu interface.
Each module runs in isolation to reduce operational risk.
</p>

<h3>Execution Flow</h3>
<ol>
  <li>Technician identification at session start</li>
  <li>Automatic registration in generated reports</li>
  <li>Module execution based on ticket requirements</li>
  <li>Automatic report generation after each action</li>
</ol>

<hr/>

<h2>🔎 Operational Modules</h2>
<h3>1️⃣ Hardware Diagnostics</h3>
<ul>
  <li>Collects OS, CPU, motherboard, RAM, video controllers, and storage info (read-only)</li>
  <li>Detects SSD/HDD and free disk space, with alert thresholds</li>
</ul>

<h3>2️⃣ Network Status</h3>
<ul>
  <li>Lists active network adapters, IP/DNS, gateway connectivity</li>
  <li>Performs DNS resolution tests</li>
</ul>

<h3>3️⃣ Printer Diagnostics & Repair</h3>
<ul>
  <li>Status and queue inspection</li>
  <li>Port and IP mapping</li>
  <li>Connectivity tests and spooler operations</li>
  <li>Default printer settings, test page generation, force online</li>
</ul>

<h3>4️⃣ RAM Optimization</h3>
<ul>
  <li>Global memory cleanup and selective high-consumption process reduction</li>
  <li>Permanent Chrome memory-saving flags</li>
</ul>

<h3>5️⃣ Network Repair</h3>
<ul>
  <li>Full network stack repair, DNS flush, connectivity tests, adapter reset</li>
</ul>

<h3>6️⃣ Windows System Repair</h3>
<ul>
  <li>DISM RestoreHealth, SFC /Scannow, StartComponentCleanup</li>
  <li>Requires reboot; does not delete user data</li>
</ul>

<hr/>

<h2>📄 Reporting System</h2>
<p>
Each action generates a <code>.txt</code> report in:
<pre>C:\services\relatorios</pre>
</p>
<ul>
  <li>Date and time</li>
  <li>Technician name</li>
  <li>Machine name</li>
  <li>IP address</li>
  <li>Executed module</li>
  <li>Operation result</li>
</ul>

<hr/>

<h2>▶ Execution Options</h2>
<p>Choose the method that fits your network and security policies:</p>

<h3>1️⃣ Direct in-memory (proxy-friendly)</h3>
<pre><code>$wc = New-Object Net.WebClient
$wc.Proxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
IEX $wc.DownloadString('https://raw.githubusercontent.com/emanoeI/QuickFIX/main/QuickFIX.ps1')
</code></pre>

<h3>2️⃣ Download and execute locally (safer, auditable)</h3>
<pre><code>$wc = New-Object Net.WebClient
$wc.Proxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
$wc.DownloadFile('https://raw.githubusercontent.com/emanoeI/QuickFIX/main/QuickFIX.ps1', 'QuickFIX.ps1')
powershell -ExecutionPolicy Bypass -File .\QuickFIX.ps1
</code></pre>

<h3>3️⃣ Short method via Invoke-RestMethod</h3>
<pre><code>powershell -ExecutionPolicy Bypass -Command "iex (irm https://raw.githubusercontent.com/emanoeI/QuickFIX/main/QuickFIX.ps1 -Proxy $null)"
</code></pre>

<h3>4️⃣ Download with SHA256 verification (enterprise-safe)</h3>
<pre><code>$wc = New-Object Net.WebClient
$wc.Proxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials
$wc.DownloadFile('https://raw.githubusercontent.com/emanoeI/QuickFIX/main/QuickFIX.ps1', 'QuickFIX.ps1')
$hash = Get-FileHash .\QuickFIX.ps1 -Algorithm SHA256
if ($hash.Hash -eq 'INSERT_OFFICIAL_SHA256') {
    powershell -ExecutionPolicy Bypass -File .\QuickFIX.ps1
} else {
    Write-Host "ERROR: File hash mismatch! Aborting execution."
}
</code></pre>

<p><strong>Tip:</strong> Replace <code>INSERT_OFFICIAL_SHA256</code> with the verified hash of the official script.</p>

<hr/>

<h2>⚙ Technical Requirements</h2>
<ul>
  <li>Windows 10 or higher</li>
  <li>Local administrator privileges</li>
  <li>PowerShell 5.0 or higher</li>
  <li>ExecutionPolicy set to Bypass</li>
</ul>

<hr/>

<h2>📌 Notes</h2>
<p>
QuickFix standardizes IT operations, accelerates diagnostics, and ensures traceability.
Restricted to internal use by the  IT Department.
</p>
