<h1 align="center">🛠 QuickFix</h1>
<h3 align="center">Corporate Support & Diagnostic Script</h3>
<p align="center"><strong>Clivalemais – IT Department</strong></p>

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
<strong>QuickFix</strong> is a corporate PowerShell-based tool developed to standardize IT support procedures,
reduce ticket resolution time, and ensure full traceability of technical interventions performed within Clivalemais.
</p>

<p>
The script runs natively on Windows environments and does not require installation or third-party dependencies.
</p>

<h3>Core Design Principles</h3>

<ul>
  <li>Operational standardization</li>
  <li>Minimal system intervention</li>
  <li>Full action traceability</li>
  <li>Modular and isolated execution</li>
</ul>

<p>
No potentially impactful action is executed without explicit technician confirmation.
</p>

<hr/>

<h2>🏗 Technical Architecture</h2>

<p>
QuickFix is structured into <strong>6 independent modules</strong>, accessible through a numeric terminal menu.
Each module operates in isolation to reduce operational risk.
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

<p>Performs a complete system information scan without modifying any configuration.</p>

<strong>Data Collected:</strong>

<ul>
  <li>Operating System</li>
  <li>CPU</li>
  <li>Motherboard</li>
  <li>RAM modules (individual detection)</li>
  <li>Video controllers</li>
  <li>Storage devices (SSD/HDD + free space analysis)</li>
</ul>

<strong>Technical Logic:</strong>

<ul>
  <li>WMI/CIM queries (Win32_OperatingSystem, Win32_Processor, etc.)</li>
  <li>Automatic SSD/HDD detection via SpindleSpeed</li>
  <li>Free disk space percentage calculation with internal alert thresholds</li>
</ul>

<p><strong>Risk Level:</strong> Read-only</p>

<hr/>

<h3>2️⃣ Network Status</h3>

<p>Provides structured connectivity diagnostics.</p>

<strong>Functions:</strong>

<ul>
  <li>Active network adapters</li>
  <li>IP configuration (DHCP or static)</li>
  <li>DNS servers</li>
  <li>Gateway connectivity test</li>
  <li>External connectivity test (8.8.8.8)</li>
  <li>DNS resolution validation</li>
</ul>

<p><strong>Risk Level:</strong> Read-only</p>

<hr/>

<h3>3️⃣ Printer Diagnostics & Repair</h3>

<ul>
  <li>Status and queue inspection</li>
  <li>Port and IP mapping</li>
  <li>Printer connectivity test</li>
  <li>Spooler service restart</li>
  <li>Manual queue cleanup (.SPL / .SHD removal)</li>
  <li>Set default printer</li>
  <li>Custom test page generation</li>
  <li>Force printer online state</li>
</ul>

<p>Some operations may interrupt active print jobs and require confirmation.</p>

<hr/>

<h3>4️⃣ RAM Optimization</h3>

<ul>
  <li>Global memory cleanup via EmptyWorkingSet</li>
  <li>Selective high-consumption process reduction</li>
  <li>Permanent Chrome memory-saving flags</li>
</ul>

<p><strong>Risk Level:</strong> Light impact (temporary performance adjustment)</p>

<hr/>

<h3>5️⃣ Network Repair</h3>

<ul>
  <li>Full network stack repair (Release / Renew / Winsock / RegisterDNS)</li>
  <li>DNS cache flush</li>
  <li>Layered connectivity test</li>
  <li>Network adapter reset</li>
</ul>

<p>May temporarily interrupt active connection.</p>

<hr/>

<h3>6️⃣ Windows System Repair</h3>

<ul>
  <li>DISM /RestoreHealth</li>
  <li>SFC /Scannow</li>
  <li>DISM /StartComponentCleanup</li>
</ul>

<p>
Average duration: 10–20 minutes<br/>
Requires reboot<br/>
Does not remove user data
</p>

<hr/>

<h2>📄 Reporting System</h2>

<p>
Each executed action generates a <code>.txt</code> report stored at:
</p>

<pre><code>C:\services\relatorios</code></pre>

<strong>Logged Information:</strong>

<ul>
  <li>Date and time</li>
  <li>Technician name</li>
  <li>Machine name</li>
  <li>IP address</li>
  <li>Executed module</li>
  <li>Operation result</li>
</ul>

<hr/>

<h2>🔐 Risk Classification</h2>

<table>
  <tr>
    <th>Level</th>
    <th>Description</th>
    <th>Examples</th>
  </tr>
  <tr>
    <td>Level 1</td>
    <td>No impact</td>
    <td>Diagnostics</td>
  </tr>
  <tr>
    <td>Level 2</td>
    <td>Light impact</td>
    <td>Spooler restart, DNS flush</td>
  </tr>
  <tr>
    <td>Level 3</td>
    <td>Structured repair</td>
    <td>DISM, SFC, Winsock reset</td>
  </tr>
</table>

<hr/>

<h2>⚙ Technical Requirements</h2>

<ul>
  <li>Windows 10 or higher</li>
  <li>Local administrator privileges</li>
  <li>PowerShell 5.0 or higher</li>
  <li>ExecutionPolicy set to Bypass</li>
</ul>

<hr/>

<h2>▶ Execution</h2>

<pre><code>powershell -ExecutionPolicy Bypass -File quickfix.ps1</code></pre>

<hr/>

<h2>📌 Final Notes</h2>

<p>
QuickFix does not replace technical expertise.  
It standardizes operational procedures, accelerates diagnostics, and strengthens IT governance practices.
</p>

<p><strong>Restricted for internal use by the Clivalemais IT Department.</strong></p>
