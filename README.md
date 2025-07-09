# Windows 11 WMI Audit Script

## Overview

This script collects hardware and system information from all domain-joined computers in Active Directory using WMI, and checks whether each machine is running **Windows 11**.

The results are:

* Saved hourly in timestamped CSV files
* Merged into a master CSV file that tracks the latest known state of each device

## Collected Information

For each computer, the script gathers:

| Field        | Description                                     |
| ------------ | ----------------------------------------------- |
| ComputerName | Hostname of the machine                         |
| OS           | Operating system caption (e.g., Windows 11 Pro) |
| BuildNumber  | OS build number (e.g., 26100)                   |
| Version      | Full OS version                                 |
| IsWindows11  | `True` if build number is ≥ 22000               |
| CPU          | Name of the processor                           |
| RAM\_GB      | Total installed RAM (in GB)                     |
| LoggedInUser | Currently logged-in user                        |
| IPAddress    | First non-link-local IP address                 |
| LastChecked  | Timestamp of when the machine was scanned       |
| Error        | WMI or RPC failure message (if any)             |

## Features

- Uses WMI (`Win32_OperatingSystem`) to retrieve OS version and build
- Determines whether the OS is Windows 11 (`BuildNumber >= 22000`)
- Runs in parallel (`ForEach-Object -Parallel`) for faster execution
- Filters out unreachable computers with `Test-Connection`
- Saves each scan to a timestamped CSV
- Maintains a cumulative master list (one row per computer)
- Does **not** overwrite good data with failed/unreachable results

---

## Output Files

All files are saved in:

```
C:\change_me\wmi_output
```

* `Windows11_Scan_yyyy-MM-dd_HH.csv` – the latest hourly scan
* `Windows11_Complete.csv` – the cumulative file, updated on each run

## How to Run

1. **Requirements**:

   * PowerShell 7+
   * AD module (`Import-Module ActiveDirectory`)
   * Domain-joined system with rights to query remote machines via WMI
   * User account must have **Log on as a batch job** rights *(or run as SYSTEM)*

2. **Run manually**:

   ```powershell
   pwsh -File .\YourScriptName.ps1
   ```

3. **Schedule with Task Scheduler**:

   * Trigger: Repeat every 1 hour
   * Action:

     * Program: `pwsh.exe`
     * Arguments: `-File "C:\Path\To\YourScript.ps1"`

## Notes

* The script uses `ForEach-Object -Parallel` for performance (PowerShell 7 only)
* If WMI or RPC is blocked on a machine, it will be logged in the `Error` field
* Offline or unreachable machines will not stop the script

## Troubleshooting

* Ensure firewalls allow WMI (RPC) traffic
* Ensure `Remote Administration` and `Windows Management Instrumentation` services are running on targets
* You may need admin rights on remote machines


