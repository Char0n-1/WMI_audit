Import-Module ActiveDirectory

# Output paths
$outputDir = "C:\change_me\wmi_output"
$timestamp = Get-Date -Format "yyyy-MM-dd_HH"
$latestResult = Join-Path $outputDir "Windows11_Scan_$timestamp.csv"
$masterResult = Join-Path $outputDir "Windows11_Complete.csv"

# Ensure output directory exists
if (!(Test-Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}

# Get computer list
$computers = Get-ADComputer -Filter * -Property Name | Select-Object -ExpandProperty Name

# Run in parallel
$results = $computers | ForEach-Object -Parallel {
    $computerName = $_

    function Get-TotalRamGB {
        try {
            $ram = Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $computerName -ErrorAction Stop
            $total = ($ram | Measure-Object -Property Capacity -Sum).Sum
            [math]::Round($total / 1GB, 2)
        } catch {
            "Unavailable"
        }
    }

    function Get-IPAddress {
        try {
            $nic = Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $computerName -Filter "IPEnabled = True" -ErrorAction Stop
            ($nic.IPAddress | Where-Object { $_ -notlike "169.*" -and $_ -notlike "::*" })[0]
        } catch {
            "Unavailable"
        }
    }

    try {
        $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computerName -ErrorAction Stop
        $cs = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computerName -ErrorAction Stop
        $cpu = Get-WmiObject -Class Win32_Processor -ComputerName $computerName -ErrorAction Stop

        $buildNumber = [int]$os.BuildNumber
        $isWindows11 = $buildNumber -ge 22000
        $ramGB = Get-TotalRamGB
        $ipAddress = Get-IPAddress

        [PSCustomObject]@{
            ComputerName  = $computerName
            OS            = $os.Caption
            BuildNumber   = $buildNumber
            Version       = $os.Version
            IsWindows11   = $isWindows11
            CPU           = $cpu.Name
            RAM_GB        = $ramGB
            LoggedInUser  = $cs.UserName
            IPAddress     = $ipAddress
            LastChecked   = (Get-Date).ToString("s")
            Error         = ""
        }
    } catch {
        [PSCustomObject]@{
            ComputerName  = $computerName
            OS            = "Unavailable"
            BuildNumber   = "N/A"
            Version       = "N/A"
            IsWindows11   = $false
            CPU           = "Unavailable"
            RAM_GB        = "Unavailable"
            LoggedInUser  = "Unavailable"
            IPAddress     = "Unavailable"
            LastChecked   = (Get-Date).ToString("s")
            Error         = $_.Exception.Message
        }
    }
} -ThrottleLimit 20

# Save latest scan
$results | Export-Csv -Path $latestResult -NoTypeInformation -Encoding UTF8

# Merge with master
if (Test-Path $masterResult) {
    $master = Import-Csv $masterResult
    foreach ($item in $results) {
        $existing = $master | Where-Object { $_.ComputerName -eq $item.ComputerName }
        if ($existing) {
            $index = $master.IndexOf($existing)
            $master[$index] = $item
        } else {
            $master += $item
        }
    }
} else {
    $master = $results
}

# Save master file
$master | Export-Csv -Path $masterResult -NoTypeInformation -Encoding UTF8

Write-Host "Latest result saved to: $latestResult"
Write-Host "Master list updated at: $masterResult"
