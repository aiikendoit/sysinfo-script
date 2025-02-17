# System Information Collector
# Exports Local System Information to CSV including Printers and Monitors
# Version with detailed hardware components

try {
    # Get hostname first for file naming
    $hostname = [System.Net.Dns]::GetHostName()
    
    # Define output path with hostname in filename
    $outputFile = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\$hostname-system-info.csv"

    # Collect system information with error handling
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
    $computerBIOS = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
    $computerOS = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $computerCPU = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop
    $computerHDD = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop
    
    # Get Printer Information
    $printers = Get-CimInstance -ClassName Win32_Printer
    $printerInfo = $printers | ForEach-Object {
        "{0} (Port: {1}, Status: {2})" -f $_.Name, $_.PortName, $_.Status
    }

    # Get Monitor Information
    $monitors = Get-CimInstance -ClassName WmiMonitorID -Namespace root\wmi
    $monitorInfo = $monitors | ForEach-Object {
        $manufacturerName = [System.Text.Encoding]::ASCII.GetString($_.ManufacturerName).Trim(([char]0))
        $productCodeName = [System.Text.Encoding]::ASCII.GetString($_.ProductCodeID).Trim(([char]0))
        $serialNumber = [System.Text.Encoding]::ASCII.GetString($_.SerialNumberID).Trim(([char]0))
        "Manufacturer: $manufacturerName, Model: $productCodeName, Serial: $serialNumber"
    }

    # Get network adapter information
    $networkAdapters = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration |
        Where-Object { $_.IPEnabled -eq $true }

    # Process disk information
    $diskInfo = $computerHDD | ForEach-Object {
        "{0} ({1:N2} GB, {2:N2}% free)" -f $_.DeviceID, 
            ($_.Size/1GB), 
            ($_.FreeSpace/$_.Size*100)
    }

    # Process network information
    $networkInfo = $networkAdapters | ForEach-Object {
        "{0} - IP: {1}, MAC: {2}" -f $_.Description,
            ($_.IPAddress[0]),
            $_.MACAddress
    }

    # Create system information object
    $systemInfo = [PSCustomObject]@{
        Hostname = $hostname
        Domain = if ($computerSystem.Domain -eq $hostname) { "WORKGROUP" } else { $computerSystem.Domain }
        PCName = $computerSystem.Name
        Manufacturer = $computerSystem.Manufacturer
        Model = $computerSystem.Model
        SerialNumber = $computerBIOS.SerialNumber
        RAM_GB = [math]::Round(($computerSystem.TotalPhysicalMemory/1GB), 2)
        Drives = $diskInfo -join ' | '
        NetworkAdapters = $networkInfo -join ' | '
        IPAddresses = ($networkAdapters.IPAddress | Where-Object { $_ -match '^\d+\.' }) -join ' | '
        MACAddresses = ($networkAdapters.MACAddress) -join ' | '
        CPU = $computerCPU.Name
        OperatingSystem = $computerOS.Caption
        ServicePack = $computerOS.ServicePackMajorVersion
        Username = $computerSystem.UserName
        LastBootTime = $computerOS.LastBootUpTime
        Printers = $printerInfo -join ' | '
        Monitors = $monitorInfo -join ' | '
        CollectionDate = (Get-Date)
    }

    # Export to CSV with error handling
    $systemInfo | Export-Csv -Path $outputFile -NoTypeInformation -Append -ErrorAction Stop

    # Display success message with the actual filename
    Write-Host "System information successfully collected and saved to $outputFile" -ForegroundColor Green
    
    # Display the collected information
    $systemInfo | Format-List
}
catch {
    # Display error in console
    Write-Host "Error collecting system information: $($_.Exception.Message)" -ForegroundColor Red
}