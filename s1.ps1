# System Information Collector
# Exports Local System Information to CSV
# Version with proper multi-drive handling

# Define output path to current user's desktop
$outputFile = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\system-info.csv"

try {
    # Collect system information with error handling
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
    $computerBIOS = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
    $computerOS = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $computerCPU = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop
    $computerHDD = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType=3" -ErrorAction Stop

    # Process disk information
    $diskInfo = $computerHDD | ForEach-Object {
        "{0} ({1:N2} GB, {2:N2}% free)" -f $_.DeviceID, 
            ($_.Size/1GB), 
            ($_.FreeSpace/$_.Size*100)
    }

    # Create system information object
    $systemInfo = [PSCustomObject]@{
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
        CollectionDate = (Get-Date)
    }

    # Export to CSV with error handling
    $systemInfo | Export-Csv -Path $outputFile -NoTypeInformation -Append -ErrorAction Stop

    # Display success message
    Write-Host "System information successfully collected and saved to $outputFile" -ForegroundColor Green
    
    # Display the collected information
    $systemInfo | Format-List
}
catch {
    # Display error in console
    Write-Host "Error collecting system information: $($_.Exception.Message)" -ForegroundColor Red
}