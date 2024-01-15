# Checking for ImmyBot Agent and uninstalling if it exists
Write-Host "ImmyBot Agent Checks: Started"
# Define the path of the executable
$exePath = "C:\Program Files (x86)\ImmyBot\Immybot.Agent.Service.exe"

# Check if the executable exists
if (Test-Path $exePath) {
    Write-Host "ImmyBot Agent found. Attempting to uninstall..."

    # Attempt to uninstall using msiexec
    try {
        $product = Get-WmiObject -Class Win32_Product| Where-Object { $_.Name -eq "ImmyBot Agent" }
if ($product) {
$identifyingNumber = $product.IdentifyingNumber
$uninstallString = "/x " + $identifyingNumber + " /quiet /noreboot"
Start-Process "msiexec" -Args $uninstallString -Wait

bash

        # Check if the executable still exists after uninstallation
        if (Test-Path $exePath) {
            Write-Host "Failed to uninstall ImmyBot Agent."
        } else {
            Write-Host "ImmyBot Agent successfully

uninstalled."
}
} else {
Write-Host "Product 'ImmyBot Agent' not found in installed programs."
}
} catch {
Write-Host "An error occurred during uninstallation: $_"
}
} else {
Write-Host "ImmyBot Agent not found. No action needed."
}

# Checking and terminating ScreenConnect.WindowsClient
$screenConnectWindowsClient = Get-Process "ScreenConnect.WindowsClient" -ErrorAction SilentlyContinue
if ($screenConnectWindowsClient) {
    $screenConnectWindowsClient | Stop-Process -Force
    Write-Host "Terminated ScreenConnect.WindowsClient instances."
} else {
    Write-Host "Did not find any ScreenConnect.WindowsClient instances."
}

# Checking and terminating ScreenConnect.ClientService
$screenConnectClientService = Get-Process "ScreenConnect.ClientService" -ErrorAction SilentlyContinue
if ($screenConnectClientService) {
    $screenConnectClientService | Stop-Process -Force
    Write-Host "Terminated ScreenConnect.ClientService instances."
} else {
    Write-Host "Did not find any ScreenConnect.ClientService instances."
}

# Delete specified directory
$directoryPath = "C:\Users\Administrator\AppData\Local\Apps\2.0"
if (Test-Path $directoryPath) {
    Remove-Item -Path $directoryPath -Recurse -Force
    Write-Host "Deleted contents of directory: $directoryPath"
} else {

Write-Host "Directory $directoryPath not found. No action needed."

}
#Additional checks for ScreenConnect.WindowsClient

$screenConnectWindowsClientCheck = Get-Process "ScreenConnect.WindowsClient" -ErrorAction SilentlyContinue
if ($screenConnectWindowsClientCheck) {
Write-Host "Warning: ScreenConnect.WindowsClient instances are still running."
} else {
Write-Host "No remaining instances of ScreenConnect.WindowsClient."
}


Write-Host "Initiating snapshot process..."

# Server and Authentication Details
$baseUrl = "https://pve.murraycompany.com:8006/api2/json"
$tokenId = "root@pam!PVE01API"
$secret = "4cb65774-6e37-4dfd-ac7b-e27316c05384"

# Headers
$headers = @{
    "Authorization" = "PVEAPIToken=$tokenId=$secret"
}

# Image Selection
Write-Host "Select the Image to snapshot:"
Write-Host "1. Detailing Image"
Write-Host "2. Accounting Image"
Write-Host "3. Operations"
$selection = Read-Host "Enter your choice (1, 2, or 3)"

$node = "pve" # Cluster Node Name
switch ($selection) {
    "1" { $vmid = "103" }
    "2" { $vmid = "108" }
    "3" { $vmid = "105" }
    default { Write-Host "Invalid selection"; Exit }
}

# Snapshot Name based on Current Date and Time
$snapname = "Sysprep_" + (Get-Date -Format "MM_dd_yyyy_HHmm")

#Description
$description = "Auto Snapshot before Sysprep"


# Take Snapshot and Get UPID
$snapshotUrl = "$baseUrl/nodes/$node/qemu/$vmid/snapshot"
$snapshotBody = @{
    snapname = $snapname
    description = $description
    vmstate = 1  # Optional: Include RAM state in the snapshot
}
try {
    $response = Invoke-RestMethod -Uri $snapshotUrl -Method POST -Headers $headers -Body $snapshotBody
    $upid = $response.data
    Write-Host "Snapshot task initiated with UPID: $upid"
} catch {
    Write-Host "Error initiating snapshot: $_"
    Exit
}
Write-Host "Entering task completion loop..."
# Start Timer
$startTime = Get-Date

# Function to Get Task Status using UPID
function Get-TaskStatus {
    param ($upid)
    $taskStatusUrl = "$baseUrl/nodes/$node/tasks/$upid/status"
    try {
        $taskStatusResponse = Invoke-RestMethod -Uri $taskStatusUrl -Method GET -Headers $headers
        return $taskStatusResponse
    } catch {
        Write-Host "Error retrieving task status: $_"
        return $null
    }
}

# Wait for Task to Complete
while ($true) {
    $taskStatus = Get-TaskStatus -upid $upid
    $currentTime = Get-Date
    $elapsedTime = $currentTime - $startTime
    $elapsedSeconds = [math]::Round($elapsedTime.TotalSeconds, 0)
    Write-Host "Checking task status..."
    if ($taskStatus -and $taskStatus.data) {
        if ($taskStatus.data.status -eq "stopped") {
            if ($taskStatus.data.exitstatus -eq "OK") {
                Write-Host "Snapshot task completed successfully. Time elapsed: $elapsedSeconds seconds"
                break
            } else {
                Write-Host "Snapshot task stopped with exit status: $($taskStatus.data.exitstatus). Time elapsed: $elapsedSeconds seconds"
                break
            }
        } else {
            Write-Host "Current Task Status: $($taskStatus.data.status). Time elapsed: $elapsedSeconds seconds"
        }
    } else {
        Write-Host "Waiting for task status update... Time elapsed: $elapsedSeconds seconds"
    }
    Start-Sleep -Seconds 5
}

Write-Host "Snapshot $snapname for VM ID $vmid has been successfully created."

Write-Host "Prompting for sysprep confirmation..."


# Sysprep Confirmation
$sysprepConfirmation = Read-Host "Ready to Sysprep? (Yes/No)"
if ($sysprepConfirmation -ne "Yes") {
    Write-Host "Exiting without running sysprep."
    Exit
}

# Sysprep Command Execution
Write-Host "Running Sysprep..."
$sysprepPath = "$env:windir\system32\sysprep\sysprep.exe"
if (Test-Path $sysprepPath) {
    & $sysprepPath /generalize /oobe /shutdown /unattend:c:\unattend.xml
    Write-Host "Sysprep command executed."
} else {
    Write-Host "Sysprep executable not found."
}
