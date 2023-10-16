
<#
.SYNOPSIS
    Install USB device driver using WinUSB and run a driver installer.

.DESCRIPTION
    This PowerShell script downloads and installs a USB device driver using WinUSB and then runs a driver installer. It also deletes temporary files after the installation.

.NOTES
    File Name      : installer-WinUSB.ps1 
    Prerequisite   : PowerShell 5.0 or later

.EXAMPLE
    .\installer-WinUSB.ps1 

    This example downloads and runs the script to install a USB device driver .

#>

# Function to update the progress bar
function Update-ProgressBar($Percentage, $Status) {
    Write-Progress -Activity "Installing WinUSB Device Driver" -Status $Status -PercentComplete $Percentage
}

# Function to start a process and display a progress bar while waiting for it to complete
function Start-ProcessWithProgressBar($Options, $Status) {
    $process = Start-Process @Options -PassThru

    # Poll the process for its exit while updating the progress
    while (!$process.HasExited) {
        $percentageComplete = ($process.TotalProcessorTime.TotalMilliseconds / $process.StartTime.AddMinutes(5).TotalMilliseconds) * 100
        Update-ProgressBar $percentageComplete $Status
        Start-Sleep -Milliseconds 500
    }
}

$installPath = "$env:LOCALAPPDATA"
Write-Verbose "installPath = $installPath"
# Download win_usb installer

$repoUrl = "https://api.github.com/repos/Sensing-Dev/sensing-dev-installer/releases/latest"
$response = Invoke-RestMethod -Uri $repoUrl
$version = $response.tag_name
$version
Write-Verbose "Latest version: $version" 

if ($version -match 'v(\d+\.\d+\.\d+)(-\w+)?') {
    $versionNum = $matches[1] 
    Write-Output "Installing version: $version" 
}
$installerName = "winusb"

$Url = "https://github.com/Sensing-Dev/sensing-dev-installer/releases/download/v${versionNum}/${installerName}.zip"

$Url

if ($Url.EndsWith("zip")) {
    # Download ZIP to a temp location

    $tempZipPath = "${env:TEMP}\${installerName}.zip"
    Invoke-WebRequest -Uri $Url -OutFile $tempZipPath -Verbose

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $tempExtractionPath = "$installPath\_tempWinUSBExtraction"
    # Create the temporary extraction directory if it doesn't exist
    if (-not (Test-Path $tempExtractionPath)) {
        New-Item -Path $tempExtractionPath -ItemType Directory
    }
    # Attempt to extract to the temporary extraction directory
    try {
        [System.IO.Compression.ZipFile]::ExtractToDirectory($tempZipPath, $tempExtractionPath)
        ls $tempExtractionPath
    }
    catch {
        Write-Error "Extraction failed. Original contents remain unchanged."
        # Optional: Cleanup the temporary extraction directory
        Remove-Item -Path $tempExtractionPath -Force -Recurse
    }    
     # Optionally delete the ZIP file after extraction
     Remove-Item -Path $tempZipPath -Force
}

if (Test-Path $tempExtractionPath) {    

    # Run Winusb installer
    Write-Host "This may take a few minutes. Starting the installation..."

    Write-Verbose "Start winUsb installer"
    $TempDir = "$tempExtractionPath/winusb/temp"
        
    New-item -Path "$TempDir" -ItemType Directory
    $winUSBOptions = @{
        FilePath               = "${tempExtractionPath}/winusb/winusb_installer.exe"
        ArgumentList           = "054c"
        WorkingDirectory       = "$TempDir"
        Wait                   = $true
        Verb                   = "RunAs"  # This attempts to run the process as an administrator
    }
    # Start winusb_installer.exe process with progress bar
    Start-ProcessWithProgressBar @winUSBOptions "Executing winUsb installer..."

    Write-Verbose "End winUsb installer"
}

# Run Driver installer
Write-Verbose "Start Driver installer"

$infPath = "$TempDir/target.inf"
if (-not (Test-Path -Path $infPath -PathType Leaf) ){
    Write-Error "$infPath does not exist."
}
else{
    $pnputilOptions = @{
        FilePath = "PUNPUTIL"
        ArgumentList           = "-i -a $infPath"
        WorkingDirectory       = "$TempDir"
        Wait                   = $true
        Verb                   = "RunAs"  # This attempts to run the process as an administrator
    }
    try {
        # Start Pnputil process with progress bar
        Start-ProcessWithProgressBar @pnputilOptions "Installing driver..."  
    }
    catch {
        Write-Error "An error occurred while running pnputil: $_"
        # You can choose to handle the error as needed, such as logging or taking corrective action.
    }
}

# delete temp files

Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command `"Remove-Item -Path '$tempExtractionPath' -Recurse -Force -Confirm:`$false`""

Write-Verbose "End Driver installer"
