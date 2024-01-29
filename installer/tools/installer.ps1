<#
.SYNOPSIS
Installs the Sensing SDK.

.DESCRIPTION
This script downloads and installs the Sensing SDK. You can specify a particular version or the latest version will be installed by default. It supports both .zip and .msi installers.

.PARAMETER version
Specifies the version of the Sensing SDK to be installed. Default is 'latest'.

.PARAMETER user
Specifies the username for which the Sensing SDK will be installed. This determines the installation path in the user's LOCALAPPDATA.

.PARAMETER Url
URL of the Sensing SDK installer. If not provided, the script constructs the URL based on the specified or default version.

.PARAMETER installPath
The installation path for the Sensing SDK. Default is the sensing-dev-installer directory in the user's LOCALAPPDATA.

.PARAMETER InstallOpenCV
If set, the script will also install OpenCV. This is not done by default.

.EXAMPLE
PS C:\> .\installer.ps1 -version 'v24.09.03' -user 'Admin' -Url 'http://example.com'

This example demonstrates how to run the script with custom version, user, and URL values.

.EXAMPLE
PS C:\> .\installer.ps1 -InstallOpenCV

This example demonstrates how to run the script with the default settings and includes the installation of OpenCV.

.NOTES
Ensure that you have the necessary permissions to install software and write to the specified directories.

.LINK
http://example.com/documentation-link

#>

param(
  [string]$version,
  [string]$user,
  [string]$Url,
  [string]$installPath,
  [switch]$InstallOpenCV = $false 
)

function Test-WritePermission {
  param (
    [string]$path,
    [string]$user = $env:USERNAME
  )
  if (-not $user) {
    $user = $env:USERNAME
  }
  $user = "$env:USERDOMAIN\$user"
  $writeAllowed = $false
  $acl = Get-Acl $path
  $securityIdentifier = New-Object System.Security.Principal.NTAccount($user)

  foreach ($access in $acl.Access) {
    if ($access.IdentityReference -eq $securityIdentifier) {
      if (
        ( 
          $access.FileSystemRights -match "Write" -or 
          $access.FileSystemRights -match "FullControl"
        ) -and 
        $access.AccessControlType -eq "Allow") {
        Write-Host "$user has write permission for $path"
        $writeAllowed = $true
        break
      }
    }
  }

  if (-not $writeAllowed) {
    Write-Host "$user needs write permission for $path"
  }

  return $writeAllowed
}
function Get-LatestVersion {
  [CmdletBinding()]
  param (
    [string]$Repository = "Sensing-Dev/sensing-dev-installer"
  )

  $RepoApiUrl = "https://api.github.com/repos/$Repository/releases/latest"

  try {
    $response = Invoke-RestMethod -Uri $RepoApiUrl -Headers @{Accept = "application/vnd.github.v3+json" }
    $latestVersion = $response.tag_name

    if ($latestVersion) {
      Write-Output $latestVersion
    }
    else {
      Write-Error "Latest version not found."
    }
  }
  catch {
    Write-Error "Error fetching the latest version: $_"
  }
}

function Install-ZipPackage {
  param (
    [string]$Url,
    [string]$installPath,
    [string]$installerName,
    [string]$versionNum
  )

  $tempZipPath = Join-Path $env:TEMP "$installerName.zip"
  Invoke-WebRequest -Uri $Url -OutFile $tempZipPath

  $tempExtractionPath = Join-Path $installPath "_tempExtraction"
  if (!(Test-Path $tempExtractionPath)) {
    New-Item -Path $tempExtractionPath -ItemType Directory | Out-Null
  }

  Expand-Archive -Path $tempZipPath -DestinationPath $tempExtractionPath

  $finalInstallPath = Join-Path $installPath $installerName
  if (Test-Path $finalInstallPath) {
    Remove-Item -Path $finalInstallPath -Force -Recurse
  }
  New-Item -Path $finalInstallPath -ItemType Directory | Out-Null

  Move-Item -Path "$tempExtractionPath\*" -Destination $finalInstallPath -Force

  Remove-Item -Path $tempZipPath -Force
  Remove-Item -Path $tempExtractionPath -Force -Recurse

  Write-Host "Installation completed successfully."
}

function Install-MsiPackage {
  param (
    [string]$Url,
    [string]$installPath,
    [string]$installerName,
    [bool]$hasAccess
  )

  $tempMsiPath = Join-Path $env:TEMP "$installerName.msi"
  Invoke-WebRequest -Uri $Url -OutFile $tempMsiPath

  $logPath = Join-Path $env:TEMP "$installerName-install.log"
  $msiExecArgs = "/i `"$tempMsiPath`" INSTALL_ROOT=`"$installPath`" /qb /l*v `"$logPath`""

  if ($hasAccess) {
    Start-Process -Wait -FilePath "msiexec.exe" -ArgumentList $msiExecArgs
  }
  else {
    Start-Process -Wait -FilePath "msiexec.exe" -ArgumentList $msiExecArgs -Verb RunAs
  }
  if ($?) {
    Write-Host "${installerName} installed at ${installPath}. See detailed log here ${log} "
  }
  else {
    Write-Error "The ${installerName} installation encountered an error. See detailed log here ${log}"        
  }
  if (Test-Path $logPath) {
    Write-Host "Installation log created at $logPath"
  }

  Remove-Item -Path $tempMsiPath -Force

  Write-Host "Installation completed successfully."
}


# Set default installPath if not provided
if (-not $installPath) {
  $installPath = "$env:LOCALAPPDATA"
  if ($user) {
    $UserProfilePath = "C:\Users\$user"
    $installPath = Join-Path -Path $UserProfilePath -ChildPath "AppData\Local"
  }
}
Write-Verbose "installPath = $installPath"

$hasAccess = Test-WritePermission -user $user -path $installPath

$installerName = "sensing-dev"
$installerPostfixName = if ($InstallOpenCV) { "" } else { "-no-opencv" }

# Construct download URL if not provided
if (-not $Url) {
  $baseUrl = "https://github.com/Sensing-Dev/sensing-dev-installer/releases/download/"

  if (-not $version) {
    $version = Get-LatestVersion
  }

  if ($version -match 'v(\d+\.\d+\.\d+)(-\w+)?') {
    $versionNum = $matches[1] 
    Write-Output "Installing version: $version" 
  }

  $downloadBase = "${baseUrl}${version}/${installerName}${installerPostfixName}-${versionNum}-win64"
  $Url = if ($user) { "${downloadBase}.zip" } else { "${downloadBase}.msi" }
  Write-Host "URL : $Url"
}

# Check if the URL ends with .zip or .msi and call the respective function
if ($url.EndsWith("zip")) {
  Install-ZipPackage -Url $url -installPath $installPath -installerName $installerName -versionNum $versionNum
}
elseif ($url.EndsWith("msi")) {
  Install-MsiPackage -Url $url -installPath $installPath -installerName $installerName -hasAccess $hasAccess
}
else {
  Write-Error "Unsupported installer format."
}

if (Test-Path -Path ${installPath}) {
  $relativeScriptPath = "tools\Env.ps1"
  # Run the .ps1 file from the installed package
  $ps1ScriptPath = Join-Path -Path $installPath -ChildPath $relativeScriptPath
  Write-Verbose "ps1ScriptPath = $ps1ScriptPath"
  if (Test-Path -Path $ps1ScriptPath -PathType Leaf) {
    & $ps1ScriptPath
  }
  else {
    Write-Error "Script at $relativeScriptPath not found in the installation path!"
  }
}


