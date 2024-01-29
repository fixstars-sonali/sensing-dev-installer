<#
.SYNOPSIS
Script to update enviroment for sensing dev installer usage.

.DESCRIPTION
Script to update enviroment for sensing dev installer usage.

.PARAMETER installPath
Path of installation.

.EXAMPLE
.\Env.ps1
Runs the script with the default installPath set to the parent directory.

.EXAMPLE
.\YourScript.ps1 -InputPath "C:\example"
Runs the script with the installPath set to "C:\example".

.NOTES
File Name      : Env.ps1
Prerequisite   : PowerShell V3
#>
param(
    [string]$installPath = (Split-Path $PSScriptRoot -Parent)    
)

# Define the paths you want to add
$newPath = "${installPath}\bin"
$newPythonPath = "${installPath}\lib\site-packages"

Write-Output "installPath : $installPath"

# Get current PATH and PYTHONPATH
$currentPath = [Environment]::GetEnvironmentVariable("Path", "User")
$currentPythonPath = [Environment]::GetEnvironmentVariable("PYTHONPATH", "User")

# Update PATH if the new path is not already in it
if (-not $currentPath.Contains($newPath)) {
    $currentPath += ";$newPath"
    [Environment]::SetEnvironmentVariable("Path", $currentPath, "User")
}
Write-Verbose "Updated PATH: $currentPath"

# Update PYTHONPATH if the new path is not already in it
if (-not $currentPythonPath -or (-not $currentPythonPath.Contains($newPythonPath))) {
    if ($currentPythonPath) {
        $currentPythonPath += ";$newPythonPath"
    } else {
        $currentPythonPath = $newPythonPath
    }
    [Environment]::SetEnvironmentVariable("PYTHONPATH", $currentPythonPath, "User")
}
Write-Verbose "Updated PYTHONPATH: $currentPythonPath"

# Update SENSING_DEV_ROOT if the new path is not already in it
[Environment]::SetEnvironmentVariable("SENSING_DEV_ROOT", $installPath, "User")
Write-Verbose "Updated SENSING_DEV_ROOT: $installPath"

[Environment]::SetEnvironmentVariable("GST_PLUGIN_PATH", $installPath, "User")
Write-Verbose "Updated GST_PLUGIN_PATH: $installPath"
