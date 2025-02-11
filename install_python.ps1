# Example usage: .\install_python.ps1 -version 3.10.11
# Version: 1.0

param (
    [string]$version = "3.10.11"  # Default version as 3.10
    [bool]$upgrade = $false  # Upgrade flag
)

# list of allowed versions
# latest stable versions of 3.8, 3.9, 3.10, 3.11, 3.12 
# 3.8.18, 3.9.21, 3.10.16, 3.11.9, 3.12.5
$allowedVersions = @("3.8.18", "3.9.21", "3.10.11", "3.11.9", "3.12.5")

# log file in user's local temp directory
$logFile = "$env:TEMP\python_install_log.txt"

# log messages to file and console
function Write-Log {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $message"
    Write-Host $logEntry
    Add-Content -Path $logFile -Value $logEntry
}

# check for administrator privileges and restart script with elevated privileges if not already
function Ensure-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Log "Script is not running as administrator. Restarting with elevated privileges..."
        Start-Process PowerShell -ArgumentList "-File `"$PSCommandPath`" -version `"$version`"" -Verb RunAs
        exit
    }
}

# check Windows version (Windows 8 or later required for python 3.8+)
function Check-WindowsVersion {
    $os = Get-WmiObject Win32_OperatingSystem
    $major = $os.Version.Split('.')[0]
    $minor = $os.Version.Split('.')[1]

    if ([int]$major -lt 6 -or ([int]$major -eq 6 -and [int]$minor -lt 2)) {
        Write-Log "ERROR: Python $version is not supported on Windows 7 or earlier. Exiting..."
        exit 1
    }
}

# get the installed Python version
function Get-InstalledPythonVersion {
    try {
        # Try multiple possible Python commands
        $pythonCommands = @("python3", "python", "py")
        $pythonPath = $null
        
        foreach ($cmd in $pythonCommands) {
            try {
                $pythonPath = (Get-Command $cmd -ErrorAction Stop).Source
                break  # Exit loop if command is found
            } catch {
                continue  # Try next command
            }
        }

        if (-not $pythonPath) {
            return $null  # No Python installation found
        }

        $installedVersion = & $pythonPath --version 2>&1
        if ([string]::IsNullOrEmpty($installedVersion)) {
            return $null
        }
        
        # Extract version number using regex
        if ($installedVersion -match "\d+\.\d+\.\d+") {
            return $matches[0]
        }
        return $null
    } catch {
        Write-Log "Error checking Python version: $_"
        return $null
    }
}

function Is-NewVersionHigher {
    param (
        [string]$installedVersion,
        [string]$newVersion
    )
    if (-not $installedVersion) { return $true }  # No Python installed, install new
    return ([version]$newVersion -gt [version]$installedVersion)
}

# function to get cpu architecture
function Get-CPUArchitecture {
    $arch = (Get-WmiObject Win32_Processor).AddressWidth
    # https://www.python.org/ftp/python/3.12.5/python-3.12.5-amd64.exe - 64-bit
    # https://www.python.org/ftp/python/3.12.5/python-3.12.5.exe - 32-bit
    if ($arch -eq 64) {
        return "-amd64"
    } else {
        return ""
    }
}

# function to check if the temp file exists and remove if it does
function Remove-TempFile {
    Write-Log "Checking if temp file exists from previous run..."
    param ([string]$filePath)
    if (Test-Path $filePath) {
        Write-Log "Removing temp installer file: $filePath"
        Remove-Item -Path $filePath -Force -ErrorAction SilentlyContinue
    }
}

# function to download Python installer
function Download-PythonInstaller {
    param ([string]$pythonVersion)

    $arch = Get-CPUArchitecture

    Remove-TempFile "$env:TEMP\python-$pythonVersion.exe"
    
    # https://www.python.org/ftp/python/3.10.11/python-3.10.11-amd64.exe
    $installerUrl = "https://www.python.org/ftp/python/$pythonVersion/python-$pythonVersion$arch.exe"
    # store in temp path
    $installerPath = "$env:TEMP\python-$pythonVersion.exe"

    Write-Log "Downloading Python $pythonVersion from $installerUrl..."
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -ErrorAction Stop

    if (!(Test-Path $installerPath)) {
        Write-Log "Download failed!"
        exit 1
    }

    Write-Log "Download completed: $installerPath"
    return $installerPath
}

# Actually do the install Python stuff
function Install-Python {
    param ([string]$installerPath)
    
    Write-Log "Starting Python installation..."
    # ref: https://www.python.org/download/releases/2.5/msi/
    Start-Process -FilePath $installerPath -ArgumentList "/qn", "ALLUSERS=1" -Wait -NoNewWindow

    # check installation was successful
    $installed = Get-InstalledPythonVersion
    if ($installed) {
        Write-Log "Python version $installed installed successfully!"
    } else {
        Write-Log "Installation failed!"
        exit 1
    }
}

# ensure the script is running as admin
Ensure-Admin

# Check Windows version is compatible
Check-WindowsVersion

# Main program logic starts here
# start logging
Write-Log "===== Start install ====="

if ($version -notin $allowedVersions) {
    Write-Log "Invalid version provided! Allowed versions: $($allowedVersions -join ', ')"
    exit 1
}

$currentVersion = Get-InstalledPythonVersion
Write-Log "Currently installed Python version: $(if ($currentVersion) { $currentVersion } else { 'None' })"

if (Is-NewVersionHigher -installedVersion $currentVersion -newVersion $version) {
    if ($currentVersion) {
        if ($upgrade) {
            Write-Log "Upgrade flag set. Proceeding with upgrading from $currentVersion to $version..."
        } else {
            Write-Log "New version $version is available. But upgrade flag not set. Skipping installation..."
            exit 0
        }
    } else {
        Write-Log "Python is not installed. Proceeding with installation of version $version..."
    }
    $installerPath = Download-PythonInstaller -pythonVersion $version
    Install-Python -installerPath $installerPath
    Remove-Item -Path $installerPath -Force -ErrorAction SilentlyContinue  # Cleanup and remove installer
    Write-Log "Installer file removed: $installerPath"
} else {
    Write-Log "Python is already up-to-date. No installation/ upgrading needed."
}

Write-Log "===== Finish install ====="