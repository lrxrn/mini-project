# Example usage: .\install_python.ps1 -version 3.10.11 -upgrade $true -reinstall $false
# Version: 1.1

param (
    [string]$version = "3.12.9", # Default version as 3.12
    [bool]$upgrade = $true,  # Upgrade flag
    [bool]$reinstall = $false # Flag to reinstall python if already installed (Remove and download again)
)

# list of allowed versions
# latest stable versions of 3.8, 3.9, 3.10, 3.11, 3.12 
$allowedVersions = @("3.8.10", "3.9.13", "3.10.11", "3.11.9", "3.12.9")

# log file in user's local temp directory
$logFile = "$env:TEMP\python_install_log.txt"

# log messages to file and console
function Write-Log {
    param (
        [string]$message,
        [string]$color = "White"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $message"
    Write-Host $logEntry -ForegroundColor $color
    Add-Content -Path $logFile -Value $logEntry
}

# check for administrator privileges and restart script with elevated privileges if not already
function Ensure-Admin {
    if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Log "Script is not running as administrator. Restarting with elevated privileges..." -color "Yellow"
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
        Write-Log "ERROR: Python $version is not supported on Windows 7 or earlier. Exiting..." -color "Red"
        exit 1
    }
}

# get the installed Python version
function Get-InstalledPythonVersion {
    try {
        # Try multiple possible Python commands
        $pythonCommands = @("python")
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
        Write-Log "Error checking Python version: $_" -color "Red"
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
        Write-Log "64-bit architecture detected." -color "Gray"
        return "-amd64"
    } else {
        Write-Log "32-bit architecture detected." -color "Gray"
        return ""
    }
}

# function to check if the temp file exists and remove if it does
function Remove-TempFile {
    param ([string]$filePath)
    Write-Log "Checking if temp file exists from previous run..."
    if (Test-Path $filePath) {
        Write-Log "Removing temp installer file: $filePath" -color "Gray"
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
        Write-Log "Download failed!" -color "Red"
        exit 1
    }

    Write-Log "Download completed: $installerPath" -color "Green"
    return $installerPath
}

# Actually do the install Python stuff
function Install-Python {
    param ([string]$installerPath)
    
    Write-Log "Starting Python installation..."
    # ref: https://www.python.org/download/releases/2.5/msi/
    Start-Process -FilePath $installerPath -ArgumentList "/quiet", "InstallAllUsers=1", "AppendPath=1" -Wait -NoNewWindow

    # check installation was successful
    $installed = Get-InstalledPythonVersion
    if ($installed) {
        Write-Log "Python version $installed installed successfully!" -color "Green"
    } else {
        Write-Log "Installation failed!" -color "Red"
        exit 1
    }
}

function Uninstall-AllPythonVersions {
    # Search for Python installations
    $uninstallString = "Python*"
    # search registry keys for uninstall information
    $uninstallKeys = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue | 
                     Where-Object { $_.DisplayName -like $uninstallString }

    if ($uninstallKeys) {
        foreach ($key in $uninstallKeys) {
            # make sure DisplayName exists and is not null
            if ($key.DisplayName) {
                $pythonVersion = $key.DisplayName
                Write-Log "Uninstalling: $pythonVersion" -color "Gray"
                # Uninstalling the version using msiexec
                Start-Process "msiexec.exe" -ArgumentList "/x $($key.PSChildName) /quiet" -Wait -NoNewWindow
            }
        }
    } else {
        Write-Log "No Python versions found to uninstall."
    }
}


# ensure the script is running as admin
Ensure-Admin

# Check Windows version is compatible
Check-WindowsVersion

# Main program logic starts here
# start logging
Write-Host "Logging into file: $logFile" -ForegroundColor "Gray"
Write-Log "===== Start install ====="

if ($version -notin $allowedVersions) {
    Write-Log "Invalid version provided! Allowed versions: $($allowedVersions -join ', ')" -color "Red"
    exit 1
}

$currentVersion = Get-InstalledPythonVersion
Write-Log "Currently installed Python version: $(if ($currentVersion) { $currentVersion } else { 'None' })"
if ($reinstall -and $currentVersion) {
    Write-Log "Reinstall flag set. Checking for existing Python installations..." -color "Yellow"
    Uninstall-AllPythonVersions
    Write-Log "Existing Python installation removed." -color "Gray"
    $currentVersion = $null
}

if (Is-NewVersionHigher -installedVersion $currentVersion -newVersion $version) {
    if ($currentVersion) {
        if ($upgrade) {
            Write-Log "Upgrade flag set. Proceeding with upgrading from $currentVersion to $version..." -color "Gray"
        } else {
            Write-Log "New version $version is available. But upgrade flag not set. Skipping upgrade..." -color "Yellow"
            exit 0
        }
    } else {
        Write-Log "Python is not installed. Proceeding with installation of version $version..." -color "Yellow"
    }
    $installerPath = Download-PythonInstaller -pythonVersion $version
    Install-Python -installerPath $installerPath
    Remove-Item -Path $installerPath -Force -ErrorAction SilentlyContinue  # Cleanup and remove installer
    Write-Log "Installer file removed: $installerPath" -color "Gray"
} else {
    Write-Log "Python is already up-to-date. No installation/ upgrading needed." -color "Green"
}

Write-Log "===== Finish install ====="