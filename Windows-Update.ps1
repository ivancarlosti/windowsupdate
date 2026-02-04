# Bypass execution policy for current session
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue

# Get the directory where the script is located
$ScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

# Define logging path (same directory as script)
$LogDirectory = $ScriptDirectory
$LogFile = Join-Path -Path $LogDirectory -ChildPath "windowsupdate.log"

# Function for consistent error handling
function Handle-Error {
    param (
        [string]$ErrorMessage,
        [int]$ExitCode = 1
    )
    Write-Output "[ERROR] $ErrorMessage"
    try {
        Stop-Transcript -ErrorAction SilentlyContinue
    } catch {
        Write-Output "[WARNING] Failed to stop transcript: $_"
    }
    Exit $ExitCode
}

# Ensure the log directory exists (script directory)
try {
    if (!(Test-Path -Path $LogDirectory)) {
        New-Item -ItemType Directory -Force -Path $LogDirectory | Out-Null
        Write-Output "[INFO] Created log directory: $LogDirectory"
    }
} catch {
    Handle-Error -ErrorMessage "Failed to create log directory: $_"
}

# Start logging
try {
    Start-Transcript -Path $LogFile -Append -ErrorAction Stop
    Write-Output "`n========================="
    Write-Output "  Windows Update Script  "
    Write-Output "=========================`n"
    Write-Output "[INFO] Logging to: $LogFile"
} catch {
    Handle-Error -ErrorMessage "Failed to start transcript: $_"
}

# Function to check and install missing modules
function Ensure-Module {
    param (
        [string]$ModuleName,
        [string]$ProviderName = $null
    )
    try {
        if (!(Get-Module -Name $ModuleName -ListAvailable)) {
            Write-Output "[INFO] Installing module: $ModuleName"
            
            # Install NuGet provider if needed (with compatibility for all PS versions)
            if ($ProviderName -and !(Get-PackageProvider -Name $ProviderName -ListAvailable -ErrorAction SilentlyContinue)) {
                try {
                    # First try modern approach
                    $params = @{
                        Name          = $ProviderName
                        Force         = $true
                        ErrorAction   = 'Stop'
                    }
                    
                    # Only add SkipPublisherCheck if parameter exists
                    if (Get-Command Install-PackageProvider -ParameterName SkipPublisherCheck -ErrorAction SilentlyContinue) {
                        $params.SkipPublisherCheck = $true
                    }
                    
                    Install-PackageProvider @params
                } catch {
                    # Fallback method if standard installation fails
                    try {
                        Write-Output "[INFO] Trying fallback provider installation method..."
                        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                        $sourceNugetExe = "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe"
                        $targetNugetExe = "$env:TEMP\nuget.exe"
                        Invoke-WebRequest $sourceNugetExe -OutFile $targetNugetExe
                        & $targetNugetExe install NuGet.CommandLine -ExcludeVersion -OutputDirectory "$env:ProgramFiles\PackageManagement\ProviderAssemblies"
                        Write-Output "[SUCCESS] Provider installed via fallback method"
                    } catch {
                        Handle-Error -ErrorMessage "Failed to install package provider $($ProviderName): $($_)"
                    }
                }
            }
            
            # Install the module (with compatibility for all PS versions)
            try {
                $params = @{
                    Name          = $ModuleName
                    Force         = $true
                    ErrorAction   = 'Stop'
                }
                
                # Only add SkipPublisherCheck if parameter exists
                if (Get-Command Install-Module -ParameterName SkipPublisherCheck -ErrorAction SilentlyContinue) {
                    $params.SkipPublisherCheck = $true
                }
                
                Install-Module @params
                Write-Output "[SUCCESS] Module $ModuleName installed successfully"
            } catch {
                Handle-Error -ErrorMessage "Failed to install module $($ModuleName): $($_)"
            }
        } else {
            Write-Output "[INFO] Module already installed: $ModuleName"
        }
    } catch {
        Handle-Error -ErrorMessage "Module check failed for $($ModuleName): $($_)"
    }
}

# Ensure required modules are installed
Ensure-Module -ModuleName "PSWindowsUpdate" -ProviderName "NuGet"

# Import the PSWindowsUpdate module
try {
    Import-Module PSWindowsUpdate -Force -ErrorAction Stop
    Write-Output "[SUCCESS] PSWindowsUpdate module imported."
} catch {
    Handle-Error -ErrorMessage "Could not import PSWindowsUpdate module: $_"
}

# Function to check and install Windows updates
function Run-WindowsUpdate {
    try {
        Write-Output "`n[INFO] Checking for Windows updates..."
        $Updates = Get-WindowsUpdate -IgnoreReboot -ErrorAction Stop
        
        if ($Updates) {
            Write-Output "[INFO] Found $($Updates.Count) updates available, proceeding with installation..."
            try {
                $installResult = Install-WindowsUpdate -AcceptAll -IgnoreReboot -ErrorAction Stop
                Write-Output "[SUCCESS] Windows updates installed successfully!"
                Write-Output ($installResult | Out-String)
            } catch {
                Handle-Error -ErrorMessage "Update installation failed: $_"
            }
        } else {
            Write-Output "[INFO] No updates available."
        }
    } catch {
        Handle-Error -ErrorMessage "Windows Update check failed: $_"
    }
}

# Run the Windows Update process
Run-WindowsUpdate

# Stop logging
try {
    Stop-Transcript
    Write-Output "[INFO] Logging completed. Transcript saved to $LogFile"
} catch {
    Write-Output "[WARNING] Failed to stop transcript properly: $_"
    Exit 0  # This is non-critical, so we exit with success
}

Exit 0