# Launch Windows Sandbox with automatic DNS configuration and winget installation
# This script creates a sandbox configuration file and launches it

# Function to download winget dependencies using BITS
function Download-WingetDependencies-BITS {
    param(
        [string]$DownloadPath
    )
    
    Write-Host "Downloading winget dependencies using BITS to: $DownloadPath" -ForegroundColor Green
    
    # Create download directory if it doesn't exist
    if (!(Test-Path $DownloadPath)) {
        New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null
        Write-Host "Created download directory: $DownloadPath" -ForegroundColor Cyan
    }
    
    # Check if BITS service is running
    $bitsService = Get-Service -Name "BITS" -ErrorAction SilentlyContinue
    if (-not $bitsService -or $bitsService.Status -ne "Running") {
        Write-Host "BITS service is not running. Starting BITS service..." -ForegroundColor Yellow
        try {
            Start-Service -Name "BITS" -ErrorAction Stop
            Write-Host "BITS service started successfully" -ForegroundColor Green
        } catch {
            Write-Host "Failed to start BITS service: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Falling back to Invoke-WebRequest..." -ForegroundColor Yellow
            Download-WingetDependencies -DownloadPath $DownloadPath
            return
        }
    }
    
    # Download URLs (latest stable versions)
    $downloads = @(
        @{
            Name = "Visual C++ Redistributable Libraries"
            Url = "https://aka.ms/Microsoft.VCLibs.arm64.14.00.Desktop.appx"
            FileName = "VCLibs.appx"
        },
        @{
            Name = "UI.Xaml"
            Url = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.arm64.appx"
            FileName = "UIXaml.appx"
        },
        @{
            Name = "winget"
            Url = "https://github.com/microsoft/winget-cli/releases/latest/download/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
            FileName = "winget.msixbundle"
        }
    )
    
    $jobs = @()
    
    # Start all downloads
    foreach ($download in $downloads) {
        $destinationPath = Join-Path $DownloadPath $download.FileName
        
        Write-Host "Starting BITS download for $($download.Name)..." -ForegroundColor Cyan
        
        try {
            # Remove existing file if it exists
            if (Test-Path $destinationPath) {
                Remove-Item $destinationPath -Force
            }
            
            # Start BITS transfer
            $job = Start-BitsTransfer -Source $download.Url -Destination $destinationPath -Priority Foreground -Asynchronous
            $jobs += @{
                Job = $job
                Name = $download.Name
                FileName = $download.FileName
            }
            
            Write-Host "BITS job started for $($download.Name) (Job ID: $($job.JobId))" -ForegroundColor Green
            
        } catch {
            Write-Host "Failed to start BITS download for $($download.Name): $($_.Exception.Message)" -ForegroundColor Red
            
            # Fallback to Invoke-WebRequest for this file
            Write-Host "Falling back to Invoke-WebRequest for $($download.Name)..." -ForegroundColor Yellow
            try {
                Invoke-WebRequest -Uri $download.Url -OutFile $destinationPath -UseBasicParsing
                Write-Host "$($download.Name) downloaded successfully using fallback method" -ForegroundColor Green
            } catch {
                Write-Host "Fallback download also failed for $($download.Name): $($_.Exception.Message)" -ForegroundColor Red
                throw
            }
        }
    }
    
    # Monitor and wait for all BITS jobs to complete
    if ($jobs.Count -gt 0) {
        Write-Host "`nMonitoring BITS downloads..." -ForegroundColor Yellow
        
        $completedJobs = 0
        $totalJobs = $jobs.Count
        
        do {
            Start-Sleep -Seconds 2
            $stillRunning = $false
            
            foreach ($jobInfo in $jobs) {
                $job = Get-BitsTransfer -JobId $jobInfo.Job.JobId -ErrorAction SilentlyContinue
                
                if ($job) {
                    switch ($job.JobState) {
                        "Transferred" {
                            Write-Host "Completing download for $($jobInfo.Name)..." -ForegroundColor Cyan
                            try {
                                Complete-BitsTransfer -BitsJob $job
                                Write-Host "$($jobInfo.Name) downloaded successfully" -ForegroundColor Green
                                $completedJobs++
                            } catch {
                                Write-Host "Failed to complete BITS transfer for $($jobInfo.Name): $($_.Exception.Message)" -ForegroundColor Red
                                Remove-BitsTransfer -BitsJob $job
                                throw
                            }
                        }
                        "Transferring" {
                            $stillRunning = $true
                            $progress = [math]::Round(($job.BytesTransferred / $job.BytesTotal) * 100, 1)
                            Write-Host "Downloading $($jobInfo.Name): $progress%" -ForegroundColor Yellow
                        }
                        "Error" {
                            Write-Host "BITS transfer failed for $($jobInfo.Name): $($job.ErrorDescription)" -ForegroundColor Red
                            Remove-BitsTransfer -BitsJob $job
                            throw "BITS transfer failed for $($jobInfo.Name)"
                        }
                        "Cancelled" {
                            Write-Host "BITS transfer was cancelled for $($jobInfo.Name)" -ForegroundColor Red
                            Remove-BitsTransfer -BitsJob $job
                            throw "BITS transfer was cancelled for $($jobInfo.Name)"
                        }
                        default {
                            $stillRunning = $true
                        }
                    }
                }
            }
            
            # Remove completed jobs from monitoring list
            $jobs = $jobs | Where-Object { 
                $job = Get-BitsTransfer -JobId $_.Job.JobId -ErrorAction SilentlyContinue
                $job -and $job.JobState -notin @("Transferred", "Error", "Cancelled")
            }
            
        } while ($stillRunning -and $jobs.Count -gt 0)
        
        Write-Host "`nAll BITS downloads completed!" -ForegroundColor Green
    }
    
    # Verify all files were downloaded
    foreach ($download in $downloads) {
        $filePath = Join-Path $DownloadPath $download.FileName
        if (Test-Path $filePath) {
            $fileSize = (Get-Item $filePath).Length
            Write-Host "✓ $($download.Name): $([math]::Round($fileSize / 1MB, 2)) MB" -ForegroundColor Green
        } else {
            Write-Host "✗ $($download.Name): File not found" -ForegroundColor Red
            throw "Download verification failed for $($download.Name)"
        }
    }
    
    Write-Host "All winget dependencies downloaded successfully using BITS!" -ForegroundColor Green
}

# Function to download winget dependencies
function Download-WingetDependencies {
    param(
        [string]$DownloadPath
    )
    
    Write-Host "Downloading winget dependencies to: $DownloadPath" -ForegroundColor Green
    
    # Create download directory if it doesn't exist
    if (!(Test-Path $DownloadPath)) {
        New-Item -ItemType Directory -Path $DownloadPath -Force | Out-Null
        Write-Host "Created download directory: $DownloadPath" -ForegroundColor Cyan
    }
    
    # Download URLs (latest stable versions)
    $vcLibsUrl = "https://aka.ms/Microsoft.VCLibs.arm64.14.00.Desktop.appx"
    $uiXamlUrl = "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.arm64.appx"
    $wingetUrl = "https://github.com/microsoft/winget-cli/releases/latest/download/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
    
    # Download dependencies
    Write-Host "Downloading Visual C++ Redistributable Libraries..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $vcLibsUrl -OutFile "$DownloadPath\VCLibs.appx" -UseBasicParsing
        Write-Host "VCLibs downloaded successfully" -ForegroundColor Green
    } catch {
        Write-Host "Failed to download VCLibs: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
    
    Write-Host "Downloading UI.Xaml..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $uiXamlUrl -OutFile "$DownloadPath\UIXaml.appx" -UseBasicParsing
        Write-Host "UI.Xaml downloaded successfully" -ForegroundColor Green
    } catch {
        Write-Host "Failed to download UI.Xaml: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
    
    Write-Host "Downloading winget..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $wingetUrl -OutFile "$DownloadPath\winget.msixbundle" -UseBasicParsing
        Write-Host "winget downloaded successfully" -ForegroundColor Green
    } catch {
        Write-Host "Failed to download winget: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
    
    Write-Host "All winget dependencies downloaded successfully!" -ForegroundColor Green
}

# Create the DNS configuration script content (modified to use pre-downloaded files)
$dnsScript = @'
# Set DNS servers for primary network adapter (Windows Sandbox compatible)
# Primary DNS: 1.1.1.1 (Cloudflare)
# Secondary DNS: 8.8.8.8 (Google)

# Set up logging
$logFile = "C:\Users\WDAGUtilityAccount\Desktop\SandboxFiles\sandbox_setup.log"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Function to write to both console and log file
function Write-LogHost {
    param(
        [string]$Message,
        [string]$ForegroundColor = "White",
        [string]$LogLevel = "INFO"
    )
    
    $logEntry = "[$timestamp] [$LogLevel] $Message"
    Write-Host $Message -ForegroundColor $ForegroundColor
    Add-Content -Path $logFile -Value $logEntry -Encoding UTF8
}

# Function to write error to log
function Write-LogError {
    param(
        [string]$Message,
        [System.Management.Automation.ErrorRecord]$ErrorRecord = $null
    )
    
    $logEntry = "[$timestamp] [ERROR] $Message"
    if ($ErrorRecord) {
        $logEntry += " - Exception: $($ErrorRecord.Exception.Message)"
        $logEntry += " - StackTrace: $($ErrorRecord.ScriptStackTrace)"
    }
    
    Write-Host $Message -ForegroundColor Red
    Add-Content -Path $logFile -Value $logEntry -Encoding UTF8
}

# Initialize log file
"=== Windows Sandbox Setup Log ===" | Out-File -FilePath $logFile -Encoding UTF8
"Session started at: $timestamp" | Add-Content -Path $logFile -Encoding UTF8
"=============================================" | Add-Content -Path $logFile -Encoding UTF8

Write-LogHost "Setting DNS servers for Windows Sandbox..." "Green"

# Method 1: Using netsh (most reliable in sandbox)
Write-LogHost "Attempting to set DNS using netsh..." "Yellow"

# Get active network interface name using WMI
try {
    $networkAdapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true -and $_.DefaultIPGateway -ne $null }
    
    if ($networkAdapters) {
        $adapter = $networkAdapters | Select-Object -First 1
        $interfaceName = $adapter.Description
        
        Write-LogHost "Found active adapter: $interfaceName" "Green"
        
        # Set DNS servers using netsh
        $result1 = netsh interface ip set dns name="$interfaceName" static 1.1.1.1
        $result2 = netsh interface ip add dns name="$interfaceName" 8.8.8.8 index=2
        
        Write-LogHost "Primary DNS set: $result1" "Cyan"
        Write-LogHost "Secondary DNS set: $result2" "Cyan"
        
        # Verify using ipconfig
        Write-LogHost "`nCurrent DNS configuration:" "Yellow"
        $dnsInfo = ipconfig /all | Select-String "DNS Servers"
        Write-LogHost $dnsInfo "White"
        
    } else {
        Write-LogError "No active network adapter found via WMI."
    }
    
} catch {
    Write-LogError "WMI method failed" $_
    
    # Fallback: Try setting DNS for all interfaces
    Write-LogHost "Trying fallback method for all interfaces..." "Yellow"
    
    try {
        # Get interface names from netsh
        $interfaces = netsh interface show interface | Select-String "Connected" | ForEach-Object { ($_ -split '\s+')[3] }
        
        foreach ($interface in $interfaces) {
            if ($interface -and $interface.Trim() -ne "") {
                Write-LogHost "Setting DNS for interface: $interface" "Cyan"
                $result1 = netsh interface ip set dns name="$interface" static 1.1.1.1
                $result2 = netsh interface ip add dns name="$interface" 8.8.8.8 index=2
                Write-LogHost "DNS set for $interface - Primary: $result1, Secondary: $result2" "Cyan"
            }
        }
    } catch {
        Write-LogError "Fallback DNS method also failed" $_
    }
}

# Flush DNS cache
Write-LogHost "`nFlushing DNS cache..." "Yellow"
try {
    $flushResult = ipconfig /flushdns
    Write-LogHost "DNS cache flushed successfully" "Green"
} catch {
    Write-LogError "Failed to flush DNS cache" $_
}

Write-LogHost "`nDNS configuration complete!" "Green"

# Wait a moment for DNS to propagate
Write-LogHost "Waiting for DNS to propagate..." "Yellow"
Start-Sleep -Seconds 3

Write-LogHost "`n=== Installing winget (Windows Package Manager) ===" "Magenta"

# Function to install winget from pre-downloaded files
function Install-WingetFromLocal {
    try {
        Write-LogHost "Installing winget from pre-downloaded files..." "Yellow"
        
        # Path to pre-downloaded files in sandbox
        $wingetFilesPath = "C:\Users\WDAGUtilityAccount\Desktop\SandboxFiles\WingetDependencies"
        
        if (!(Test-Path $wingetFilesPath)) {
            Write-LogError "Winget dependencies folder not found: $wingetFilesPath"
            throw "Winget dependencies not available"
        }
        
        # Check if all required files exist
        $requiredFiles = @("VCLibs.appx", "UIXaml.appx", "winget.msixbundle")
        foreach ($file in $requiredFiles) {
            $filePath = Join-Path $wingetFilesPath $file
            if (!(Test-Path $filePath)) {
                Write-LogError "Required file not found: $filePath"
                throw "Missing required file: $file"
            }
            Write-LogHost "Found required file: $file" "Green"
        }
        
        # Install dependencies first
        Write-LogHost "Installing Visual C++ Redistributable Libraries..." "Green"
        try {
            Add-AppxPackage -Path "$wingetFilesPath\VCLibs.appx" -ErrorAction SilentlyContinue
            Write-LogHost "VCLibs installed successfully" "Green"
        } catch {
            Write-LogError "Failed to install VCLibs" $_
        }
        
        Write-LogHost "Installing UI.Xaml..." "Green"
        try {
            Add-AppxPackage -Path "$wingetFilesPath\UIXaml.appx" -ErrorAction SilentlyContinue
            Write-LogHost "UI.Xaml installed successfully" "Green"
        } catch {
            Write-LogError "Failed to install UI.Xaml" $_
        }
        
        # Install winget
        Write-LogHost "Installing winget..." "Green"
        try {
            Add-AppxPackage -Path "$wingetFilesPath\winget.msixbundle" -ErrorAction SilentlyContinue
            Write-LogHost "winget package installed successfully" "Green"
        } catch {
            Write-LogError "Failed to install winget package" $_
        }
        
        # Wait for installation to complete
        Write-LogHost "Waiting for installation to complete..." "Yellow"
        Start-Sleep -Seconds 5
        
        # Verify installation
        Write-LogHost "Verifying winget installation..." "Yellow"
        
        # Add winget to PATH for current session
        $wingetPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps"
        if ($env:PATH -notlike "*$wingetPath*") {
            $env:PATH += ";$wingetPath"
            Write-LogHost "Added winget to PATH: $wingetPath" "Cyan"
        }
        
        # Test winget
        $wingetExe = "$env:LOCALAPPDATA\Microsoft\WindowsApps\winget.exe"
        if (Test-Path $wingetExe) {
            Write-LogHost "winget installed successfully!" "Green"
            try {
                $version = & $wingetExe --version
                Write-LogHost "winget version: $version" "Green"
            } catch {
                Write-LogError "winget executable found but failed to run" $_
            }
            
            # Show some useful commands
            Write-LogHost "`nUseful winget commands:" "Cyan"
            Write-LogHost "  winget search <app-name>    - Search for applications" "White"
            Write-LogHost "  winget install <app-name>   - Install an application" "White"
            Write-LogHost "  winget list                 - List installed applications" "White"
            Write-LogHost "  winget upgrade              - List available upgrades" "White"
            Write-LogHost "  winget upgrade --all        - Upgrade all applications" "White"
            
        } else {
            Write-LogError "winget executable not found at expected location: $wingetExe"
            Write-LogHost "winget installation may not be complete. Try running 'winget' from command prompt." "Yellow"
        }
        
    } catch {
        Write-LogError "Error installing winget from local files" $_
        Write-LogHost "You may need to install winget manually from the Microsoft Store or GitHub releases." "Yellow"
    }
}

# Install winget from pre-downloaded files
Install-WingetFromLocal

Write-LogHost "`n=== Setup Complete ===" "Green"
Write-LogHost "DNS configured and winget installed!" "Green"
Write-LogHost "You can now use winget to install applications." "Yellow"

# Log session completion
$endTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content -Path $logFile -Value "=============================================" -Encoding UTF8
Add-Content -Path $logFile -Value "Session completed at: $endTimestamp" -Encoding UTF8
Add-Content -Path $logFile -Value "Log file location: $logFile" -Encoding UTF8

Write-LogHost "`nLog file saved to: $logFile" "Green"

# Keep the window open
Write-LogHost "`nPress any key to continue..." "Cyan"
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
'@

# Find OneDrive for Business directory (OneDrive - Company Name) for backup copy
$oneDriveBusinessDir = $null
$userProfile = $env:USERPROFILE

# Look for OneDrive - [Company Name] directories
$oneDriveDirs = Get-ChildItem -Path $userProfile -Directory | Where-Object { $_.Name -match "^OneDrive - " }

if ($oneDriveDirs) {
    # Use the first OneDrive for Business directory found
    $oneDriveBusinessDir = $oneDriveDirs[0].FullName
    Write-Host "Found OneDrive for Business: $oneDriveBusinessDir" -ForegroundColor Green
} else {
    # Fallback to regular OneDrive if no business account found
    $regularOneDrive = "$userProfile\OneDrive"
    if (Test-Path $regularOneDrive) {
        $oneDriveBusinessDir = $regularOneDrive
        Write-Host "Using regular OneDrive: $oneDriveBusinessDir" -ForegroundColor Yellow
    } else {
        # Final fallback to Documents folder
        $oneDriveBusinessDir = "$userProfile\Documents"
        Write-Host "OneDrive not found, using Documents folder: $oneDriveBusinessDir" -ForegroundColor Red
    }
}

# Create permanent directory for sandbox scripts on C: drive (primary location)
$scriptsDir = "C:\SandboxScripts"
if (!(Test-Path $scriptsDir)) {
    New-Item -ItemType Directory -Path $scriptsDir -Force | Out-Null
}

# Create backup directory in OneDrive for syncing
$oneDriveScriptsDir = "$oneDriveBusinessDir\Documents\SandboxScripts"
if (!(Test-Path $oneDriveScriptsDir)) {
    New-Item -ItemType Directory -Path $oneDriveScriptsDir -Force | Out-Null
}

# Download winget dependencies to local folder
$wingetDependenciesPath = "$scriptsDir\WingetDependencies"
Write-Host "`nDownloading winget dependencies..." -ForegroundColor Green
try {
    # Use BITS for downloading (change to Download-WingetDependencies for standard method)
    Download-WingetDependencies-BITS -DownloadPath $wingetDependenciesPath
} catch {
    Write-Host "Failed to download winget dependencies: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Continuing without winget installation..." -ForegroundColor Yellow
}

# Save the DNS script to both C: drive (primary) and OneDrive (backup)
$dnsScriptPath = "$scriptsDir\SetDNS_and_Winget.ps1"
$dnsScript | Out-File -FilePath $dnsScriptPath -Encoding UTF8 -Force

$oneDriveDnsScriptPath = "$oneDriveScriptsDir\SetDNS_and_Winget.ps1"
$dnsScript | Out-File -FilePath $oneDriveDnsScriptPath -Encoding UTF8 -Force

# Create a batch file in permanent location
$batchContent = @"
@echo off
echo Running DNS configuration and winget installation script...
powershell.exe -ExecutionPolicy Unrestricted -NoProfile -File "C:\Users\WDAGUtilityAccount\Desktop\SandboxFiles\SetDNS_and_Winget.ps1"
pause
"@

$batchPath = "$scriptsDir\RunDNS_and_Winget.bat"
$batchContent | Out-File -FilePath $batchPath -Encoding ASCII -Force

$oneDriveBatchPath = "$oneDriveScriptsDir\RunDNS_and_Winget.bat"
$batchContent | Out-File -FilePath $oneDriveBatchPath -Encoding ASCII -Force

# Create the Windows Sandbox configuration file
$sandboxConfig = @"
<Configuration>
    <VGpu>Enable</VGpu>
    <Networking>Enable</Networking>
    <MappedFolders>
        <MappedFolder>
            <HostFolder>$scriptsDir</HostFolder>
            <SandboxFolder>C:\Users\WDAGUtilityAccount\Desktop\SandboxFiles</SandboxFolder>
            <ReadOnly>false</ReadOnly>
        </MappedFolder>
    </MappedFolders>
    <LogonCommand>
  <Command>powershell.exe -NoExit -ExecutionPolicy Unrestricted -File "C:\Users\WDAGUtilityAccount\Desktop\SandboxFiles\SetDNS_and_Winget.ps1"</Command>
    </LogonCommand>
</Configuration>
"@

# Save the sandbox configuration to permanent location
$configPath = "$scriptsDir\SandboxWithDNS_and_Winget.wsb"
$sandboxConfig | Out-File -FilePath $configPath -Encoding UTF8 -Force

Write-Host "Created sandbox configuration files:" -ForegroundColor Green
Write-Host "DNS + winget script (C: drive): $dnsScriptPath" -ForegroundColor Yellow
Write-Host "DNS + winget script (OneDrive backup): $oneDriveDnsScriptPath" -ForegroundColor Yellow
Write-Host "Batch file (C: drive): $batchPath" -ForegroundColor Yellow
Write-Host "Batch file (OneDrive backup): $oneDriveBatchPath" -ForegroundColor Yellow
Write-Host "Sandbox config: $configPath" -ForegroundColor Yellow
Write-Host "Winget dependencies: $wingetDependenciesPath" -ForegroundColor Yellow

# Check if Windows Sandbox is available
if (!(Get-WindowsOptionalFeature -Online -FeatureName "Containers-DisposableClientVM" | Where-Object {$_.State -eq "Enabled"})) {
    Write-Host "Windows Sandbox is not enabled. Please enable it first:" -ForegroundColor Red
    Write-Host "1. Open 'Turn Windows features on or off'" -ForegroundColor Yellow
    Write-Host "2. Check 'Windows Sandbox'" -ForegroundColor Yellow
    Write-Host "3. Restart your computer" -ForegroundColor Yellow
    Write-Host "4. Run this script again" -ForegroundColor Yellow
    exit
}

# Launch Windows Sandbox
Write-Host "`nLaunching Windows Sandbox with DNS auto-configuration and winget installation..." -ForegroundColor Green
Write-Host "The DNS script will run automatically when the sandbox starts, followed by winget installation from pre-downloaded files." -ForegroundColor Cyan
Write-Host "You can reuse the .wsb file later: $configPath" -ForegroundColor Cyan

try {
    Start-Process -FilePath "C:\Windows\System32\WindowsSandbox.exe" -ArgumentList $configPath -Wait
} catch {
    Write-Host "Error launching sandbox: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "You can manually run the sandbox configuration file: $configPath" -ForegroundColor Yellow
}

Write-Host "`nSandbox session ended." -ForegroundColor Yellow
Write-Host "Files saved to C: drive: $scriptsDir" -ForegroundColor Green
Write-Host "Backup files saved to OneDrive: $oneDriveScriptsDir" -ForegroundColor Green
Write-Host "Winget dependencies saved to: $wingetDependenciesPath" -ForegroundColor Green
Write-Host "Done!" -ForegroundColor Green
