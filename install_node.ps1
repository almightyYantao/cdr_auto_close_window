# Clawdbot Node Windows Install Script
# Run in PowerShell (Admin)

param(
    [string]$GatewayHost = "10.10.12.76",
    [string]$GatewayPort = "18789",
    [string]$NodeName = "Windows-Test-Node",
    [string]$Proxy = "http://10.10.12.76:7897"
)

Write-Host "========================================"
Write-Host "  Clawdbot Node Install Script"
Write-Host "========================================"
Write-Host ""

# Set proxy
Write-Host "[0/4] Setting proxy: $Proxy"
$env:HTTP_PROXY = $Proxy
$env:HTTPS_PROXY = $Proxy
$env:ALL_PROXY = $Proxy
[System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($Proxy)
Write-Host "  Done"

# Check Node.js version
Write-Host "[1/4] Checking Node.js..."
$nodeVersion = node --version 2>$null
$needInstall = $false

if (-not $nodeVersion) {
    Write-Host "  Node.js not found"
    $needInstall = $true
} else {
    # Check version >= 22
    $versionNum = [int]($nodeVersion -replace 'v(\d+)\..*', '$1')
    if ($versionNum -lt 22) {
        Write-Host "  Node.js $nodeVersion is too old (need v22+)"
        $needInstall = $true
    } else {
        Write-Host "  Node.js $nodeVersion OK"
    }
}

if ($needInstall) {
    Write-Host "  Downloading Node.js v22..."
    
    $nodeUrl = "https://nodejs.org/dist/v22.13.0/node-v22.13.0-x64.msi"
    $nodeMsi = "$env:TEMP\node-installer.msi"
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    $webClient = New-Object System.Net.WebClient
    $webClient.Proxy = New-Object System.Net.WebProxy($Proxy)
    $webClient.DownloadFile($nodeUrl, $nodeMsi)
    
    Write-Host "  Installing Node.js v22..."
    Start-Process msiexec.exe -ArgumentList "/i", $nodeMsi, "/quiet", "/norestart" -Wait
    
    # Refresh PATH
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    
    # Add Node.js path manually
    $nodePath = "C:\Program Files\nodejs"
    if (Test-Path $nodePath) {
        $env:Path = "$nodePath;$env:Path"
    }
    
    $newVersion = node --version 2>$null
    if ($newVersion) {
        Write-Host "  Node.js $newVersion installed!"
    } else {
        Write-Host "  Node.js installed! Please reopen PowerShell and run again."
        Read-Host "Press Enter to exit"
        exit
    }
}

# Install Clawdbot
Write-Host "[2/4] Installing Clawdbot..."
npm config set proxy $Proxy
npm config set https-proxy $Proxy

# Clear npm cache first
npm cache clean --force 2>$null

npm install -g clawdbot
Write-Host "  Clawdbot installed!"

# Refresh PATH to find clawdbot
$npmGlobal = npm config get prefix 2>$null
if ($npmGlobal) {
    $env:Path = "$npmGlobal;$env:Path"
}

# Create desktop shortcut
Write-Host "[3/4] Creating startup script..."

$startScript = @"
@echo off
title Clawdbot Node
echo ========================================
echo   Clawdbot Node - $NodeName
echo   Gateway: ${GatewayHost}:${GatewayPort}
echo ========================================
echo.
call clawdbot node run --host $GatewayHost --port $GatewayPort --display-name $NodeName
pause
"@

$startScriptPath = "$env:USERPROFILE\Desktop\Clawdbot-Node.bat"
$startScript | Out-File -FilePath $startScriptPath -Encoding ASCII
Write-Host "  Script created: $startScriptPath"

# Start node
Write-Host "[4/4] Starting node..."
Write-Host ""
Write-Host "========================================"
Write-Host "  Node is running, waiting for approval"
Write-Host ""
Write-Host "  On Gateway run:"
Write-Host "    clawdbot nodes pending"
Write-Host "    clawdbot nodes approve <requestId>"
Write-Host "========================================"
Write-Host ""

# Try to find clawdbot
$clawdbotPath = Get-Command clawdbot -ErrorAction SilentlyContinue
if ($clawdbotPath) {
    & clawdbot node run --host $GatewayHost --port $GatewayPort --display-name $NodeName
} else {
    # Try npm global path
    $npmPrefix = npm config get prefix 2>$null
    $clawdbotExe = "$npmPrefix\clawdbot.cmd"
    if (Test-Path $clawdbotExe) {
        & $clawdbotExe node run --host $GatewayHost --port $GatewayPort --display-name $NodeName
    } else {
        Write-Host "ERROR: clawdbot not found in PATH"
        Write-Host "Please reopen PowerShell and run: clawdbot node run --host $GatewayHost --port $GatewayPort --display-name $NodeName"
        Read-Host "Press Enter to exit"
    }
}
