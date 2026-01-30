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

# Check Node.js
Write-Host "[1/4] Checking Node.js..."
$nodeVersion = node --version 2>$null
if (-not $nodeVersion) {
    Write-Host "  Node.js not found, installing..."
    
    $nodeUrl = "https://nodejs.org/dist/v20.11.0/node-v20.11.0-x64.msi"
    $nodeMsi = "$env:TEMP\node-installer.msi"
    
    Write-Host "  Downloading Node.js..."
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    $webClient = New-Object System.Net.WebClient
    $webClient.Proxy = New-Object System.Net.WebProxy($Proxy)
    $webClient.DownloadFile($nodeUrl, $nodeMsi)
    
    Write-Host "  Installing Node.js..."
    Start-Process msiexec.exe -ArgumentList "/i", $nodeMsi, "/quiet", "/norestart" -Wait
    
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    
    Write-Host "  Node.js installed!"
    Write-Host ""
    Write-Host "  Please reopen PowerShell and run this script again!"
    Read-Host "Press Enter to exit"
    exit
} else {
    Write-Host "  Node.js found: $nodeVersion"
}

# Install Clawdbot
Write-Host "[2/4] Installing Clawdbot..."
npm config set proxy $Proxy
npm config set https-proxy $Proxy
npm install -g clawdbot
Write-Host "  Clawdbot installed!"

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
clawdbot node run --host $GatewayHost --port $GatewayPort --display-name $NodeName
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

clawdbot node run --host $GatewayHost --port $GatewayPort --display-name $NodeName
