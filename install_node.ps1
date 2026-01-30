# Clawdbot Node Windows 一键安装脚本 (带代理)
# 用法: 在 PowerShell (管理员) 中运行

param(
    [string]$GatewayHost = "10.10.12.76",
    [string]$GatewayPort = "18789",
    [string]$NodeName = "Windows-Test-Node",
    [string]$Proxy = "http://127.0.0.1:7897"
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Clawdbot Node 一键安装脚本" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 设置代理
Write-Host "[0/4] 设置代理: $Proxy" -ForegroundColor Yellow
$env:HTTP_PROXY = $Proxy
$env:HTTPS_PROXY = $Proxy
$env:ALL_PROXY = $Proxy
[System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($Proxy)
Write-Host "  ✅ 代理已设置" -ForegroundColor Green

# 1. 检查 Node.js
Write-Host "[1/4] 检查 Node.js..." -ForegroundColor Yellow
$nodeVersion = node --version 2>$null
if (-not $nodeVersion) {
    Write-Host "  -> 未安装 Node.js，正在安装..." -ForegroundColor Yellow
    
    # 直接下载 MSI
    $nodeUrl = "https://nodejs.org/dist/v20.11.0/node-v20.11.0-x64.msi"
    $nodeMsi = "$env:TEMP\node-installer.msi"
    
    Write-Host "  -> 下载 Node.js (通过代理)..." -ForegroundColor Gray
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    $webClient = New-Object System.Net.WebClient
    $webClient.Proxy = New-Object System.Net.WebProxy($Proxy)
    $webClient.DownloadFile($nodeUrl, $nodeMsi)
    
    Write-Host "  -> 安装 Node.js..." -ForegroundColor Gray
    Start-Process msiexec.exe -ArgumentList "/i", $nodeMsi, "/quiet", "/norestart" -Wait
    
    # 刷新环境变量
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    
    Write-Host "  ✅ Node.js 安装完成" -ForegroundColor Green
    Write-Host ""
    Write-Host "  ⚠️ 请重新打开 PowerShell 再运行此脚本！" -ForegroundColor Yellow
    Read-Host "按 Enter 退出"
    exit
} else {
    Write-Host "  ✅ Node.js 已安装: $nodeVersion" -ForegroundColor Green
}

# 2. 设置 npm 代理并安装 Clawdbot
Write-Host "[2/4] 安装 Clawdbot..." -ForegroundColor Yellow
npm config set proxy $Proxy
npm config set https-proxy $Proxy
npm install -g clawdbot
Write-Host "  ✅ Clawdbot 安装完成" -ForegroundColor Green

# 3. 创建桌面快捷启动脚本
Write-Host "[3/4] 创建启动脚本..." -ForegroundColor Yellow

$startScript = @"
@echo off
title Clawdbot Node
echo ========================================
echo   Clawdbot Node - $NodeName
echo   Gateway: $GatewayHost`:$GatewayPort
echo ========================================
echo.
clawdbot node run --host $GatewayHost --port $GatewayPort --display-name "$NodeName"
pause
"@

$startScriptPath = "$env:USERPROFILE\Desktop\Clawdbot-Node.bat"
$startScript | Out-File -FilePath $startScriptPath -Encoding ASCII
Write-Host "  ✅ 启动脚本: $startScriptPath" -ForegroundColor Green

# 4. 直接启动节点
Write-Host "[4/4] 启动节点连接..." -ForegroundColor Yellow
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  节点正在运行，等待 Gateway 审批..." -ForegroundColor Cyan
Write-Host "" -ForegroundColor Cyan
Write-Host "  Gateway 端执行:" -ForegroundColor White
Write-Host "    clawdbot nodes pending" -ForegroundColor Yellow
Write-Host "    clawdbot nodes approve <requestId>" -ForegroundColor Yellow
Write-Host "" -ForegroundColor Cyan
Write-Host "  审批后节点即可使用" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 启动节点
& clawdbot node run --host $GatewayHost --port $GatewayPort --display-name $NodeName
