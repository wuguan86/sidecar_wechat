param (
    [string]$env = "dev"
)

$ErrorActionPreference = "Stop"

Write-Host "Starting build for environment: $env"

# 1. Clean previous builds
if (Test-Path "dist") {
    Remove-Item -Path "dist" -Recurse -Force
}
if (Test-Path "build") {
    Remove-Item -Path "build" -Recurse -Force
}

# 2. Run PyInstaller
Write-Host "Running PyInstaller..."
pyinstaller wechat_bridge.spec

# 3. Copy configuration
$configFile = "config_${env}.yaml"
if (-not (Test-Path $configFile)) {
    Write-Error "Config file not found: $configFile"
}

Write-Host "Copying $configFile to dist/config.yaml..."
Copy-Item -Path $configFile -Destination "dist/config.yaml"

Write-Host "Build complete. Executable and config are in dist/"
