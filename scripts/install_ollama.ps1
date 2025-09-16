# Manice AI - Ollama Installation Script
# Simple installer for Ollama on Windows

Write-Host "🚀 Manice AI - Ollama Installation" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "⚠️  This script requires Administrator privileges for installation." -ForegroundColor Yellow
    Write-Host "Please run PowerShell as Administrator and try again." -ForegroundColor Yellow
    pause
    exit 1
}

Write-Host "✅ Running with Administrator privileges" -ForegroundColor Green

# Check if Ollama is already installed
$ollamaPath = Get-Command ollama -ErrorAction SilentlyContinue
if ($ollamaPath) {
    Write-Host "✅ Ollama is already installed!" -ForegroundColor Green
    Write-Host "Location: $($ollamaPath.Source)" -ForegroundColor Gray
    
    # Test if it's working
    try {
        $version = ollama version 2>$null
        Write-Host "✅ Ollama version: $version" -ForegroundColor Green
        
        Write-Host "`n🎉 Ollama is ready! You can now run the 8GB model setup." -ForegroundColor Green
        Write-Host "Next step: python scripts\setup_8gb_optimized.py" -ForegroundColor Cyan
    }
    catch {
        Write-Host "⚠️  Ollama is installed but not working properly." -ForegroundColor Yellow
    }
    
    pause
    exit 0
}

Write-Host "📥 Ollama not found. Starting installation..." -ForegroundColor Yellow

# Create temp directory
$tempDir = "$env:TEMP\ManiceOllamaInstall"
if (-not (Test-Path $tempDir)) {
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
}

Write-Host "📂 Created temp directory: $tempDir" -ForegroundColor Gray

# Download Ollama installer
$ollamaUrl = "https://ollama.ai/download/OllamaSetup.exe"
$installerPath = "$tempDir\OllamaSetup.exe"

Write-Host "📥 Downloading Ollama installer..." -ForegroundColor Yellow
Write-Host "URL: $ollamaUrl" -ForegroundColor Gray

try {
    # Use .NET WebClient for better progress
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile($ollamaUrl, $installerPath)
    $webClient.Dispose()
    
    Write-Host "✅ Download completed!" -ForegroundColor Green
}
catch {
    Write-Host "❌ Download failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please download manually from: https://ollama.ai" -ForegroundColor Yellow
    pause
    exit 1
}

# Verify download
if (-not (Test-Path $installerPath)) {
    Write-Host "❌ Installer not found after download!" -ForegroundColor Red
    Write-Host "Please download manually from: https://ollama.ai" -ForegroundColor Yellow
    pause
    exit 1
}

$fileSize = (Get-Item $installerPath).Length / 1MB
Write-Host "✅ Installer size: $([math]::Round($fileSize, 2)) MB" -ForegroundColor Green

# Run installer
Write-Host "`n🚀 Starting Ollama installation..." -ForegroundColor Cyan
Write-Host "Please follow the installation wizard." -ForegroundColor Yellow
Write-Host "The installer will open in a new window." -ForegroundColor Yellow

try {
    # Start installer and wait for completion
    $process = Start-Process -FilePath $installerPath -Wait -PassThru
    
    if ($process.ExitCode -eq 0) {
        Write-Host "✅ Ollama installation completed successfully!" -ForegroundColor Green
    }
    else {
        Write-Host "⚠️  Installer exited with code: $($process.ExitCode)" -ForegroundColor Yellow
        Write-Host "Installation may have been cancelled or failed." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "❌ Error running installer: $($_.Exception.Message)" -ForegroundColor Red
    pause
    exit 1
}

# Clean up
Write-Host "`n🧹 Cleaning up..." -ForegroundColor Gray
try {
    Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "✅ Temporary files cleaned up" -ForegroundColor Gray
}
catch {
    Write-Host "⚠️  Could not clean up temp files: $tempDir" -ForegroundColor Yellow
}

# Verify installation
Write-Host "`n🔍 Verifying Ollama installation..." -ForegroundColor Cyan
Start-Sleep -Seconds 3  # Give system time to update PATH

# Refresh PATH
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

$ollamaPath = Get-Command ollama -ErrorAction SilentlyContinue
if ($ollamaPath) {
    Write-Host "✅ Ollama found at: $($ollamaPath.Source)" -ForegroundColor Green
    
    try {
        $version = ollama version 2>$null
        Write-Host "✅ Ollama version: $version" -ForegroundColor Green
    }
    catch {
        Write-Host "⚠️  Ollama installed but version check failed" -ForegroundColor Yellow
        Write-Host "This is normal for fresh installations." -ForegroundColor Gray
    }
    
    Write-Host "`n🎉 Installation successful!" -ForegroundColor Green
    Write-Host "=================================" -ForegroundColor Green
    Write-Host "Ollama is now installed and ready to use." -ForegroundColor Green
    Write-Host "`nNext steps:" -ForegroundColor Cyan
    Write-Host "1. Close this PowerShell window" -ForegroundColor White
    Write-Host "2. Open a new PowerShell window (to refresh PATH)" -ForegroundColor White
    Write-Host "3. Run: python scripts\setup_8gb_optimized.py" -ForegroundColor Yellow
    Write-Host "`n🚀 Your 8GB-optimized AI models await!" -ForegroundColor Cyan
}
else {
    Write-Host "❌ Ollama installation verification failed" -ForegroundColor Red
    Write-Host "Please try one of these options:" -ForegroundColor Yellow
    Write-Host "1. Restart PowerShell and check again" -ForegroundColor White
    Write-Host "2. Restart your computer" -ForegroundColor White
    Write-Host "3. Download manually from: https://ollama.ai" -ForegroundColor White
    Write-Host "4. Check if installation was completed in the installer" -ForegroundColor White
}

Write-Host "`nPress any key to exit..." -ForegroundColor Gray
pause