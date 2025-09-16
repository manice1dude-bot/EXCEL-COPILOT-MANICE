# Manice AI Models Setup - Enhanced Version
# Save As: download_manice_models.ps1

# Configuration
$ModelsDirectory = "D:\Open_Source_AI_Models"
$OllamaModelsPath = "$ModelsDirectory\Ollama"
$LogFile = "$ModelsDirectory\download_log.txt"

# Create directories if they don't exist
Write-Host "🔧 Setting up model directory structure..."
if (-not (Test-Path $ModelsDirectory)) {
    New-Item -ItemType Directory -Path $ModelsDirectory -Force
    Write-Host "✅ Created: $ModelsDirectory"
}

if (-not (Test-Path $OllamaModelsPath)) {
    New-Item -ItemType Directory -Path $OllamaModelsPath -Force
    Write-Host "✅ Created: $OllamaModelsPath"
}

# Set Ollama models directory
$env:OLLAMA_MODELS = $OllamaModelsPath
[Environment]::SetEnvironmentVariable("OLLAMA_MODELS", $OllamaModelsPath, "User")
Write-Host "📁 Set OLLAMA_MODELS path to: $OllamaModelsPath"

Write-Host "🔍 Checking for Ollama installation..."
if (-not (Get-Command ollama -ErrorAction SilentlyContinue)) {
    Write-Host "📥 Ollama not found. Installing Ollama..."
    $OllamaInstaller = "$env:TEMP\OllamaSetup.exe"
    Invoke-WebRequest -Uri "https://ollama.ai/download/OllamaSetup.exe" -OutFile $OllamaInstaller
    Start-Process -FilePath $OllamaInstaller -Wait
    Write-Host "✅ Ollama installation completed"
}

Write-Host "Downloading models..."
Write-Host "Note: Total download size will be approximately 25-30GB"
Write-Host "Ensure you have sufficient disk space and stable internet connection."

# Download large model (DeepSeek R1 or alternative)
Write-Host "Downloading large model: DeepSeek R1 (approx 20GB)..."
try {
    ollama pull deepseek-r1
    Write-Host "✅ DeepSeek R1 downloaded successfully!"
} catch {
    Write-Host "⚠️ DeepSeek R1 not available, trying alternative..."
    ollama pull codellama:34b    # 19GB alternative
    Write-Host "✅ CodeLlama 34B downloaded as large model alternative"
}

# Download small model (Mistral 7B)
Write-Host "Downloading small model: Mistral 7B (approx 4GB)..."
ollama pull mistral:7b
Write-Host "✅ Mistral 7B downloaded successfully!"

# Download additional utility model
Write-Host "Downloading utility model: Llama2 7B Chat (approx 4GB)..."
ollama pull llama2:7b-chat
Write-Host "✅ Llama2 7B Chat downloaded successfully!"

Write-Host "✅ All models downloaded successfully! Total: ~27GB"
