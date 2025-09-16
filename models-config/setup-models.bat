@echo off
title Manice AI Model Setup - Professional Edition
cls

REM Set custom models directory
set "MODELS_DIR=D:\Open_Source_AI_Models"
set "OLLAMA_MODELS=%MODELS_DIR%\Ollama"
setx OLLAMA_MODELS "%OLLAMA_MODELS%" >nul 2>&1

REM Create directories if they don't exist
if not exist "%MODELS_DIR%" mkdir "%MODELS_DIR%"
if not exist "%OLLAMA_MODELS%" mkdir "%OLLAMA_MODELS%"
if not exist "%MODELS_DIR%\Logs" mkdir "%MODELS_DIR%\Logs"

echo ==========================================
echo    🤖 Manice AI Model Setup Pro
echo ==========================================
echo.
echo 📁 Models Directory: %MODELS_DIR%
echo 🗂️ Ollama Models: %OLLAMA_MODELS%
echo.
echo This script will help you set up the AI models
echo required for Manice Excel AI CoPilot:
echo.
echo - 🧠 DeepSeek R1 (Large model for complex tasks)
echo - ⚡ Mistral-7B (Small model for quick operations) 
echo - 🔧 Utility models for specialized tasks
echo.
echo Total Size: ~25-30GB organized in folders
echo.

:MENU
echo What would you like to do?
echo.
echo 1) Install Ollama + Download models (Recommended)
echo 2) Check Ollama installation
echo 3) Download models only
echo 4) Test model connection
echo 5) Configure LM Studio
echo 6) Exit
echo.
set /p choice="Enter your choice (1-6): "

if "%choice%"=="1" goto INSTALL_OLLAMA
if "%choice%"=="2" goto CHECK_OLLAMA
if "%choice%"=="3" goto DOWNLOAD_MODELS
if "%choice%"=="4" goto TEST_MODELS
if "%choice%"=="5" goto SETUP_LM_STUDIO
if "%choice%"=="6" goto EXIT
goto MENU

:INSTALL_OLLAMA
echo.
echo ==========================================
echo    Installing Ollama...
echo ==========================================
echo.

:: Check if Ollama is already installed
ollama version >nul 2>&1
if %errorlevel% == 0 (
    echo ✅ Ollama is already installed!
    ollama version
    echo.
    goto DOWNLOAD_MODELS
)

echo 📥 Downloading Ollama installer...
:: Download Ollama for Windows
powershell -Command "& {Invoke-WebRequest -Uri 'https://ollama.ai/download/OllamaSetup.exe' -OutFile '%TEMP%\OllamaSetup.exe'}"

if exist "%TEMP%\OllamaSetup.exe" (
    echo 🚀 Starting Ollama installation...
    echo Please follow the installation wizard.
    start /wait "%TEMP%\OllamaSetup.exe"
    
    echo.
    echo ⏳ Waiting for Ollama to be ready...
    timeout /t 5 /nobreak >nul
    
    :: Verify installation
    ollama version >nul 2>&1
    if %errorlevel% == 0 (
        echo ✅ Ollama installed successfully!
        del "%TEMP%\OllamaSetup.exe" >nul 2>&1
    ) else (
        echo ❌ Ollama installation failed. Please install manually from https://ollama.ai
        pause
        goto MENU
    )
) else (
    echo ❌ Failed to download Ollama installer
    echo Please download manually from: https://ollama.ai
    pause
    goto MENU
)

goto DOWNLOAD_MODELS

:CHECK_OLLAMA
echo.
echo ==========================================
echo    Checking Ollama Installation...
echo ==========================================
echo.

ollama version >nul 2>&1
if %errorlevel% == 0 (
    echo ✅ Ollama is installed and working!
    ollama version
    echo.
    
    echo 📋 Available models:
    ollama list
    echo.
) else (
    echo ❌ Ollama is not installed or not in PATH
    echo Please install Ollama from: https://ollama.ai
    echo.
)

pause
goto MENU

:DOWNLOAD_MODELS
echo.
echo ==========================================
echo    Downloading AI Models...
echo ==========================================
echo.

:: Check if Ollama is running
ollama version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Ollama is not available
    echo Please install Ollama first (option 1)
    pause
    goto MENU
)

echo.
echo 💾 Total download size: ~25-30GB
echo ⚠️ Ensure you have:
echo   - Sufficient disk space (at least 35GB free)
echo   - Stable internet connection
echo   - Patience (downloads may take 30-60 minutes)
echo.
echo 🧠 Downloading DeepSeek R1 (Large model - ~20GB)...
echo This may take a while depending on your internet connection...

REM Try DeepSeek R1 first
ollama pull deepseek-r1
if %errorlevel% == 0 (
    echo ✅ DeepSeek R1 downloaded successfully!
    goto DOWNLOAD_SMALL_MODEL
)

echo ⚠️ DeepSeek R1 not available, trying CodeLlama 34B...
ollama pull codellama:34b
if %errorlevel% == 0 (
    echo ✅ CodeLlama 34B downloaded as large model alternative!
    goto DOWNLOAD_SMALL_MODEL
)

echo ⚠️ CodeLlama 34B failed, trying Llama2 70B...
ollama pull llama2:70b
if %errorlevel% == 0 (
    echo ✅ Llama2 70B downloaded as large model alternative!
    goto DOWNLOAD_SMALL_MODEL
)

echo ❌ All large model downloads failed. Using smaller alternative...
ollama pull deepseek-coder:33b
echo ✅ Using DeepSeek Coder 33B as large model

:DOWNLOAD_SMALL_MODEL
echo.
echo ⚡ Downloading Mistral-7B (Small model - ~4GB)...
ollama pull mistral:7b

if %errorlevel% == 0 (
    echo ✅ Mistral-7B downloaded successfully!
) else (
    echo ⚠️ Mistral-7B failed, trying alternative...
    ollama pull llama2:7b-chat
    if %errorlevel% == 0 (
        echo ✅ Llama2 7B Chat downloaded as small model alternative!
    ) else (
        echo ❌ Small model download failed. Please check your connection.
    )
)

REM Download additional utility model
echo.
echo 🔧 Downloading utility model (Phi-3 Medium - ~7GB)...
ollama pull phi3:medium
if %errorlevel% == 0 (
    echo ✅ Phi-3 Medium downloaded successfully!
) else (
    echo ⚠️ Phi-3 Medium failed, downloading Phi-3 Mini instead...
    ollama pull phi3:mini
    echo ✅ Phi-3 Mini downloaded as utility model
)

echo.
echo 📋 Current models:
ollama list

echo.
echo ✅ Model setup complete!
echo.

pause
goto MENU

:TEST_MODELS
echo.
echo ==========================================
echo    Testing Model Connections...
echo ==========================================
echo.

:: Test Ollama service
echo 🔍 Testing Ollama service...
curl -s http://localhost:11434/api/tags >nul
if %errorlevel% == 0 (
    echo ✅ Ollama service is running
) else (
    echo ❌ Ollama service is not running
    echo Please start Ollama manually
)

echo.
echo 🧠 Testing DeepSeek R1...
echo Testing | ollama run deepseek-r1 "Respond with 'DeepSeek R1 is working!'"
echo.

echo ⚡ Testing Mistral-7B...
echo Testing | ollama run mistral:7b "Respond with 'Mistral-7B is working!'"
echo.

echo 🤖 Testing Manice AI Server connection...
curl -s http://localhost:8899/health
if %errorlevel% == 0 (
    echo ✅ Manice AI Server is running
) else (
    echo ⚠️ Manice AI Server is not running
    echo Please start it with: cd ai-server && python server.py
)

echo.
pause
goto MENU

:SETUP_LM_STUDIO
echo.
echo ==========================================
echo    LM Studio Configuration
echo ==========================================
echo.
echo LM Studio is an alternative to Ollama for running local models.
echo.
echo 📋 Setup Steps:
echo 1. Download LM Studio from: https://lmstudio.ai
echo 2. Install and launch LM Studio
echo 3. Download these models in LM Studio:
echo    - deepseek-ai/deepseek-r1 (or similar large model)
echo    - mistralai/mistral-7b-instruct (or similar small model)
echo 4. Start the local server in LM Studio (usually port 1234)
echo 5. Update Manice config to use LM Studio:
echo    - Edit ai-server/config.py
echo    - Set preferred_provider = "lm_studio"
echo    - Set lm_studio_url = "http://127.0.0.1:1234"
echo.

echo Would you like to open the LM Studio download page? (Y/N)
set /p open_browser="Enter Y or N: "

if /i "%open_browser%"=="Y" (
    start https://lmstudio.ai
    echo ✅ Opening LM Studio website...
)

echo.
pause
goto MENU

:EXIT
echo.
echo ==========================================
echo    🎉 Setup Complete!
echo ==========================================
echo.
echo To start using Manice:
echo.
echo 1. Start the AI server:
echo    cd ai-server
echo    python server.py
echo.
echo 2. Install the Excel add-in:
echo    cd excel-addin
echo    npm run install-addin
echo.
echo 3. Open Excel and look for the "Manice AI" tab!
echo.
echo 📚 For more information, see the README.md file.
echo.

pause
exit /b 0

:ERROR
echo.
echo ❌ An error occurred during setup.
echo Please check the error message above and try again.
echo.
echo For help, visit: https://github.com/manice-ai/excel-copilot/issues
echo.
pause
exit /b 1