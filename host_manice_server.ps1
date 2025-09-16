# Save As: host_manice_server.ps1
$ollamaPath = "$env:USERPROFILE\AppData\Local\Ollama"
$pythonPath = "$env:USERPROFILE\AppData\Local\Programs\Python\Python310\python.exe"

Write-Host "Starting Ollama Server..."
Start-Process -NoNewWindow -FilePath "$ollamaPath\ollama.exe"

Write-Host "Launching Excel integration server..."
Start-Process -NoNewWindow -FilePath $pythonPath -ArgumentList "manice_excel_backend.py"

# Optional: Auto-start on boot
$StartupFolder = "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
Copy-Item "host_manice_server.ps1" $StartupFolder
Write-Host "Manice AI Server configured for auto-start on Windows boot."
