# Script de arranque local — SS Automatización
# Ejecutar: .\arrancar.ps1

$env:MAESTRO_URL = "https://solutionspayroll-my.sharepoint.com/:x:/g/personal/yvega_solutionsandpayroll_com/IQDJGjXNtVSZS5tkb4uo4jNWAeaTuJV_hag9R1Z9-asNBYE?e=pU0mVr"

& "$PSScriptRoot\.venv\Scripts\python.exe" "$PSScriptRoot\server.py"
