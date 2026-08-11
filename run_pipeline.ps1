# Set working directory to project folder
Set-Location -Path $PSScriptRoot

# Force Python to use UTF-8 for stdio/stdout
$env:PYTHONIOENCODING = "utf-8"

# Activate Virtual Environment
& ".\.venv\Scripts\Activate.ps1"

# Execute Python Pipeline with UTF-8 log output
python pipeline.py 2>&1 | Out-File -FilePath "pipeline_execution.log" -Encoding utf8 -Append

# Deactivate Virtual Environment
deactivate