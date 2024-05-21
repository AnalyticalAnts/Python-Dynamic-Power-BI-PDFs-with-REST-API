# PowerShell script to run Python script
$scriptDirectory = "C:\Scripts\PythonScript"
$pythonExecutable = "C:\Users\Username\AppData\Local\Programs\Python\Python39\python.exe"
$scriptPath = "$scriptDirectory\AutomatedReports.py"

# Change to the script directory
Set-Location -Path $scriptDirectory

# Run the Python script
& $pythonExecutable $scriptPath
