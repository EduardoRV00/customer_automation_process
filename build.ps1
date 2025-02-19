$exclude = @("venv", "customer_automation_process.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "customer_automation_process.zip" -Force