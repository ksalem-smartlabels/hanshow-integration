# ===============================
# CONFIG
# ===============================

$TaskName = "HanshowIntegration"

$PythonPath = "C:\Users\ksalem\AppData\Local\Programs\Python\Python313\python.exe"
$ScriptPath = "C:\Services\HanshowIntegration\excel_to_hanshow_v5.py"
$WorkingDir = "C:\Services\HanshowIntegration\"

# ===============================
# REMOVE EXISTING TASK (if exists)
# ===============================

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Write-Host "Existing task found. Removing..."
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}

# ===============================
# CREATE ACTION
# ===============================

$Action = New-ScheduledTaskAction `
    -Execute $PythonPath `
    -Argument "`"$ScriptPath`"" `
    -WorkingDirectory $WorkingDir

# ===============================
# CREATE TRIGGER (At Startup)
# ===============================

$Trigger = New-ScheduledTaskTrigger -AtStartup

# ===============================
# SETTINGS (Auto Restart on Failure)
# ===============================

$Settings = New-ScheduledTaskSettingsSet `
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `
    -StartWhenAvailable `
    -DontStopIfGoingOnBatteries `
    -AllowStartIfOnBatteries

# ===============================
# REGISTER TASK
# ===============================

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -User "SYSTEM" `
    -RunLevel Highest

Write-Host "HanshowIntegration service installed successfully."
