<#
.SYNOPSIS
    Automates creation of examination folder structure.
    User presses one key for session and one key for exam type.
#>

# ---------------- Self-Elevation ----------------
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltinRole] "Administrator"))
{
    Write-Host "Restarting script as Administrator..."
    Start-Sleep -Seconds 2
    $argsList = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    Start-Process PowerShell -ArgumentList $argsList -Verb RunAs
    exit
}
# ------------------------------------------------

function Show-Step($msg, $detail = "") {
    Write-Host ""
    Write-Host ">>> $msg" -ForegroundColor Cyan
    if ($detail -ne "") {
        Write-Host "    $detail" -ForegroundColor DarkGray
    }
    Start-Sleep -Seconds 1
}

# --- Determine the base path as script location ---
$BasePath = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }

Clear-Host
Write-Host "=== Exam Folder Automation Tool ===" -ForegroundColor Green
Write-Host ""

# Choose session
Write-Host "Select Session:"
Write-Host "  S = Spring"
Write-Host "  F = Fall"
$sessionChoice = Read-Host "Enter (S/F)"
switch ($sessionChoice.ToUpper()) {
    "S" { $sessionName = "Spring" }
    "F" { $sessionName = "Fall" }
    default { Write-Host "Invalid choice. Use S or F."; exit }
}

# Choose exam type
Write-Host ""
Write-Host "Select Exam Type:"
Write-Host "  M = Mid Term"
Write-Host "  F = Final Term"
$examChoice = Read-Host "Enter (M/F)"
switch ($examChoice.ToUpper()) {
    "M" { $examName = "Mid Term" }
    "F" { $examName = "Final Term" }
    default { Write-Host "Invalid choice. Use M or F."; exit }
}

# Enter year
$year = Read-Host "Enter Year (e.g., 2025)"

# Enter dates (without year)
Write-Host ""
Write-Host "Enter dates as DD-MM (without year). Year will be $year."
$startInput = Read-Host "Enter Start Date (dd-MM)"
$endInput   = Read-Host "Enter End Date (dd-MM)"

# Combine user input with year
$startDate = "$startInput-$year"
$endDate   = "$endInput-$year"

# Parse dates
try {
    $start = [datetime]::ParseExact($startDate,'dd-MM-yyyy',$null)
    $end   = [datetime]::ParseExact($endDate,'dd-MM-yyyy',$null)
}
catch {
    Write-Host "Invalid date format. Use dd-MM."
    exit
}

# Build root folder name
$rootName   = "$sessionName $year $examName Examinations"
$rootFolder = Join-Path $BasePath $rootName
$logFolder  = Join-Path $BasePath "Logs"
$logFile    = Join-Path $logFolder ("ExamFolders-" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".log")

# Prepare folders
if (!(Test-Path $rootFolder)) {
    New-Item -ItemType Directory -Path $rootFolder | Out-Null
}
if (!(Test-Path $logFolder)) {
    New-Item -ItemType Directory -Path $logFolder | Out-Null
}

Start-Transcript -Path $logFile -Append
Show-Step "Starting" "Creating exam folder structure in: $rootFolder"

# Date loop
$current = $start
while ($current -le $end) {
    if ($current.DayOfWeek -ne 'Sunday') {
        $dateFolder = Join-Path $rootFolder ($current.ToString('dd-MM-yyyy'))

        if (!(Test-Path $dateFolder)) {
            New-Item -ItemType Directory -Path $dateFolder | Out-Null
            Show-Step "Created folder" $dateFolder
        }

        # Subfolders
        $subFolders = @(
            "Student Attendance",
            "DB Backup for SQL Database",
            "Get Student for Exam Data File",
            "Exam File"
        )

        foreach ($sf in $subFolders) {
            $subPath = Join-Path $dateFolder $sf
            if (!(Test-Path $subPath)) {
                New-Item -ItemType Directory -Path $subPath | Out-Null
                Write-Host "    + Created $sf"
            }
        }
    }
    else {
        Show-Step "Skipped Sunday" ($current.ToString('dd-MM-yyyy'))
    }
    $current = $current.AddDays(1)
}

Show-Step "All folders created successfully!"
Write-Host "Folders created under: $rootFolder" -ForegroundColor Green
Write-Host "Log file: $logFile"
Stop-Transcript

# Optionally open in Explorer
Start-Process explorer.exe $rootFolder

Read-Host "Press ENTER to exit"
