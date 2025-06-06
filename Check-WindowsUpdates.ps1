# Check for specific Windows Updates
param(
    [string[]]$KBNumbers
)

# Enhanced logging function
function Write-Status {
    param([string]$Message, [string]$Level = "INFO")
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$Timestamp] [$Level] $Message" -ForegroundColor $(
        switch ($Level) {
            "ERROR" { "Red" }
            "WARNING" { "Yellow" }
            "SUCCESS" { "Green" }
            default { "White" }
        }
    )
}

function Send-AnyKeyAndExit ($Quit = $false) {
    if ($Quit -eq $true) {
        Write-Host "Press any key to quit."
        [void][System.Console]::ReadKey($true)
        if ($Host.Transcribing) {
            Stop-Transcript | Out-Null
        }
        exit
    }
    else {
        return
    }
}

$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)
$LogDir = "C:\Users\$User\Documents\Logs\$ScriptName"
if (-not (Test-Path $LogDir)) {
    New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
    Write-Status "Created log directory: $LogDir" "INFO"
}
$Timestamp = Get-Date -Format "MM-dd-yyyy_HH-MM-ss"
$LogFile = Join-Path -Path $LogDir -ChildPath "get-windowsupdates-$Timestamp.log"

try {
    Start-Transcript -Path $LogFile -Append
    Write-Status "Started transcript logging: $LogFile" "INFO"
}
catch {
    Write-Status "Failed to start transcript logging: $($_.Exception.Message)" "WARNING"
}

# If no KB numbers provided as parameters, prompt user for input
if (-not $KBNumbers) {
    Write-Status "Windows Update Checker - Interactive Mode" "INFO"
    Write-Host "Enter one kb number per line, press Enter twice when done." -ForegroundColor Yellow
    Write-Host "Example: KB5021653" -ForegroundColor Gray
    Write-Host ""
    $KBNumbers = @()
    do {
        $UserInput = Read-Host "KB Number"
        if ($UserInput -and $UserInput.Trim() -ne "") {
            # Add "KB" prefix if not provided
            if ($UserInput -notmatch "^KB\d+$") {
                if ($UserInput -match "^\d+$") {
                    $UserInput = "KB$UserInput"
                }
            }
            $KBNumbers += $UserInput.Trim()
        }
    } while ($UserInput -and $UserInput.Trim() -ne "")
    
    if ($KBNumbers.Count -eq 0) {
        Write-Status "No KB numbers provided. Exiting." "ERROR"
        Send-AnyKeyAndExit -Quit $true
    }
}

Write-Status "Starting Windows Update check for $($KBNumbers.Count) KB numbers" "INFO"
Write-Host "=" * 50 

foreach ($KB in $KBNumbers) {
    Write-Status "Checking for $KB..." "INFO"
    
    try {
        # Method 1: Check using Get-HotFix
        $Hotfix = Get-HotFix -Id $KB -ErrorAction SilentlyContinue
        
        if ($Hotfix) {
            Write-Status "$KB is INSTALLED" "SUCCESS"
            Write-Host "  Description: $($Hotfix.Description)" -ForegroundColor Gray
            Write-Host "  Installed By: $($Hotfix.InstalledBy)" -ForegroundColor Gray
            Write-Host "  Installed On: $($Hotfix.InstalledOn)" -ForegroundColor Gray
        }
        else {
            # Method 2: Check Windows Update history using COM object
            $Session = New-Object -ComObject Microsoft.Update.Session
            $Searcher = $Session.CreateUpdateSearcher()
            $HistoryCount = $Searcher.GetTotalHistoryCount()
            
            if ($HistoryCount -gt 0) {
                $History = $Searcher.QueryHistory(0, $HistoryCount)
                $Found = $History | Where-Object { $_.Title -like "*$KB*" }
                
                if ($Found) {
                    Write-Status "$KB is INSTALLED (found in update history)" "SUCCESS"
                    Write-Host "  Title: $($Found.Title)" -ForegroundColor Gray
                    Write-Host "  Date: $($Found.Date)" -ForegroundColor Gray
                }
                else {
                    Write-Status "$KB is NOT FOUND" "WARNING"
                }
            }
            else {
                Write-Status "$KB is NOT FOUND (no update history available)" "WARNING"
            }
        }
    }
    catch {
        Write-Status "Error checking $KB - $($_.Exception.Message)" "ERROR"
    }
    
    Write-Host ""
}

Write-Status "Update check completed." "SUCCESS"

# Stop PowerShell transcript logging
try {
    Stop-Transcript
    Write-Status "Stopped transcript logging: $LogFile" "INFO"
}
catch {
    Write-Status "Failed to stop transcript logging: $($_.Exception.Message)" "WARNING"
}

Send-AnyKeyAndExit -Quit $true