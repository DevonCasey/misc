# Miscellaneous PowerShell Scripts

Some random scripts I have written for different tasks. Hope something is useful to someone!

## Scripts Overview

### 1. Check-WindowsUpdates.ps1

**Purpose**: Verify if specific Windows KB updates are installed on the system.

**Features**:

- **Smart KB Handling**: Automatically adds "KB" prefix if user enters just numbers
- **Multiple Detection Methods**:

  - Primary: Uses `Get-HotFix` cmdlet for installed updates
  - Fallback: Searches Windows Update history via COM objects

**Usage Examples**:

```powershell
# Interactive mode
.\Check-WindowsUpdates.ps1

# With parameters
.\Check-WindowsUpdates.ps1 -KBNumbers "KB5021653", "KB5040562", "KB5005112", "KB5040430"
```

**Output**: Displays installation status, description, install date, and installer information for each KB.

---

### 2. Begin-Robocopy.ps1

**Purpose**: Enhanced robocopy wrapper with interactive input, comprehensive logging, and error handling.

**Features**:

- **Admin Elevation**: Optional run as admin option
- **Robust Error Handling**: Interprets robocopy exit codes with detailed explanations
- **Performance Monitoring**: Tracks and reports operation duration

**Robocopy Parameters Used**:

- `/MIR` - Mirror directory tree (delete extra files in destination)
- `/copy:DAT` - Copy Data, Attributes, and Timestamps
- `/Z` - Restartable mode for network interruptions
- `/MT:16` - Multi-threaded with 16 threads
- `/R:2` - Retry 2 times on failed copies
- `/W:5` - Wait 5 seconds between retries
- `/TEE` - Output to console and log file
- `/E` - Copy subdirectories including empty ones

**Usage Examples**:

```powershell
# Interactive mode
.\Begin-Robocopy.ps1

# With parameters
.\Begin-Robocopy.ps1 -Source "C:\Source" -Destination "E:\Destination"

# With admin elevation
.\Begin-Robocopy.ps1 -Source "C:\Source" -Destination "E:\Destination" -UseAdmin
```

**Exit Code Interpretation**:

- **0-3**: Success (with varying levels of completeness)
- **4-7**: Warnings (mismatched/additional files detected)
- **8+**: Errors (failed copies)

## Log Locations

All logs are stored in: `C:\Users\$env:USERNAME\Documents\Logs\$ScriptName`

- **Windows Updates**: check-windowsupdates_MM-dd-yyyy_HH-MM-ss.log
- **Robocopy**: `robocopy_MM-dd-yyyy_HH-MM-ss.log`
