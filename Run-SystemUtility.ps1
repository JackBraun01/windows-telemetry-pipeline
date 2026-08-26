# --- 1. Paths & Setup ---
$LogFile      = "C:\IT_Projects\Windows telemetry project\SystemHealthLog.csv"
$BackupDir    = "C:\IT_Projects\Windows telemetry project\Backups"
$Timestamp    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

if (-not (Test-Path $BackupDir)) {
    New-Item -ItemType Directory -Path $BackupDir -Force | Out-Null
}

# --- 2. Base System Metrics ---
$RAMFree  = "N/A"
$RAMTotal = "N/A"
$DiskFree = "N/A"
try {
    $OS       = Get-CimInstance Win32_OperatingSystem
    $RAMFree  = [math]::Round($OS.FreePhysicalMemory / 1MB, 2)
    $RAMTotal = [math]::Round($OS.TotalVisibleMemorySize / 1MB, 2)
} catch {}
try {
    $DiskFree = [math]::Round((Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'").FreeSpace / 1GB, 2)
} catch {}

# --- 3. GPU Model Detection ---
$GpuName = "Unknown GPU"
try {
    $Video = Get-CimInstance Win32_VideoController | Select-Object -First 1
    if ($Video -and $Video.Name) { $GpuName = $Video.Name }
} catch {}

# --- 4. Live Battery Telemetry ---
$BatteryPercent = "N/A"
$BatteryStatus  = "AC Powered / Desktop"
try {
    Add-Type -AssemblyName System.Windows.Forms
    $PowerStatus = [System.Windows.Forms.SystemInformation]::PowerStatus
    if ($PowerStatus.BatteryLifePercent -ge 0) {
        $BatteryPercent = "$([int]($PowerStatus.BatteryLifePercent * 100))%"
        if ($PowerStatus.PowerLineStatus -eq 'Online') {
            $BatteryStatus = "AC Powered"
        } else {
            $BatteryStatus = "Discharging"
        }
    }
} catch {}

# --- 5. Process Categorization (Names & Counts) ---
$SystemCriticalNames = @('svchost', 'csrss', 'lsass', 'services', 'smss', 'wininit', 'winlogon', 'explorer', 'dwm', 'fontdrvhost', 'sihost', 'taskhostw', 'MsMpEng')
$AllProcs          = Get-Process -ErrorAction SilentlyContinue

$Tier0_Names = [System.Collections.Generic.HashSet[string]]::new()
$Tier1_Names = [System.Collections.Generic.HashSet[string]]::new()
$Tier2_Names = [System.Collections.Generic.HashSet[string]]::new()

foreach ($Proc in $AllProcs) {
    if ([string]::IsNullOrWhiteSpace($Proc.ProcessName)) { continue }
    
    if ($SystemCriticalNames -contains $Proc.ProcessName -or $Proc.SessionId -eq 0) {
        [void]$Tier0_Names.Add($Proc.ProcessName)
    } elseif (-not [string]::IsNullOrWhiteSpace($Proc.MainWindowTitle)) {
        [void]$Tier1_Names.Add($Proc.ProcessName)
    } else {
        [void]$Tier2_Names.Add($Proc.ProcessName)
       }
}

# --- 6. Tier 3 Threat Detection ---
$Tier3_Threats = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@('mimikatz', 'nc', 'ncat', 'psexec', 'pwdump', 'procdump')
)
$SuspiciousPathPatterns = @('\\AppData\\Local\\Temp\\', '\\Users\\Public\\')

$Tier3_ThreatCount = 0
$Tier3_ThreatNames = [System.Collections.Generic.List[string]]::new()

foreach ($Proc in $AllProcs) {
    if ([string]::IsNullOrWhiteSpace($Proc.ProcessName)) { continue }

    $IsNameMatch = $Tier3_Threats.Contains($Proc.ProcessName)

    $ProcPath = $null
    try { $ProcPath = $Proc.Path } catch {}

    $IsPathMatch = $false
    if ($ProcPath) {
        foreach ($Pattern in $SuspiciousPathPatterns) {
            if ($ProcPath -like "*$Pattern*") { $IsPathMatch = $true; break }
        }
    }

    if ($IsNameMatch -or $IsPathMatch) {
        $Tier3_ThreatCount++
        [void]$Tier3_ThreatNames.Add("$($Proc.ProcessName) (PID $($Proc.Id))")
    }
}

if ($Tier3_ThreatCount -eq 0) {
    $Overall_Security_State = "SECURE: Normal Activity"
} else {
    $Overall_Security_State = "WARNING: Suspicious Binary Detected"
}

# --- 7. Compile & Export Record ---
$LogRecord = [ordered]@{
    Timestamp               = $Timestamp
    Overall_Security_State  = $Overall_Security_State
    Tier0_Essential_Count   = $Tier0_Names.Count
    Tier0_Essential_Apps    = ($Tier0_Names -join ', ')
    Tier1_UserApps_Count    = $Tier1_Names.Count
    Tier1_UserApps_List     = ($Tier1_Names -join ', ')
    Tier2_Background_Count  = $Tier2_Names.Count
    Tier2_Background_List   = ($Tier2_Names -join ', ')
    Tier3_Threat_Count      = $Tier3_ThreatCount
    Tier3_Threat_Names      = ($Tier3_ThreatNames -join ', ')
    GPU_Model               = $GpuName
    Battery_Percent         = $BatteryPercent
    Battery_Status          = $BatteryStatus
    Free_RAM_GB             = $RAMFree
    Total_RAM_GB            = $RAMTotal
    Free_Disk_GB            = $DiskFree
}

$CSVObject = [PSCustomObject]$LogRecord

if (-not (Test-Path $LogFile)) {
    $CSVObject | Export-Csv -Path $LogFile -NoTypeInformation -Encoding utf8
} else {
    $CSVObject | Export-Csv -Path $LogFile -Append -Force -NoTypeInformation -Encoding utf8
}

Write-Host "Diagnostic record logged successfully!" -ForegroundColor Green
