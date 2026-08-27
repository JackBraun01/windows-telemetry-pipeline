# Automated System Telemetry & Performance Logging Pipeline

An automated background system telemetry pipeline built for personal Windows endpoint monitoring. The system extracts hardware performance metrics, live power telemetry, and categorized process architecture every 5 minutes using **PowerShell** and **WMI/CIM APIs**, storing diagnostic logs locally for analytical ingestion in **Microsoft Excel** via **Power Query**.

---

## Technical Context & Security Rationale

- **Purpose & Scope:** Developed strictly for local personal endpoint monitoring and resource profiling on single-user workstations.
- **Background Execution:** Uses a light VBScript wrapper (`Run-Hidden.vbs`) solely to suppress command prompt window flicker and prevent desktop focus interruption every 5 minutes during active workstation use.

---

## Architecture Overview

- **Trigger:** Windows Task Scheduler invokes the collector every 5 minutes; execution behavior depends on the task's configured logon and power conditions (see Task Scheduler Behavior below).
- **Wrapper (`Run-Hidden.vbs`):** Suppresses GUI console window creation to avoid desktop focus-stealing.
- **Core Engine (`Run-SystemUtility.ps1`):**
  - **Self-locating:** Resolves its own log and backup paths dynamically via `$PSScriptRoot`, so the script works correctly regardless of which folder it's placed in — no hardcoded paths to keep in sync.
  - **Hardware Metrics:** Queries WMI / CIM APIs (RAM usage, C-Drive storage, GPU model).
  - **Power Metrics:** Uses `System.Windows.Forms.SystemInformation` to query AC power status and battery information.
  - **Process Classification:** Evaluates running tasks into 4 distinct tiers via `.NET HashSet` lookups.
- **Data Destination (`SystemHealthLog.csv`):** Appends structured UTF-8 diagnostic records using compatible file-sharing settings, written to the same folder as the script itself.
- **Analytics Layer (`Microsoft Excel`):** Power Query engine ingests and auto-refreshes data model every 5 minutes.

### Script Inspection & Code Review
To inspect the underlying PowerShell implementation, resource collections, and file-sharing behavior, view the core script file directly:
[`Run-SystemUtility.ps1`](./Run-SystemUtility.ps1)

---

## Dashboard Preview

![Dashboard Preview Part 1](dashboard-preview1.png)

![Dashboard Preview Part 2](dashboard-preview2.png)

---

## Key Features & Reliability

- **Resource & Hardware Telemetry:** Collects available memory (GB), total RAM (GB), C-drive free space (GB), and active Display Adapter GPU model via CIM/WMI.
- **Power Telemetry & Desktop Edge Cases:** Queries `System.Windows.Forms.SystemInformation`. For desktop PCs lacking a physical battery, metrics gracefully record as `N/A / AC Powered` without throwing script exceptions.
- **4-Tiered Process Categorization:** Evaluates running processes against `.NET HashSet[string]` lookups for fast memory management:
  - **Tier 0 (Core System):** Session 0 / essential OS processes (`svchost`, `csrss`, `explorer`, `dwm`).
  - **Tier 1 (User Applications):** Active foreground user software (`chrome`, `powershell`, `mmc`).
  - **Tier 2 (Background Services):** Helper tools and unclassified services.
  - **Tier 3 (Threat Indicators):** Proactive detection tier tracking flagged processes or unauthorized binary execution patterns.

### Tier 3 Suspicious-Process Heuristics
The engine evaluates active running processes against a dedicated `.NET HashSet[string]` list of suspicious process names (`$Tier3_Threats`) and path heuristic rules during every 5-minute cycle:
- **Name Matching:** Evaluates active process names against a list of names commonly associated with remote-execution or credential-dumping tools.
- **Path Verification:** Identifies binaries executing out of temporary or non-standard user directories (e.g. `AppData\Local\Temp` or `C:\Users\Public`).
- **State Evaluation:**
  - **`0` Matches:** Flags log output as `SECURE: Normal Activity`.
  - **`>= 1` Matches:** Escalates state to `WARNING: Suspicious Binary Detected` and logs the specific process names and PIDs to the CSV payload for review.

> **Important:** Tier 3 is a heuristic indicator system, not an antivirus or EDR solution. A match means a process warrants investigation — it does not establish that the process is malicious. Legitimate software can share names or run from these paths, and genuinely malicious software can easily avoid both signals entirely. This tier is a lightweight, personal triage aid, not a security control.

- **File-Lock Resilience & Error Handling:** Opens the CSV with compatible file-sharing settings so the collector can generally append data while Excel or another application has the file open for reading. Core hardware queries (RAM, disk) are wrapped in error handling so a single failed query doesn't crash the scheduled run.

---

## Sample CSV Diagnostic Output

```csv
Timestamp,Overall_Security_State,Tier0_Essential_Count,Tier0_Essential_Apps,Tier1_UserApps_Count,Tier1_UserApps_List,Tier2_Background_Count,Tier2_Background_List,Tier3_Threat_Count,Tier3_Threat_Names,GPU_Model,Battery_Percent,Battery_Status,Free_RAM_GB,Total_RAM_GB,Free_Disk_GB
2026-08-26 12:34:55,SECURE: Normal Activity,56,"AggregatorHost, amdfendrsr, csrss, dwm...",7,"chrome, Notepad, SystemSettings",32,"AIXHost, backgroundTaskHost, chrome...",0,,AMD Radeon(TM) 840M Graphics,100%,AC Powered,8.29,23.29,335.84
```

---

## Setup & Deployment

### Prerequisites
- **OS:** Windows 10 / 11 (64-bit)
- **PowerShell:** Version 5.1 or higher
- **Permissions:** Standard User execution (Admin rights are not required for basic WMI/CIM system queries)
- **Visualization:** Microsoft Excel 2016+ with Power Query enabled

### Installation Steps

1. **Clone the repository into any folder you like** — the script no longer requires a specific path:
   ```cmd
   git clone https://github.com/JackBraun01/windows-telemetry-pipeline.git "C:\IT_Projects\Windows telemetry project"
   ```
   `Run-SystemUtility.ps1` automatically detects its own folder and writes `SystemHealthLog.csv` and `Backups\` right next to itself — no editing required inside the script itself.

2. **Point `Run-Hidden.vbs` at the script's actual location.** This is the one path that still needs to be set manually, since Windows needs to be told where to look *before* the script can run and locate itself. Open `Run-Hidden.vbs` and confirm the path matches wherever you cloned the repo:
   ```vbscript
   CreateObject("Wscript.Shell").Run "powershell.exe -ExecutionPolicy Bypass -File ""C:\IT_Projects\Windows telemetry project\Run-SystemUtility.ps1""", 0, False
   ```

3. **Set your execution policy for the session** before testing manually — Windows blocks unsigned scripts by default, and you WILL see an error without this step (see Troubleshooting below):
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```
   This is only needed for manual testing — the automated Task Scheduler run doesn't need it, since `Run-Hidden.vbs` already calls PowerShell with `-ExecutionPolicy Bypass` built in.

   **Security note:** `-Bypass` disables PowerShell's script-signing check entirely for that execution — it does not verify the script hasn't been tampered with. This is an acceptable trade-off here because you are running a script you (or a source you've reviewed) control, on your own machine, for personal monitoring. It would **not** be an acceptable pattern for scripts from an untrusted source, or in a shared/managed environment — the correct alternative there is to sign the script with a code-signing certificate and use `-ExecutionPolicy AllSigned` or `RemoteSigned` instead.

4. **Register the Scheduled Task**, pointing at wherever `Run-Hidden.vbs` actually lives:
   ```powershell
   $action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument "`"C:\IT_Projects\Windows telemetry project\Run-Hidden.vbs`""
   $trigger = New-ScheduledTaskTrigger -At (Get-Date) -Once -RepetitionInterval (New-TimeSpan -Minutes 5)
   Register-ScheduledTask -TaskName "SystemHealthMonitor" -Action $action -Trigger $trigger -Description "5-minute background system diagnostic logger"
   ```

5. **Configure Excel Ingestion:**
   1. Open Microsoft Excel and create a new blank workbook, saved into the same project folder.
   2. Navigate to **Data** → **Get Data** → **From File** → **From Text/CSV**.
   3. Select `SystemHealthLog.csv` from the project folder and click **Transform Data**, then **Close & Load**.
   4. Under **Data** → **Queries & Connections**, right-click the query → **Properties**.
   5. Enable **Refresh every 5 minutes** and **Refresh data when opening the file**.

6. **Test it:**
   ```powershell
   cd "C:\IT_Projects\Windows telemetry project"
   .\Run-SystemUtility.ps1
   ```
   You should see `Diagnostic record logged successfully!` in green, and a new row appear in `SystemHealthLog.csv`.

---

## Troubleshooting

Real issues encountered setting this up, and their fixes:

**"File ... cannot be loaded. The file ... is not digitally signed."**
Windows blocks unsigned scripts by default. Run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` in your PowerShell window before testing manually. This only applies to that one window — you'll need to run it again each time you open a fresh PowerShell session to test.

**Windows Script Host error: "Can not find script file ... Run-Hidden.vbs"**
Task Scheduler's action is pointing at the wrong path. Open the task's Properties → Actions tab → Edit, and check the **"Add arguments"** field specifically (not "Program/script," which should just say `wscript.exe`) — it needs the full, correctly quoted path to wherever `Run-Hidden.vbs` actually lives, e.g. `"C:\IT_Projects\Windows telemetry project\Run-Hidden.vbs"`.

**Script says "logged successfully" but the CSV never shows new data**
Check for a leftover duplicate Scheduled Task from an earlier setup attempt — having two tasks pointed at different script copies (one old, one current) causes exactly this kind of "it worked, but where did it go" confusion. Check Task Scheduler Library for duplicates and disable/delete the old one.

**Two scheduled tasks both firing, causing intermittent random errors**
If you registered the task more than once (e.g. once via PowerShell script, once via the Task Scheduler wizard), you may end up with two separate tasks both running every 5 minutes independently. Check Task Scheduler Library for duplicates and keep only one.

**Excel shows old/stale data after adding a new CSV column**
Power Query's `Source` step can lock in a fixed column count on first setup. If you add a column to the CSV later (e.g. customizing the Tier 3 logic), open Power Query Editor → click the `Source` step → check the formula bar for a `Columns=` parameter and update the number to match your new total column count.

**Renamed test files (e.g. testing threat detection) won't run, or vanish immediately**
This is very likely your own antivirus/EDR software correctly doing its job — renaming an executable to mimic a known suspicious tool is a real detection pattern, and legitimate security software will block or kill it. This is expected behavior, not a bug in this script.

**`git push` rejected: "Updates were rejected because the remote contains work that you do not have locally"**
Someone (including you, via GitHub's web editor) made a change directly on GitHub that your local copy doesn't have. Run `git pull origin main` first to merge those changes down, resolve any file conflicts, then push again.

**A README (or any file) you upload to GitHub doesn't render / shows a placeholder**
GitHub only auto-renders a file named *exactly* `README.md`. Windows often hides file extensions by default, so a rename can silently produce something like `README.md.md` without you noticing. Enable File Explorer → View → "File name extensions" to always see the true filename.

---

## Task Scheduler Behavior & Collector vs. Dashboard Distinction

**The collector (`Run-SystemUtility.ps1`) and the dashboard (Excel) are independent and run on different schedules — this is worth understanding clearly:**

- **The collector runs continuously, regardless of Excel.** Once the Scheduled Task is registered, it fires every 5 minutes as long as Windows is running — whether or not Excel, or any workbook, is open. Data collection does not depend on the dashboard in any way.
- **The dashboard only updates when the workbook itself refreshes.** Excel/Power Query has no visibility into the collector — it simply re-reads whatever is currently in `SystemHealthLog.csv` whenever a refresh is triggered (manually, on open, or on the 5-minute timer *while the workbook is open*). If the workbook is closed, the CSV keeps growing in the background, but the dashboard won't reflect it until you next open the file and refresh.

**Task Scheduler settings relevant to reliability** (checked via the task's Properties in Task Scheduler):

- **Runs whether logged on or not:** By default, a task created with `Register-ScheduledTask` typically runs only while you are logged into Windows. If you want telemetry collected even when logged off (e.g. at a lock screen), you need to explicitly configure the task's **"Run whether user is logged on or not"** option under the General tab — this may prompt for your account password to store credentials securely.
- **Missed runs:** If the PC is asleep, off, or the task's trigger window is missed, Task Scheduler does **not** automatically catch up on missed 5-minute intervals by default — you'll simply see a gap in the CSV's timestamps corresponding to that downtime.
- **Overlapping runs:** Each run typically completes in well under 5 minutes, so overlap is unlikely in normal use. If you do customize this script to do heavier work, be aware Task Scheduler's default behavior allows a new instance to start even if a previous one is still running — check the task's **Settings** tab ("If the task is already running...") if this becomes a concern.
- **On battery / metered connection:** Default Task Scheduler settings may pause tasks on battery power or restrict them on metered connections depending on your "Conditions" tab configuration — worth reviewing if you're monitoring a laptop that's frequently unplugged.

---

## Repository Structure

```
windows-telemetry-pipeline/
├── Run-SystemUtility.ps1      # Core PowerShell collector — hardware, power, process tiering, Tier 3 heuristics
├── Run-Hidden.vbs             # Silent launcher — suppresses the console window Task Scheduler would otherwise show
├── .gitignore                 # Excludes generated data (CSV, xlsx, Backups/) from version control
├── README.md                  # This file
├── dashboard-preview1.png     # Excel dashboard screenshot (part 1)
└── dashboard-preview2.png     # Excel dashboard screenshot (part 2)
```

*Not included in the repo (generated locally on first run):* `SystemHealthLog.csv`, `Backups/`, and your own `.xlsx` dashboard workbook.

---

## Known Limitations

- **Data and log paths are self-resolving** (via `$PSScriptRoot`), but `Run-Hidden.vbs` and the registered Scheduled Task still need to be manually pointed at wherever you place the project — Windows has to be told where to start looking before the script can locate itself. Moving the folder after setup requires updating those two references together.
- **Data files are not included in this repo:** `SystemHealthLog.csv` and any `.xlsx` dashboard file are excluded via `.gitignore`, since they contain live system data from the machine they were generated on. Cloning this repo gives you the pipeline itself, not pre-populated data — you'll start with an empty log.
- **Excel Concurrent Locks:** Power Query maintains a brief read-only lock during data syncs. Streaming appends prevent write collisions, but manually editing the raw CSV during a refresh cycle may cause temporary file access contention.
- **Virtual Machines:** On hypervisors without virtual GPU pass-through, the GPU query defaults to standard display driver strings (e.g., `Basic Display Adapter`).
- **Security software interaction:** Endpoint protection software (Windows Defender, EDR tools, etc.) may quarantine or terminate processes launched from non-standard locations such as `C:\Users\Public\`, which can interfere with manually testing the Tier 3 path-detection rule — this is expected behavior from legitimate security tooling, not a bug in the script.

---

## License
Distributed under the MIT License.
