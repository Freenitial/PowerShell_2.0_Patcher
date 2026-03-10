<# :
    @echo off & Title PowerShell 2.0 Patcher & set args=%*
    :: Author : Leo Gillet / Freenitial on GitHub
    :: Version : 1.0

    :: Windows version check
    for /f "tokens=2 delims=[]" %%v in ('ver') do for /f "tokens=2 delims=. " %%m in ("%%v") do set "WINMAJOR=%%m"
    if not defined WINMAJOR set "WINMAJOR=0"
    if %WINMAJOR% GEQ 10 goto :winVersionOk
    set "HTAFILE=%TEMP%\ps2patcher_vercheck.hta"
    > "%HTAFILE%" echo ^<html^>^<head^>^<title^>PowerShell 2.0 Patcher^</title^>
    >> "%HTAFILE%" echo ^<hta:application showintaskbar=yes border=dialog maximizeButton=no minimizeButton=no scroll=no /^>
    >> "%HTAFILE%" echo ^<script language="vbscript"^>
    >> "%HTAFILE%" echo Sub Window_OnLoad : Dim s : s = screen.deviceXDPI / 96 : Dim w : w = CInt(500 * s) : Dim h : h = CInt(240 * s) : self.resizeTo w, h : self.moveTo (screen.availWidth - w) / 2, (screen.availHeight - h) / 2 : End Sub
    >> "%HTAFILE%" echo Sub OpenLink : CreateObject("WScript.Shell").Run "https://github.com/Freenitial/PowerShell-2.0_.NET-4_for_Windows_XP-2003-Vista-2008" : End Sub
    >> "%HTAFILE%" echo ^</script^>^</head^>
    >> "%HTAFILE%" echo ^<body style="font-family:Segoe UI;font-size:9pt;padding:18px;background:#f0f0f0;overflow:hidden"^>
    >> "%HTAFILE%" echo ^<p^>^<b^>This tool requires Windows 10 or later.^</b^>^</p^>
    >> "%HTAFILE%" echo ^<p^>For PowerShell on older Windows versions, visit:^</p^>
    >> "%HTAFILE%" echo ^<p^>^<a href="#" onclick="OpenLink" style="color:#0066cc"^>github.com/Freenitial/PowerShell-2.0_.NET-4_for_Windows_XP-2003-Vista-2008^</a^>^</p^>
    >> "%HTAFILE%" echo ^<center^>^<button onclick="self.close" style="padding:4px 24px;margin-top:8px"^>OK^</button^>^</center^>
    >> "%HTAFILE%" echo ^</body^>^</html^>
    mshta "%HTAFILE%"
    del "%HTAFILE%" 2>nul
    exit /b 1
    :winVersionOk

    :: Elevate to admin
    net session >nul 2>&1 || (powershell -nologo -noprofile -ex bypass "saps '%~f0' -verb runas" & exit /b)

    :: Detect silent/unattended mode to keep console visible
    set "WINMODE=-Window Hidden"
    for %%a in (%args%) do (
        for %%k in (-Unattended -Silent -S) do (
            if /I "%%~a"=="%%k" set "WINMODE="
        )
    )
    
    :: Execute powershell code below
    powershell -NoLogo -NoProfile -Ex Bypass %WINMODE% -Command "$batFile='%~dp0';$sb=[ScriptBlock]::Create([IO.File]::ReadAllText('%~f0'));& $sb @args" %args%
    exit /b %errorlevel%
#>

#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Restores PowerShell 2.0 engine support on Windows 10/11 (x64 + x86).

.DESCRIPTION
    OPERATING MODES (auto-detected) :

    FEATURE AVAILABLE (Windows 10, Windows 11 pre-KB5063878) :
        PS 2.0 is still a Windows Optional Feature. This tool enables it
        alongside .NET Framework 3.5. No binary patching needed.

    FEATURE REMOVED (Windows 11 24H2 build 26100.4946+, 25H2+) :
        Microsoft removed the PS 2.0 feature via KB5063878 (August 2025).
        powershell.exe contains a hardcoded deprecation block that overrides
        -Version 2 to version 3 (PowerShell\3 = CLR4 = PS 5.1).
        This tool patches powershell.exe to remap version 3 -> 1
        (PowerShell\1 = CLR2 = PS 2.0) and suppresses the warning.

    TWO PATCHING STRATEGIES :
        "Duplicate" : creates powershell2.exe alongside the original.
            Original untouched. Survives Windows Update.
        "Replace"   : patches the original powershell.exe in-place.
            Backup created (.bak). Ownership transferred from TrustedInstaller
            via takeown/icacls, then restored. Reverted by any CU.

    DUAL ARCHITECTURE (64-bit OS) :
        System32\...\powershell.exe  = x64 (PE32+, one deprecation block)
        SysWOW64\...\powershell.exe  = x86 (PE32, two deprecation blocks)
        Each patched independently via architecture-specific byte patterns.

    UNATTENDED MODE (-Unattended) :
        Installs everything automatically : .NET 3.5, ps2DLC (if needed),
        duplicate patch for all architectures, desktop shortcuts.

.PARAMETER Unattended
    Runs in non-interactive mode. Installs all prerequisites and applies
    the "Duplicate" patch for both x64 and x86 automatically.

.NOTES
    Reverse-engineered from powershell.exe builds 10.0.26100.5074 / 10.0.26100.7705
    References : KB5065506 (removal notice), KB5063878 (first CU to remove PS2)
    Expects    : ps2DLC.zip alongside this script (when PS2 feature is removed)
#>
param(
    [Alias('Silent','S')]
    [switch]$Unattended
)

# ============================================================================
# C# HELPER FOR ASYNC PROCESS OUTPUT
# ============================================================================
# PowerShell scriptblocks cannot safely execute on .NET ThreadPool threads
# This compiled C# class handles OutputDataReceived on the ThreadPool
# and exposes a ConcurrentQueue<string> for the UI thread to drain.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -TypeDefinition @"
using System;
using System.Collections.Concurrent;
using System.Diagnostics;
public class ProcessOutputCollector {
    public readonly ConcurrentQueue<string> Lines = new ConcurrentQueue<string>();
    public DataReceivedEventHandler GetHandler() {
        return new DataReceivedEventHandler((sender, e) => {
            if (e.Data != null) Lines.Enqueue(e.Data);
        });
    }
}
"@

# ============================================================================
# LOGGING
# ============================================================================

$script:LogDir = [System.IO.Path]::Combine($env:TEMP, 'PowerShell 2.0 Patcher')
if (-not [System.IO.Directory]::Exists($script:LogDir)) {
    [System.IO.Directory]::CreateDirectory($script:LogDir) | Out-Null
}
$script:LogFile = [System.IO.Path]::Combine($script:LogDir, "PS2Patcher_$(Get-Date -Format 'yyyyMMdd').log")
if ([System.IO.File]::Exists($script:LogFile)) {
    [System.IO.File]::AppendAllText($script:LogFile, "`r`n`r`n------------------------------`r`n`r`n")
}
# Rotate old log files (keep last 10)
$logFilesList = [System.IO.Directory]::GetFiles($script:LogDir, '*.log')
if ($logFilesList.Count -gt 10) {
    $sortedLogs = [System.Array]::CreateInstance([System.IO.FileInfo], $logFilesList.Count)
    for ($i = 0; $i -lt $logFilesList.Count; $i++) { $sortedLogs[$i] = [System.IO.FileInfo]::new($logFilesList[$i]) }
    [System.Array]::Sort($sortedLogs, [System.Comparison[System.IO.FileInfo]]{ param($a, $b) $b.LastWriteTimeUtc.CompareTo($a.LastWriteTimeUtc) })
    for ($i = 10; $i -lt $sortedLogs.Count; $i++) { [System.IO.File]::Delete($sortedLogs[$i].FullName) }
}

# Global reference to the GUI log control (set when form is created, $null in unattended)
$script:LogRichTextBox = $null

function Write-Log {
    # Writes a timestamped message to : console (colored), log file, and GUI RichTextBox (if active).
    param(
        [string]$Message,
        [ValidateSet('Info','Warning','Error','Debug')]
        [string]$Level = 'Info'
    )
    $timestamp  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"
    # Console output with color
    switch ($Level) {
        'Error'   { Write-Host $logMessage -ForegroundColor Red }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Debug'   { Write-Host $logMessage -ForegroundColor Gray }
        default   { Write-Host $logMessage -ForegroundColor White }
    }
    # Log file
    try { [System.IO.File]::AppendAllText($script:LogFile, "$logMessage`r`n") }
    catch { Write-Host "Log write failed : $_" -ForegroundColor Red }
    # GUI RichTextBox (if present)
    if ($null -ne $script:LogRichTextBox) {
        $script:LogRichTextBox.AppendText("$Message`r`n")
        $script:LogRichTextBox.ScrollToCaret()
        $script:LogRichTextBox.Refresh()
    }
}

function Write-LogSeparator {
    # Appends a blank line to console, log file and GUI RichTextBox for visual separation.
    Write-Host ''
    try { [System.IO.File]::AppendAllText($script:LogFile, "`r`n") } catch { }
    if ($null -ne $script:LogRichTextBox) {
        $script:LogRichTextBox.AppendText("`r`n")
        $script:LogRichTextBox.ScrollToCaret()
    }
}

Write-Log ('=' * 60)
Write-Log 'PowerShell 2.0 Patcher started'
Write-Log "Running from PowerShell : $($PSVersionTable.PSVersion)"
Write-Log "User : $env:USERNAME | Computer : $env:COMPUTERNAME"
$ntVerKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey('SOFTWARE\Microsoft\Windows NT\CurrentVersion')
$osProductName    = $ntVerKey.GetValue('ProductName')
$osDisplayVersion = $ntVerKey.GetValue('DisplayVersion')
$osBuildNumber    = $ntVerKey.GetValue('CurrentBuildNumber')
$ntVerKey.Close()
if ([int]$osBuildNumber -ge 22000) { $osProductName = $osProductName -replace 'Windows 10', 'Windows 11' }
Write-Log "OS : $osProductName $osDisplayVersion"
Write-Log ('=' * 60)

# ============================================================================
# CONFIGURATION
# ============================================================================

$pathSystem32Ps      = [System.IO.Path]::Combine($env:SystemRoot, 'System32', 'WindowsPowerShell', 'v1.0')
$pathSysWOW64Ps      = [System.IO.Path]::Combine($env:SystemRoot, 'SysWOW64', 'WindowsPowerShell', 'v1.0')
$patchedExeName      = 'powershell2.exe'
$backupExeSuffix     = '.bak'
$registryKeyPath     = 'HKLM:\SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine'
$scriptDirectory     = if ($batFile) { $batFile } else { $PSScriptRoot }
$ps2DlcZipPath       = [System.IO.Path]::Combine($scriptDirectory, 'ps2DLC.zip')
$defaultShortcutDir  = [Environment]::GetFolderPath('Desktop')
$featureNameRoot     = 'MicrosoftWindowsPowerShellV2Root'
$featureNameNetFx3   = 'NetFx3'
$isOs64Bit           = [Environment]::Is64BitOperatingSystem
$notifySoundPath     = [System.IO.Path]::Combine($env:SystemRoot, 'Media', 'Windows Notify.wav')
$script:GuiMode      = (-not $Unattended)
$script:UnattendedErrors = 0

# ============================================================================
# X64 BYTE PATTERNS (PE32+, single deprecation block)
# ============================================================================
# Mask : 1 = must match, 0 = wildcard
# Wildcards : relative CALL displacement (bytes 8-12), RBP stack offset (byte 23)

$x64OriginalBytes = [byte[]]@(
    0x44, 0x8D, 0x42, 0x28,                          # lea r8d, [rdx+0x28]
    0x48, 0x8B, 0x40, 0x08,                          # mov rax, [rax+8]
    0xE8, 0x00, 0x00, 0x00, 0x00,                    # call <rel32>
    0x41, 0xC7, 0x04, 0x24, 0x03, 0x00, 0x00, 0x00, # mov dword ptr [r12], 3
    0xC7, 0x45, 0x00, 0x03, 0x00, 0x00, 0x00         # mov dword ptr [rbp+??], 3
)
$x64OriginalMask  = [byte[]]@(1,1,1,1, 1,1,1,1, 1,0,0,0,0, 1,1,1,1,1,1,1,1, 1,1,0,1,1,1,1)

$x64PatchedBytes  = [byte[]]@(
    0x44, 0x8D, 0x42, 0x28,
    0x48, 0x8B, 0x40, 0x08,
    0x90, 0x90, 0x90, 0x90, 0x90,                    # NOP x5
    0x41, 0xC7, 0x04, 0x24, 0x01, 0x00, 0x00, 0x00, # version -> 1
    0xC7, 0x45, 0x00, 0x01, 0x00, 0x00, 0x00         # copy -> 1
)
$x64PatchedMask   = [byte[]]@(1,1,1,1, 1,1,1,1, 1,1,1,1,1, 1,1,1,1,1,1,1,1, 1,1,0,1,1,1,1)

# ============================================================================
# X86 BYTE PATTERNS (PE32, two deprecation blocks)
# ============================================================================
# The x86 stub uses a different calling convention (push args + indirect call + call esi).
# There are TWO independent deprecation blocks sharing the same 20-byte call sequence.
# Version overrides use varying ModRM encodings (C7 07, C7 45 xx, C7 00) so
# we locate blocks by the call pattern, then forward-scan for mov [reg], 3.

$x86CallOriginalBytes = [byte[]]@(
    0x6A, 0x28,              # push 0x28               ; MUI resource #40
    0x6A, 0x00,              # push 0
    0x50,                    # push eax                 ; resource handle
    0x8B, 0x08,              # mov ecx, [eax]           ; vtable
    0x8B, 0x71, 0x04,        # mov esi, [ecx+4]         ; display function ptr
    0x8B, 0xCE,              # mov ecx, esi             ; thiscall self
    0xFF, 0x15, 0x00, 0x00, 0x00, 0x00, # call [import] ; wildcarded absolute addr
    0xFF, 0xD6               # call esi                 ; display warning
)
$x86CallOriginalMask = [byte[]]@(1,1, 1,1, 1, 1,1, 1,1,1, 1,1, 1,1,0,0,0,0, 1,1)

$x86CallPatchedBytes = [byte[]]@(
    0x6A, 0x28, 0x6A, 0x00, 0x50, 0x8B, 0x08, 0x8B, 0x71, 0x04, 0x8B, 0xCE,
    0x90, 0x90, 0x90, 0x90, 0x90, 0x90, # NOP x6 (was call [import])
    0x90, 0x90                           # NOP x2 (was call esi)
)
$x86CallPatchedMask = [byte[]]@(1,1, 1,1, 1, 1,1, 1,1,1, 1,1, 1,1,1,1,1,1, 1,1)

# ============================================================================
# CORE UTILITY FUNCTIONS
# ============================================================================

function Find-BytePattern {
    # Returns the file offset of the FIRST match, or -1 if not found.
    param([byte[]]$Data, [byte[]]$Pattern, [byte[]]$Mask)
    $len = $Pattern.Length
    $limit = $Data.Length - $len
    $first = $Pattern[0]
    for ($i = 0; $i -le $limit; $i++) {
        if ($Data[$i] -ne $first) { continue }
        $ok = $true
        for ($j = 1; $j -lt $len; $j++) {
            if ($Mask[$j] -eq 1 -and $Data[$i + $j] -ne $Pattern[$j]) { $ok = $false; break }
        }
        if ($ok) { return $i }
    }
    return -1
}

function Find-AllBytePatternMatches {
    # Returns an array of ALL matching offsets. Used for x86 (two blocks).
    param([byte[]]$Data, [byte[]]$Pattern, [byte[]]$Mask)
    $results = [System.Collections.Generic.List[int]]::new()
    $len = $Pattern.Length
    $limit = $Data.Length - $len
    $first = $Pattern[0]
    for ($i = 0; $i -le $limit; $i++) {
        if ($Data[$i] -ne $first) { continue }
        $ok = $true
        for ($j = 1; $j -lt $len; $j++) {
            if ($Mask[$j] -eq 1 -and $Data[$i + $j] -ne $Pattern[$j]) { $ok = $false; break }
        }
        if ($ok) { $results.Add($i) }
    }
    return $results.ToArray()
}

function New-RunAsShortcut {
    # Creates a .lnk with SLDF_RUNAS_USER (bit 0x2000 in LinkFlags at offset 0x14).
    param([string]$Path, [string]$Target, [string]$Arguments, [string]$WorkDir, [string]$Desc, [string]$Icon)
    $ws = New-Object -ComObject WScript.Shell
    $lnk = $ws.CreateShortcut($Path)
    $lnk.TargetPath = $Target; $lnk.Arguments = $Arguments; $lnk.WorkingDirectory = $WorkDir
    $lnk.Description = $Desc; $lnk.IconLocation = $Icon; $lnk.Save()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($lnk) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
    # Patch the .lnk binary to set the RunAsAdmin flag
    $raw = [System.IO.File]::ReadAllBytes($Path)
    $raw[0x15] = $raw[0x15] -bor 0x20
    [System.IO.File]::WriteAllBytes($Path, $raw)
}

function Invoke-NotifySound {
    # Plays the Windows notification sound if the file exists and we are in GUI mode
    if ($script:GuiMode -and [System.IO.File]::Exists($notifySoundPath)) {
        try {
            $player = New-Object System.Media.SoundPlayer($notifySoundPath)
            $player.Play()
            $player.Dispose()
        } catch { }
    }
}

function Set-PrereqIcon {
    # Sets a prerequisite label to OK (green check), Fail (red cross), or Busy (orange hourglass)
    param([System.Windows.Forms.Label]$Label, [ValidateSet('OK','Fail','Busy')][string]$State)
    switch ($State) {
        'OK'   { $Label.Text = [char]0x2714; $Label.ForeColor = [System.Drawing.Color]::Green }
        'Fail' { $Label.Text = [char]0x2716; $Label.ForeColor = [System.Drawing.Color]::Red }
        'Busy' { $Label.Text = [char]0x231B; $Label.ForeColor = [System.Drawing.Color]::Orange }
    }
}

function Disable-AllActionButtons {
    # Disables every action button during an operation to prevent concurrent execution.
    $btnNetFx3.Enabled = $false
    $btnPs2.Enabled = $false
    if ($null -ne $btnDuplicate) { $btnDuplicate.Enabled = $false }
    if ($null -ne $btnReplace)   { $btnReplace.Enabled = $false }
    if ($null -ne $btnOpenPs2)   { $btnOpenPs2.Enabled = $false }
    $btnUninstall.Enabled = $false
    $btnShortcutCreate.Enabled = $false
    $btnPs2Download.Enabled = $false
    $btnRefresh.Enabled = $false
}

function Reset-DismModule {
    # Forces DISM module to clear cached feature state (it caches per-session).
    Remove-Module -Name Dism -Force -ErrorAction SilentlyContinue
    Import-Module -Name Dism -Force -ErrorAction SilentlyContinue
}

# ============================================================================
# DETECTION
# ============================================================================

function Test-FeatureAvailable {
    # Determines if PS 2.0 is still a Windows Optional Feature (queryable via DISM).
    # Returns $true if the feature exists (enabled or disabled), $false if purged from CBS.
    $f = Get-WindowsOptionalFeature -Online -FeatureName $featureNameRoot -ErrorAction SilentlyContinue
    return ($null -ne $f)
}

function Get-ArchPatchState {
    # Returns a hashtable describing all patch artifacts for a given architecture directory.
    # HasDuplicate   : powershell2.exe exists and contains the patched pattern
    # HasInPlace     : powershell.exe itself is patched AND a .bak backup exists
    # HasOrphanPatch : powershell.exe is patched but no backup (manual tampering or lost backup)
    param([string]$PsDir)
    $origPath   = [System.IO.Path]::Combine($PsDir, 'powershell.exe')
    $dupPath    = [System.IO.Path]::Combine($PsDir, $patchedExeName)
    $bakPath    = [System.IO.Path]::Combine($PsDir, "powershell.exe$backupExeSuffix")
    $hasBak     = [System.IO.File]::Exists($bakPath)
    # Check duplicate
    $hasDup = $false
    if ([System.IO.File]::Exists($dupPath)) {
        $dupBytes = [System.IO.File]::ReadAllBytes($dupPath)
        $hasDup = ((Find-BytePattern $dupBytes $x64PatchedBytes $x64PatchedMask) -ge 0) -or
                  ((Find-AllBytePatternMatches $dupBytes $x86CallPatchedBytes $x86CallPatchedMask).Count -gt 0)
    }
    # Check in-place
    $origPatched = $false
    if ([System.IO.File]::Exists($origPath)) {
        $origBytes = [System.IO.File]::ReadAllBytes($origPath)
        $origPatched = ((Find-BytePattern $origBytes $x64PatchedBytes $x64PatchedMask) -ge 0) -or
                       ((Find-AllBytePatternMatches $origBytes $x86CallPatchedBytes $x86CallPatchedMask).Count -gt 0)
    }
    return @{
        HasDuplicate   = $hasDup
        HasInPlace     = ($origPatched -and $hasBak)
        HasOrphanPatch = ($origPatched -and (-not $hasBak))
    }
}

# ============================================================================
# .NET 3.5 INSTALLATION
# ============================================================================

function Get-DismErrorMessage {
    # Returns a human-readable error message for common DISM exit codes.
    param([int]$ExitCode)
    $hex = '0x{0:X}' -f ([uint32]$ExitCode)
    switch ($hex) {
        '0x800F081F' { return "DISM error $hex : source files not found. An internet connection is required to download .NET 3.5, or use a Windows ISO as source." }
        '0x800F0954' { return "DISM error $hex : download blocked by WSUS policy. Configure Group Policy to allow Windows Update or use a Windows ISO as source." }
        '0x80070490' { return "DISM error $hex : component not found in the store. The Windows image may be corrupted (try sfc /scannow)." }
        default      { return "DISM failed with exit code $ExitCode ($hex)." }
    }
}

function Install-NetFx3Feature {
    # Enables .NET 3.5 via dism.exe with real-time progress.
    # GUI mode   : async via ProcessOutputCollector + WinForms Timer, single-line progress.
    # Unattended : synchronous dism.exe, output piped directly to console.
    if (-not $script:GuiMode) {
        Write-Log 'Installing .NET Framework 3.5 via DISM...'
        $script:dismLastWasProgress = $false
        & "$env:SystemRoot\System32\dism.exe" /Online /Enable-Feature /FeatureName:NetFx3 /All /NoRestart 2>&1 | ForEach-Object {
            $line = "$_"
            if ([string]::IsNullOrWhiteSpace($line)) { return }
            $trimmed = $line.Trim()
            $isProgress = ($trimmed.StartsWith('[') -and $trimmed.EndsWith(']'))
            if ($isProgress) {
                # Overwrite the same console line with carriage return
                Write-Host "`r  $($trimmed.PadRight(60))" -NoNewline
                $script:dismLastWasProgress = $true
            } else {
                if ($script:dismLastWasProgress) { Write-Host ''; $script:dismLastWasProgress = $false }
                Write-Log "  $trimmed" 'Debug'
            }
        }
        if ($script:dismLastWasProgress) { Write-Host '' }
        if ($LASTEXITCODE -eq 0) { Write-Log '.NET 3.5 enabled.'; return $true }
        elseif ($LASTEXITCODE -eq 3010) { Write-Log '.NET 3.5 enabled (RESTART REQUIRED).' 'Warning'; return $true }
        else {
            Write-Log (Get-DismErrorMessage $LASTEXITCODE) 'Error'
            return $false
        }
    }
    # GUI mode : async with progress
    Write-Log 'Enabling .NET Framework 3.5 via DISM...'
    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName  = [System.IO.Path]::Combine($env:SystemRoot, 'System32', 'dism.exe')
    $startInfo.Arguments = '/Online /Enable-Feature /FeatureName:NetFx3 /All /NoRestart'
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError  = $true
    $startInfo.CreateNoWindow = $true
    $oemEnc = [System.Text.Encoding]::GetEncoding([System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage)
    $startInfo.StandardOutputEncoding = $oemEnc
    $startInfo.StandardErrorEncoding  = $oemEnc
    try { $script:dismProc = [System.Diagnostics.Process]::Start($startInfo) }
    catch { Write-Log "ERROR launching DISM : $($_.Exception.Message)" 'Error'; return $false }
    $stdoutCol = New-Object ProcessOutputCollector
    $stderrCol = New-Object ProcessOutputCollector
    $script:dismProc.add_OutputDataReceived($stdoutCol.GetHandler())
    $script:dismProc.BeginOutputReadLine()
    $script:dismProc.add_ErrorDataReceived($stderrCol.GetHandler())
    $script:dismProc.BeginErrorReadLine()
    # Track inline progress bar position for single-line overwrite
    $script:dismProgressIdx = -1
    $script:dismDone = $false
    $script:dismOk   = $false
    $uiTimer = New-Object System.Windows.Forms.Timer
    $uiTimer.Interval = 100
    $uiTimer.Add_Tick({
        $line = $null
        while ($stdoutCol.Lines.TryDequeue([ref]$line)) {
            if ([string]::IsNullOrWhiteSpace($line)) { continue }
            $trimmed = $line.Trim()
            # DISM progress lines look like "[====   50.0%   ====]"
            $isProg = ($trimmed.StartsWith('[') -and $trimmed.EndsWith(']'))
            if ($isProg) {
                # Overwrite the previous progress line in the RichTextBox
                if ($script:dismProgressIdx -ge 0) {
                    $script:LogRichTextBox.Select($script:dismProgressIdx, $script:LogRichTextBox.TextLength - $script:dismProgressIdx)
                    $script:LogRichTextBox.SelectedText = "  $trimmed"
                } else {
                    $script:dismProgressIdx = $script:LogRichTextBox.TextLength
                    $script:LogRichTextBox.AppendText("  $trimmed")
                }
                $script:LogRichTextBox.Select($script:LogRichTextBox.TextLength, 0)
            } else {
                # Regular line : finalize any pending progress, then append
                if ($script:dismProgressIdx -ge 0) {
                    $script:LogRichTextBox.AppendText("`r`n")
                    $script:dismProgressIdx = -1
                }
                $script:LogRichTextBox.AppendText("  $trimmed`r`n")
            }
            $script:LogRichTextBox.ScrollToCaret()
        }
        # Check if the DISM process has finished and all output has been drained
        if ($script:dismProc.HasExited -and $stdoutCol.Lines.IsEmpty -and $stderrCol.Lines.IsEmpty) {
            $uiTimer.Stop()
            if ($script:dismProgressIdx -ge 0) {
                $script:LogRichTextBox.AppendText("`r`n")
                $script:dismProgressIdx = -1
            }
            $exitCode = $script:dismProc.ExitCode
            $script:dismProc.Dispose()
            if ($exitCode -eq 0) {
                Write-Log '.NET Framework 3.5 enabled.'
                $script:dismOk = $true
            } elseif ($exitCode -eq 3010) {
                Write-Log '.NET 3.5 enabled but RESTART REQUIRED.' 'Warning'
                [System.Windows.Forms.MessageBox]::Show(
                    ".NET Framework 3.5 installed but a restart is required.`nPlease reboot and run this tool again.",
                    'Restart Required', 'OK', 'Warning') | Out-Null
                $script:dismOk = $true
            } else {
                Write-Log (Get-DismErrorMessage $exitCode) 'Error'
                $script:dismOk = $false
            }
            $script:dismDone = $true
        }
    })
    $uiTimer.Start()
    # Block the caller while pumping the message loop (keeps GUI responsive)
    while (-not $script:dismDone) {
        [System.Windows.Forms.Application]::DoEvents()
        [System.Threading.Thread]::Sleep(15)
    }
    $uiTimer.Dispose()
    return $script:dismOk
}

# ============================================================================
# PS 2.0 FEATURE ENABLE (feature-available systems only)
# ============================================================================

function Install-Ps2Feature {
    Write-Log 'Enabling PowerShell 2.0 feature...'
    try {
        $result = Enable-WindowsOptionalFeature -Online -FeatureName $featureNameRoot -All -NoRestart -ErrorAction Stop
        if ($result.RestartNeeded) {
            Write-Log 'PS 2.0 feature enabled but RESTART REQUIRED.' 'Warning'
            if ($script:GuiMode) {
                [System.Windows.Forms.MessageBox]::Show(
                    "PowerShell 2.0 feature enabled but a restart is required.`nPlease reboot and run this tool again.",
                    'Restart Required', 'OK', 'Warning') | Out-Null
            }
        } else {
            Write-Log 'PS 2.0 feature enabled.'
        }
        return $true
    } catch {
        Write-Log "ERROR : $($_.Exception.Message)" 'Error'
        return $false
    }
}

# ============================================================================
# PS2DLC INSTALLATION (feature-removed systems only)
# ============================================================================

function Install-Ps2DlcPackage {
    # Extracts ps2DLC.zip and installs assemblies for BOTH architectures + registry.
    if (-not [System.IO.File]::Exists($ps2DlcZipPath)) {
        Write-Log "ps2DLC.zip not found at $ps2DlcZipPath" 'Error'
        return $false
    }
    $extractPath = [System.IO.Path]::Combine($env:TEMP, 'ps2DLC_extract')
    if ([System.IO.Directory]::Exists($extractPath)) { [System.IO.Directory]::Delete($extractPath, $true) }
    Write-Log 'Extracting ps2DLC.zip...'
    try { Expand-Archive -Path $ps2DlcZipPath -DestinationPath $extractPath -Force }
    catch { Write-Log "Extraction failed : $($_.Exception.Message)" 'Error'; return $false }
    $ps2Root = [System.IO.Path]::Combine($extractPath, 'ps2DLC')
    if (-not [System.IO.Directory]::Exists($ps2Root)) {
        Write-Log 'ps2DLC subfolder missing in archive.' 'Error'; return $false
    }
    try {
        [System.Reflection.Assembly]::Load('System.EnterpriseServices, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a') | Out-Null
        $gac = New-Object System.EnterpriseServices.Internal.Publish
    } catch {
        Write-Log "GAC API load failed : $($_.Exception.Message)" 'Error'; return $false
    }
    # Install binaries for each architecture present
    $archFolders = @()
    if ($isOs64Bit) { $archFolders += 'amd64' }
    $archFolders += 'x86'
    foreach ($arch in $archFolders) {
        $binPath = [System.IO.Path]::Combine($ps2Root, 'Binaries', $arch)
        if (-not [System.IO.Directory]::Exists($binPath)) { continue }
        Write-Log "Installing binaries ($arch)..."
        foreach ($dll in [System.IO.Directory]::GetFiles($binPath, '*.dll')) {
            Write-Log "  GAC : $([System.IO.Path]::GetFileName($dll))" 'Debug'
            $gac.GacInstall($dll)
        }
        # Localized resources (fallback to en-US)
        $locale = (Get-WinSystemLocale).Name
        $resPath = [System.IO.Path]::Combine($ps2Root, 'ResourceBinaries', $arch, $locale)
        if (-not [System.IO.Directory]::Exists($resPath)) {
            Write-Log "  Locale '$locale' unavailable, using en-US." 'Debug'
            $resPath = [System.IO.Path]::Combine($ps2Root, 'ResourceBinaries', $arch, 'en-US')
        }
        if ([System.IO.Directory]::Exists($resPath)) {
            Write-Log "Installing resources ($arch)..."
            foreach ($dll in [System.IO.Directory]::GetFiles($resPath, '*.dll')) {
                $gac.GacInstall($dll)
            }
        }
    }
    # Registry import
    Write-Log 'Importing registry (PowerShell\1)...'
    $reg1 = [System.IO.Path]::Combine($ps2Root, 'regkeysNew', 'regkeyPowerShell.reg')
    $reg2 = [System.IO.Path]::Combine($ps2Root, 'regkeysNew', 'regkeyPowerShellEngine.reg')
    if ([System.IO.File]::Exists($reg1)) { reg import $reg1 2>&1 | Out-Null }
    if ([System.IO.File]::Exists($reg2)) { reg import $reg2 2>&1 | Out-Null }
    # Cleanup temp extraction folder
    try { [System.IO.Directory]::Delete($extractPath, $true) } catch { }
    # Verify installation
    $gacOk = [System.IO.Directory]::Exists([System.IO.Path]::Combine($env:SystemRoot, 'assembly', 'GAC_MSIL', 'System.Management.Automation'))
    $regOk = $null -ne (Get-ItemProperty -Path $registryKeyPath -ErrorAction SilentlyContinue)
    if ($gacOk -and $regOk) { Write-Log 'ps2DLC installed.'; return $true }
    else { Write-Log 'ps2DLC installation may be incomplete.' 'Warning'; return $false }
}

# ============================================================================
# BINARY PATCHING
# ============================================================================

function Install-BinaryPatch {
    # Patches a single powershell.exe. Auto-detects PE32 vs PE32+ and applies
    # the appropriate pattern (x64 single-block or x86 multi-block).
    # For in-place patching, uses NTFS rename to work around the running-exe lock :
    # the original file is renamed to .bak (NTFS allows renaming open executables),
    # which frees the filename, then patched bytes are written to a brand-new file.
    param(
        [string]$PsDir,       # Directory containing powershell.exe
        [string]$ArchLabel,   # 'x64' or 'x86' for logging
        [bool]$InPlace        # $true = patch original, $false = create duplicate
    )
    $origPath   = [System.IO.Path]::Combine($PsDir, 'powershell.exe')
    $dupPath    = [System.IO.Path]::Combine($PsDir, $patchedExeName)
    $bakPath    = [System.IO.Path]::Combine($PsDir, "powershell.exe$backupExeSuffix")
    $tmpPath    = [System.IO.Path]::Combine($PsDir, 'powershell.exe.patchtmp')
    if (-not [System.IO.File]::Exists($origPath)) {
        Write-Log "$origPath not found." 'Error'; return $false
    }
    $verInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($origPath)
    Write-Log "Source ($ArchLabel) : $($verInfo.FileVersion)"
    $exeBytes = [System.IO.File]::ReadAllBytes($origPath)
    Write-Log "Size : $($exeBytes.Length) bytes" 'Debug'
    # Detect PE type (PE32 vs PE32+)
    $peOffset = [BitConverter]::ToInt32($exeBytes, 0x3C)
    $peMagic  = [BitConverter]::ToUInt16($exeBytes, $peOffset + 4 + 20)
    $isPE32Plus = ($peMagic -eq 0x20B)
    Write-Log "PE type : $(if ($isPE32Plus) {'PE32+ (x64)'} else {'PE32 (x86)'})" 'Debug'
    $patchApplied = $false
    if ($isPE32Plus) {
        # --- x64 : single deprecation block ---
        $offset = Find-BytePattern $exeBytes $x64OriginalBytes $x64OriginalMask
        if ($offset -lt 0) {
            if ((Find-BytePattern $exeBytes $x64PatchedBytes $x64PatchedMask) -ge 0) {
                Write-Log "  Already patched (x64)."
                if (-not $InPlace -and -not [System.IO.File]::Exists($dupPath)) {
                    [System.IO.File]::Copy($origPath, $dupPath, $true)
                    Write-Log "  Copied to $patchedExeName." 'Debug'
                }
                return $true
            }
            Write-Log "x64 deprecation pattern not found (build $($verInfo.FileVersion))." 'Error'; return $false
        }
        Write-Log "  Deprecation block at 0x$($offset.ToString('X'))"
        # NOP call (5 bytes at +8), version 03->01 at +17 and +24
        for ($i = 0; $i -lt 5; $i++) { $exeBytes[$offset + 8 + $i] = 0x90 }
        $exeBytes[$offset + 17] = 0x01
        $exeBytes[$offset + 24] = 0x01
        Write-Log '  NOP warning call | Version remap 3->1'
        # Verify patch was applied correctly
        if ((Find-BytePattern $exeBytes $x64PatchedBytes $x64PatchedMask) -ne $offset) {
            Write-Log 'Patch verification failed.' 'Error'; return $false
        }
        $patchApplied = $true
    } else {
        # --- x86 : two deprecation blocks with forward-scan ---
        $patchedCheck = Find-AllBytePatternMatches $exeBytes $x86CallPatchedBytes $x86CallPatchedMask
        if ($patchedCheck.Count -gt 0) {
            Write-Log "  Already patched (x86, $($patchedCheck.Count) block(s))."
            if (-not $InPlace -and -not [System.IO.File]::Exists($dupPath)) {
                [System.IO.File]::Copy($origPath, $dupPath, $true)
            }
            return $true
        }
        $matches = Find-AllBytePatternMatches $exeBytes $x86CallOriginalBytes $x86CallOriginalMask
        if ($matches.Count -eq 0) {
            Write-Log "x86 deprecation pattern not found (build $($verInfo.FileVersion))." 'Error'; return $false
        }
        Write-Log "  Found $($matches.Count) deprecation block(s)"
        foreach ($blockOff in $matches) {
            Write-Log "  Block at 0x$($blockOff.ToString('X')) :"
            # NOP call [import] (6 bytes at +12) and call esi (2 bytes at +18)
            for ($i = 0; $i -lt 6; $i++) { $exeBytes[$blockOff + 12 + $i] = 0x90 }
            $exeBytes[$blockOff + 18] = 0x90
            $exeBytes[$blockOff + 19] = 0x90
            Write-Log '    NOP warning calls (8 bytes)' 'Debug'
            # Forward-scan for mov [reg], 3 (C7 ModRM [disp8?] 03 00 00 00)
            for ($scan = 20; $scan -lt 50; $scan++) {
                if (($blockOff + $scan + 6) -ge $exeBytes.Length) { break }
                if ($exeBytes[$blockOff + $scan] -ne 0xC7) { continue }
                $modRM = $exeBytes[$blockOff + $scan + 1]
                $mod = ($modRM -shr 6) -band 3
                $immOff = -1
                if ($mod -eq 0) { $immOff = 2 }      # no displacement
                elseif ($mod -eq 1) { $immOff = 3 }   # 8-bit displacement
                else { continue }
                $pos = $blockOff + $scan + $immOff
                if ($exeBytes[$pos] -eq 0x03 -and $exeBytes[$pos+1] -eq 0x00 -and
                    $exeBytes[$pos+2] -eq 0x00 -and $exeBytes[$pos+3] -eq 0x00) {
                    $exeBytes[$pos] = 0x01
                    Write-Log "    Version remap 3->1 at +$scan" 'Debug'
                }
            }
        }
        $patchApplied = $true
    }
    if (-not $patchApplied) { return $false }
    # ---- Write the patched binary to disk ----
    if ($InPlace) {
        # In-place patching strategy :
        # powershell.exe is locked by the current process (we are running from it).
        # NTFS allows renaming a file whose executable image is mapped, so we :
        #   1. Save ACL, takeown/icacls to get write permission
        #   2. RENAME the locked original -> .bak (frees the filename)
        #   3. Write patched bytes to a BRAND-NEW file with the original name
        #   4. Restore ACL on the new file
        Write-Log '  Saving ACL...' 'Debug'
        $savedAcl = $null
        try { $savedAcl = [System.IO.File]::GetAccessControl($origPath, [System.Security.AccessControl.AccessControlSections]::All) } catch { }
        Write-Log '  Taking ownership...' 'Debug'
        # Resolve Administrators group name from SID (locale-independent)
        $adminSid = [System.Security.Principal.SecurityIdentifier]::new('S-1-5-32-544')
        $adminAccount = $adminSid.Translate([System.Security.Principal.NTAccount]).Value
        $takeownOut = & takeown /f $origPath 2>&1
        Write-Log "  takeown : $takeownOut" 'Debug'
        $icaclsOut = & icacls $origPath /grant "${adminAccount}:F" 2>&1
        Write-Log "  icacls : $icaclsOut" 'Debug'
        # Fallback : take ownership via .NET if icacls failed
        try {
            $acl = [System.IO.File]::GetAccessControl($origPath)
            $acl.SetOwner($adminSid)
            $fullControl = [System.Security.AccessControl.FileSystemAccessRule]::new(
                $adminSid,
                [System.Security.AccessControl.FileSystemRights]::FullControl,
                [System.Security.AccessControl.AccessControlType]::Allow)
            $acl.SetAccessRule($fullControl)
            [System.IO.File]::SetAccessControl($origPath, $acl)
            Write-Log '  .NET ACL override applied.' 'Debug'
        } catch {
            Write-Log "  .NET ACL override failed : $($_.Exception.Message)" 'Warning'
        }
        try {
            if (-not [System.IO.File]::Exists($bakPath)) {
                # Rename the locked original to .bak (NTFS allows renaming mapped executables)
                [System.IO.File]::Move($origPath, $bakPath)
                Write-Log "  Backup created : $bakPath"
            } else {
                Write-Log "  Backup already exists : $bakPath"
                # Backup already exists (retry scenario) : rename to a temp name
                if ([System.IO.File]::Exists($tmpPath)) {
                    try { [System.IO.File]::Delete($tmpPath) } catch { }
                }
                [System.IO.File]::Move($origPath, $tmpPath)
                Write-Log '  Renamed active exe to .patchtmp' 'Debug'
            }
        } catch {
            Write-Log "ERROR renaming exe : $($_.Exception.Message)" 'Error'
            if ([System.IO.File]::Exists($bakPath)) {
                Write-Log "  Backup is safe at : $bakPath" 'Warning'
            }
            return $false
        }
        try {
            # Write patched bytes to a new file (filename is now free)
            [System.IO.File]::WriteAllBytes($origPath, $exeBytes)
            Write-Log "  Patched in-place : powershell.exe"
            Write-Log "  Original backed up at : $bakPath"
        } catch {
            Write-Log "ERROR writing patched file : $($_.Exception.Message)" 'Error'
            Write-Log "  Attempting automatic restore from backup : $bakPath" 'Warning'
            try {
                [System.IO.File]::Move($bakPath, $origPath)
                Write-Log '  Restore successful.' 'Warning'
            } catch {
                Write-Log "  Automatic restore failed! Backup remains at : $bakPath" 'Error'
            }
            if ([System.IO.File]::Exists($tmpPath)) {
                try { [System.IO.File]::Delete($tmpPath) } catch { }
            }
            return $false
        }
        # Restore original TrustedInstaller ACL on the new file
        if ($null -ne $savedAcl) {
            try { [System.IO.File]::SetAccessControl($origPath, $savedAcl); Write-Log '  ACL restored.' 'Debug' }
            catch { Write-Log '  Could not restore original ACL.' 'Warning' }
        }
        # Cleanup leftover .patchtmp from a previous rename if it exists
        if ([System.IO.File]::Exists($tmpPath)) {
            try { [System.IO.File]::Delete($tmpPath) } catch { }
        }
    } else {
        # Duplicate : write patched copy alongside original
        try {
            [System.IO.File]::WriteAllBytes($dupPath, $exeBytes)
            Write-Log "  Written : $patchedExeName"
        } catch {
            Write-Log "ERROR : $($_.Exception.Message)" 'Error'; return $false
        }
    }
    return $true
}

# ============================================================================
# UNINSTALL
# ============================================================================

function Uninstall-BinaryPatch {
    # Removes ALL patch artifacts for a single architecture directory.
    # Handles : duplicate only, in-place only, both, orphan (no backup).
    param([string]$PsDir, [string]$ArchLabel)
    $origPath = [System.IO.Path]::Combine($PsDir, 'powershell.exe')
    $dupPath  = [System.IO.Path]::Combine($PsDir, $patchedExeName)
    $bakPath  = [System.IO.Path]::Combine($PsDir, "powershell.exe$backupExeSuffix")
    $tmpPath  = [System.IO.Path]::Combine($PsDir, 'powershell.exe.patchtmp')
    $state = Get-ArchPatchState -PsDir $PsDir
    $did = $false
    # Remove duplicate
    if ($state.HasDuplicate) {
        try { [System.IO.File]::Delete($dupPath); Write-Log "  Removed ($ArchLabel) : $patchedExeName"; $did = $true }
        catch { Write-Log "ERROR removing duplicate ($ArchLabel) : $($_.Exception.Message)" 'Error' }
    }
    # Restore original from backup (uses rename to work around running-exe lock)
    if ($state.HasInPlace) {
        Write-Log "  Restoring original ($ArchLabel) from $bakPath..."
        $savedAcl = $null
        try { $savedAcl = [System.IO.File]::GetAccessControl($origPath, [System.Security.AccessControl.AccessControlSections]::All) } catch { }
        # Take ownership via SID (locale-independent)
        $adminSid = [System.Security.Principal.SecurityIdentifier]::new('S-1-5-32-544')
        $adminAccount = $adminSid.Translate([System.Security.Principal.NTAccount]).Value
        & takeown /f $origPath 2>&1 | Out-Null
        & icacls $origPath /grant "${adminAccount}:F" 2>&1 | Out-Null
        try {
            # Rename the patched (locked) exe out of the way
            if ([System.IO.File]::Exists($tmpPath)) {
                try { [System.IO.File]::Delete($tmpPath) } catch { }
            }
            [System.IO.File]::Move($origPath, $tmpPath)
            # Restore from backup
            [System.IO.File]::Move($bakPath, $origPath)
            Write-Log '  Restored from backup.'
            $did = $true
            if ($null -ne $savedAcl) {
                try { [System.IO.File]::SetAccessControl($origPath, $savedAcl) } catch { }
            }
            # Cleanup temp file
            try { [System.IO.File]::Delete($tmpPath) } catch { }
        } catch { Write-Log "ERROR restoring ($ArchLabel) : $($_.Exception.Message)" 'Error' }
    }
    if ($state.HasOrphanPatch) {
        Write-Log "  WARNING ($ArchLabel) : powershell.exe is patched but no backup exists!" 'Warning'
        Write-Log '  A Windows Update (CU) will replace the file.' 'Warning'
    }
    if (-not $did -and -not $state.HasOrphanPatch) { Write-Log "  Nothing to remove ($ArchLabel)." }
}

function Uninstall-Shortcuts {
    # Removes PS2 shortcut files from a directory (tries multiple naming conventions)
    param([string]$Dir)
    foreach ($name in @('PowerShell 2.0.lnk', 'PowerShell 2.0 (x64).lnk', 'PowerShell 2.0 (x86).lnk')) {
        $lnkPath = [System.IO.Path]::Combine($Dir, $name)
        if ([System.IO.File]::Exists($lnkPath)) {
            [System.IO.File]::Delete($lnkPath)
            Write-Log "  Removed shortcut : $name"
        }
    }
}

# ============================================================================
# OPEN PS 2.0 HELPER
# ============================================================================

function Get-Ps2LaunchInfo {
    # Determines the best executable and arguments to launch PS 2.0.
    # Returns a hashtable with ExePath and InnerExeName, or $null if nothing is available.
    # Priority : x64 Replace > x64 Duplicate > x86 Replace > x86 Duplicate.
    # On feature-available systems, just use the standard powershell.exe.
    if ($script:FeatureAvailable) {
        # Feature-available : standard powershell.exe with -Version 2
        # Prefer x64 on 64-bit OS
        $targetDir = $pathSystem32Ps
        if (-not [System.IO.Directory]::Exists($targetDir)) { $targetDir = $pathSystem32Ps }
        return @{
            ExePath      = [System.IO.Path]::Combine($targetDir, 'powershell.exe')
            InnerExeName = 'powershell'
        }
    }
    # Feature-removed : check patch states in priority order
    $candidates = @(
        @{ Dir = $pathSystem32Ps; Label = if ($isOs64Bit) { 'x64' } else { 'x86' } }
    )
    if ($isOs64Bit -and [System.IO.Directory]::Exists($pathSysWOW64Ps)) {
        $candidates += @{ Dir = $pathSysWOW64Ps; Label = 'x86' }
    }
    # Pass 1 : prefer in-place (Replace) for each arch (x64 first)
    foreach ($c in $candidates) {
        $st = Get-ArchPatchState -PsDir $c.Dir
        if ($st.HasInPlace -or $st.HasOrphanPatch) {
            return @{
                ExePath      = [System.IO.Path]::Combine($c.Dir, 'powershell.exe')
                InnerExeName = 'powershell'
            }
        }
    }
    # Pass 2 : fall back to duplicate for each arch (x64 first)
    foreach ($c in $candidates) {
        $st = Get-ArchPatchState -PsDir $c.Dir
        if ($st.HasDuplicate) {
            return @{
                ExePath      = [System.IO.Path]::Combine($c.Dir, $patchedExeName)
                InnerExeName = [System.IO.Path]::GetFileNameWithoutExtension($patchedExeName)
            }
        }
    }
    return $null
}

# ============================================================================
# MODE DETECTION
# ============================================================================

try {
    $script:FeatureAvailable = Test-FeatureAvailable
} catch {
    Write-Log "WARNING : could not query PS 2.0 feature state (CBS/DISM error). Assuming removed." 'Warning'
    Write-Log "  $($_.Exception.Message)" 'Debug'
    $script:FeatureAvailable = $false
}
Write-Log "PS 2.0 Windows Feature available : $($script:FeatureAvailable)"

# ============================================================================
# UNATTENDED MODE
# ============================================================================

if ($Unattended) {
    Write-Log '--- UNATTENDED MODE ---'
    $script:UnattendedErrors = 0
    # On feature-removed systems, verify ps2DLC.zip availability FIRST (fail fast)
    if (-not $script:FeatureAvailable) {
        if (-not [System.IO.File]::Exists($ps2DlcZipPath)) {
            Write-Log 'ps2DLC.zip not found, attempting download from Microsoft...'
            $downloadUrl = 'https://download.microsoft.com/download/2b37839b-e146-465a-a78c-c9066609c553/ps2DLC.zip'
            $tempDownloadPath = [System.IO.Path]::Combine($scriptDirectory, 'ps2DLC.zip.downloading')
            try {
                # Manual streaming with progress reporting (synchronous, no threading issues)
                $request = [System.Net.HttpWebRequest]::Create($downloadUrl)
                $request.Timeout = 30000
                $response = $request.GetResponse()
                $totalBytes = $response.ContentLength
                $responseStream = $response.GetResponseStream()
                $fileStream = [System.IO.File]::Create($tempDownloadPath)
                $buffer = New-Object byte[] 65536
                $bytesRead = 0
                $totalRead = 0
                $lastPct = -1
                while (($bytesRead = $responseStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
                    $fileStream.Write($buffer, 0, $bytesRead)
                    $totalRead += $bytesRead
                    if ($totalBytes -gt 0) {
                        $pct = [int](($totalRead / $totalBytes) * 100)
                        if ($pct -ne $lastPct) {
                            $sizeMB = '{0:N1}' -f ($totalRead / 1MB)
                            $totalMB = '{0:N1}' -f ($totalBytes / 1MB)
                            Write-Host "`r  Downloading : $pct% ($sizeMB / $totalMB MB)   " -NoNewline
                            $lastPct = $pct
                        }
                    }
                }
                $fileStream.Close()
                $responseStream.Close()
                $response.Close()
                Write-Host ''
                if ([System.IO.File]::Exists($ps2DlcZipPath)) { [System.IO.File]::Delete($ps2DlcZipPath) }
                [System.IO.File]::Move($tempDownloadPath, $ps2DlcZipPath)
                Write-Log "ps2DLC.zip downloaded to $ps2DlcZipPath"
            } catch {
                Write-Log "Download failed : $($_.Exception.Message)" 'Error'
                # Cleanup partial/corrupt temp file
                foreach ($disposable in @($fileStream, $responseStream, $response)) {
                    if ($null -ne $disposable) { try { $disposable.Close() } catch { } }
                }
                if ([System.IO.File]::Exists($tempDownloadPath)) {
                    try { [System.IO.File]::Delete($tempDownloadPath) } catch { }
                }
                Write-Log 'Cannot proceed without ps2DLC.zip. Place it next to the script or ensure internet access.' 'Error'
                exit 1
            }
        } else {
            Write-Log "ps2DLC.zip found at $ps2DlcZipPath"
        }
        Write-LogSeparator
    }
    # Step 1 : .NET 3.5
    Reset-DismModule
    $netFx3 = Get-WindowsOptionalFeature -Online -FeatureName $featureNameNetFx3 -ErrorAction SilentlyContinue
    if (-not ($netFx3 -and $netFx3.State -eq 'Enabled')) {
        if (-not (Install-NetFx3Feature)) { $script:UnattendedErrors++ }
    } else { Write-Log '.NET 3.5 already enabled.' }
    Write-LogSeparator
    if ($script:FeatureAvailable) {
        # Feature-available : enable the PS2 feature
        $ps2f = Get-WindowsOptionalFeature -Online -FeatureName $featureNameRoot -ErrorAction SilentlyContinue
        if (-not ($ps2f -and $ps2f.State -eq 'Enabled')) {
            if (-not (Install-Ps2Feature)) { $script:UnattendedErrors++ }
        } else { Write-Log 'PS 2.0 feature already enabled.' }
        # Shortcuts for feature-available (standard powershell.exe -Version 2)
        Write-LogSeparator
        foreach ($entry in @(
            @{ Dir = $pathSystem32Ps; Suffix = if ($isOs64Bit) {' (x64)'} else {''} },
            @{ Dir = $pathSysWOW64Ps; Suffix = ' (x86)' }
        )) {
            if (-not $isOs64Bit -and $entry.Suffix -eq ' (x86)') { continue }
            $lnkPath   = [System.IO.Path]::Combine($defaultShortcutDir, "PowerShell 2.0$($entry.Suffix).lnk")
            $targetExe = [System.IO.Path]::Combine($entry.Dir, 'powershell.exe')
            if ([System.IO.File]::Exists($targetExe) -and -not [System.IO.File]::Exists($lnkPath)) {
                New-RunAsShortcut -Path $lnkPath -Target $targetExe -Arguments '-Version 2 -NoExit' -WorkDir '%USERPROFILE%' -Desc "PowerShell 2.0$($entry.Suffix)" -Icon "$targetExe,0"
                Write-Log "Shortcut : $lnkPath"
            }
        }
    } else {
        # Feature-removed : install ps2DLC + duplicate both archs
        $gacOk = [System.IO.Directory]::Exists([System.IO.Path]::Combine($env:SystemRoot, 'assembly', 'GAC_MSIL', 'System.Management.Automation'))
        $regOk = $null -ne (Get-ItemProperty -Path $registryKeyPath -ErrorAction SilentlyContinue)
        $dlcReady = ($gacOk -and $regOk)
        if (-not $dlcReady) {
            $dlcReady = Install-Ps2DlcPackage
            if (-not $dlcReady) { $script:UnattendedErrors++ }
        } else { Write-Log 'ps2DLC already installed.' }
        if ($dlcReady) {
            Write-LogSeparator
            # Patch x64
            $stateX64 = Get-ArchPatchState -PsDir $pathSystem32Ps
            if (-not $stateX64.HasDuplicate) {
                Write-Log '--- Patching x64 ---'
                if (-not (Install-BinaryPatch -PsDir $pathSystem32Ps -ArchLabel 'x64' -InPlace $false)) { $script:UnattendedErrors++ }
            } else { Write-Log 'x64 duplicate already exists.' }
            # Patch x86 (if 64-bit OS)
            if ($isOs64Bit) {
                $stateX86 = Get-ArchPatchState -PsDir $pathSysWOW64Ps
                if (-not $stateX86.HasDuplicate) {
                    Write-Log '--- Patching x86 ---'
                    if (-not (Install-BinaryPatch -PsDir $pathSysWOW64Ps -ArchLabel 'x86' -InPlace $false)) { $script:UnattendedErrors++ }
                } else { Write-Log 'x86 duplicate already exists.' }
            }
            Write-LogSeparator
            # Shortcuts for feature-removed (patched duplicate)
            foreach ($entry in @(
                @{ Dir = $pathSystem32Ps; Suffix = if ($isOs64Bit) {' (x64)'} else {''} },
                @{ Dir = $pathSysWOW64Ps; Suffix = ' (x86)' }
            )) {
                if (-not $isOs64Bit -and $entry.Suffix -eq ' (x86)') { continue }
                $lnkPath = [System.IO.Path]::Combine($defaultShortcutDir, "PowerShell 2.0$($entry.Suffix).lnk")
                $target  = [System.IO.Path]::Combine($entry.Dir, $patchedExeName)
                $icon    = [System.IO.Path]::Combine($entry.Dir, 'powershell.exe')
                if ([System.IO.File]::Exists($target) -and -not [System.IO.File]::Exists($lnkPath)) {
                    $innerName = [System.IO.Path]::GetFileNameWithoutExtension($patchedExeName)
                    $launchArgs = "-NoLogo -NoProfile -ExecutionPolicy Bypass -NoExit -Command Write-Host '$innerName -Version 2' -ForegroundColor Yellow; $innerName -Version 2"
                    New-RunAsShortcut -Path $lnkPath -Target $target -Arguments $launchArgs -WorkDir '%USERPROFILE%' -Desc "PowerShell 2.0$($entry.Suffix)" -Icon "$icon,0"
                    Write-Log "Shortcut : $lnkPath"
                }
            }
        } else {
            Write-Log 'ps2DLC installation failed, skipping binary patch and shortcuts.' 'Error'
        }
    }
    if ($script:UnattendedErrors -gt 0) {
        Write-Log "--- UNATTENDED FINISHED WITH $($script:UnattendedErrors) ERROR(S) ---" 'Error'
        Invoke-NotifySound
        exit 1
    }
    Write-Log '--- UNATTENDED COMPLETE ---'
    Invoke-NotifySound
    exit 0
}

# ============================================================================
# GUI MODE
# ============================================================================

# DPI awareness must be set before creating any window
if (-not ('DPIAware' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class DPIAware
{
    public static readonly IntPtr UNAWARE              = (IntPtr) (-1);
    public static readonly IntPtr SYSTEM_AWARE         = (IntPtr) (-2);
    public static readonly IntPtr PER_MONITOR_AWARE    = (IntPtr) (-3);
    public static readonly IntPtr PER_MONITOR_AWARE_V2 = (IntPtr) (-4);
    public static readonly IntPtr UNAWARE_GDISCALED    = (IntPtr) (-5);
    [DllImport("user32.dll", EntryPoint = "SetProcessDpiAwarenessContext", SetLastError = true)]
    private static extern bool NativeSetProcessDpiAwarenessContext(IntPtr Value);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool SetProcessDPIAware();
    public static void SetDpiAwareness(IntPtr context)
    {
        if (!NativeSetProcessDpiAwarenessContext(context)) { SetProcessDPIAware(); }
    }
}
'@
}
try { [System.Windows.Forms.Application]::EnableVisualStyles() }        catch { }
try { [DPIAware]::SetDpiAwareness([DPIAware]::PER_MONITOR_AWARE_V2) }  catch { }

$iconBase64 = "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"
$iconBytes  = [System.Convert]::FromBase64String($iconBase64)
$iconStream = [System.IO.MemoryStream]::new($iconBytes)
$iconImage  = [System.Drawing.Image]::FromStream($iconStream)
$iconBitmap = [System.Drawing.Bitmap]::new($iconImage)
$iconHandle = $iconBitmap.GetHicon()
# Set custom AppUserModelID so the taskbar uses Form.Icon instead of powershell.exe icon
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class AppId
{
    [DllImport("shell32.dll", SetLastError = true)]
    public static extern int SetCurrentProcessExplicitAppUserModelID(
        [MarshalAs(UnmanagedType.LPWStr)] string AppID);
}
'@
[AppId]::SetCurrentProcessExplicitAppUserModelID('Freenitial.PS2Patcher') | Out-Null

# ── Loading splash form ──
$loadingForm = New-Object System.Windows.Forms.Form
$loadingForm.SuspendLayout()
$loadingForm.Text            = "PowerShell 2.0 Patcher"
$loadingForm.Font            = New-Object System.Drawing.Font("Microsoft Sans Serif", 8.25, [System.Drawing.FontStyle]::Regular)
$loadingForm.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
$loadingForm.AutoScaleMode   = [System.Windows.Forms.AutoScaleMode]::Dpi
$loadingForm.ClientSize      = New-Object System.Drawing.Size(300, 120)
$loadingForm.StartPosition   = "CenterScreen"
$loadingForm.FormBorderStyle = "FixedDialog"
$loadingForm.ControlBox      = $false
$loadingForm.Cursor          = [System.Windows.Forms.Cursors]::WaitCursor
$loadingIcon = New-Object System.Windows.Forms.PictureBox
$loadingIcon.Location = New-Object System.Drawing.Point(126, 10)
$loadingIcon.Size     = New-Object System.Drawing.Size(48, 48)
$loadingIcon.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Normal
$loadingIcon.Tag      = $iconImage
$loadingIcon.Add_Paint({
    param($s, $e)
    $g = $e.Graphics
    $g.InterpolationMode  = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.PixelOffsetMode    = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
    $g.DrawImage($s.Tag, 0, 0, $s.Width, $s.Height)
})
$loadingForm.Controls.Add($loadingIcon)
$loadingLabel = New-Object System.Windows.Forms.Label
$loadingLabel.Location  = New-Object System.Drawing.Point(10, 68)
$loadingLabel.Size      = New-Object System.Drawing.Size(280, 18)
$loadingLabel.Text      = "Loading interface..."
$loadingLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$loadingLabel.Anchor    = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$loadingForm.Controls.Add($loadingLabel)
$launch_progressBar = New-Object System.Windows.Forms.ProgressBar
$launch_progressBar.Location = New-Object System.Drawing.Point(10, 92)
$launch_progressBar.Size     = New-Object System.Drawing.Size(280, 20)
$launch_progressBar.Style    = "Continuous"
$launch_progressBar.Value    = 10
$launch_progressBar.Anchor   = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$loadingForm.Controls.Add($launch_progressBar)
$loadingForm.ResumeLayout($true)
$loadingForm.Show()
[System.Windows.Forms.Application]::DoEvents()

# ============================================================================
# THEME DETECTION & APPLICATION
# ============================================================================

function Test-SystemDarkMode {
    # Returns $true if Windows is set to dark app theme.
    try {
        $regPath = 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize'
        $value = (Get-ItemProperty -Path $regPath -Name 'AppsUseLightTheme' -ErrorAction Stop).AppsUseLightTheme
        return ($value -eq 0)
    } catch { return $false }
}

if (-not ('DwmTheme' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class DwmTheme
{
    // DWMWA_USE_IMMERSIVE_DARK_MODE = 20 (Windows 10 20H1+ / Windows 11)
    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
    public static void SetTitleBarDarkMode(IntPtr hwnd, bool dark)
    {
        int val = dark ? 1 : 0;
        DwmSetWindowAttribute(hwnd, 20, ref val, sizeof(int));
    }
}
'@
}

$script:IsDarkMode = Test-SystemDarkMode

# Color palettes
if ($script:IsDarkMode) {
    $script:ThemeColors = @{
        FormBack       = [System.Drawing.Color]::FromArgb(32, 32, 32)
        FormFore       = [System.Drawing.Color]::FromArgb(230, 230, 230)
        GroupBoxFore   = [System.Drawing.Color]::FromArgb(200, 200, 200)
        GroupBoxBorder = [System.Drawing.Color]::FromArgb(50,90,120)
        ButtonBack     = [System.Drawing.Color]::FromArgb(55, 55, 55)
        ButtonFore     = [System.Drawing.Color]::FromArgb(230, 230, 230)
        ButtonFlat     = [System.Drawing.Color]::FromArgb(80,80,80)
        TextBoxBack    = [System.Drawing.Color]::FromArgb(45, 45, 45)
        TextBoxFore    = [System.Drawing.Color]::FromArgb(220, 220, 220)
        LogBack        = [System.Drawing.Color]::FromArgb(20, 20, 20)
        LogFore        = [System.Drawing.Color]::FromArgb(204, 204, 204)
        CheckBoxFore   = [System.Drawing.Color]::FromArgb(220, 220, 220)
        BannerFore     = [System.Drawing.Color]::FromArgb(75,190,250)
    }
} else {
    $script:ThemeColors = @{
        FormBack       = [System.Drawing.SystemColors]::Control
        FormFore       = [System.Drawing.SystemColors]::ControlText
        GroupBoxFore   = [System.Drawing.SystemColors]::ControlText
        GroupBoxBorder = [System.Drawing.SystemColors]::ControlDark
        ButtonBack     = [System.Drawing.SystemColors]::Control
        ButtonFore     = [System.Drawing.SystemColors]::ControlText
        ButtonFlat     = [System.Drawing.SystemColors]::ControlDark
        TextBoxBack    = [System.Drawing.SystemColors]::Window
        TextBoxFore    = [System.Drawing.SystemColors]::WindowText
        LogBack        = [System.Drawing.Color]::FromArgb(200, 200, 200)
        LogFore        = [System.Drawing.Color]::FromArgb(20, 20, 20)
        CheckBoxFore   = [System.Drawing.SystemColors]::ControlText
        BannerFore     = [System.Drawing.Color]::DarkBlue
    }
}

function Set-ControlTheme {
    # Recursively applies theme colors to a control and all its children.
    param([System.Windows.Forms.Control]$Control)
    switch ($Control.GetType().Name) {
        'Form' {
            $Control.BackColor = $script:ThemeColors.FormBack
            $Control.ForeColor = $script:ThemeColors.FormFore
        }
        'GroupBox' {
            $Control.ForeColor = $script:ThemeColors.GroupBoxFore
            $Control.BackColor = $script:ThemeColors.FormBack
            if ($script:IsDarkMode) {
                $Control.Add_Paint({
                    param($s, $e)
                    $g   = $e.Graphics
                    $box = $s
                    $textSize = $g.MeasureString($box.Text, $box.Font)
                    $textLeft = 8
                    $halfText = [int]($textSize.Height / 2)
                    $bgBrush = [System.Drawing.SolidBrush]::new($box.BackColor)
                    $g.FillRectangle($bgBrush, 0, 0, $box.Width, $box.Height)
                    $bgBrush.Dispose()
                    $textBrush = [System.Drawing.SolidBrush]::new($box.ForeColor)
                    $g.DrawString($box.Text, $box.Font, $textBrush, $textLeft, 0)
                    $textBrush.Dispose()
                    $pen = [System.Drawing.Pen]::new($script:ThemeColors.GroupBoxBorder)
                    $rect = [System.Drawing.Rectangle]::new(0, $halfText, $box.Width - 1, $box.Height - $halfText - 1)
                    $g.DrawLine($pen, $rect.X, $rect.Y, $textLeft - 2, $rect.Y)
                    $g.DrawLine($pen, $textLeft + [int]$textSize.Width + 1, $rect.Y, $rect.Right, $rect.Y)
                    $g.DrawLine($pen, $rect.X, $rect.Y, $rect.X, $rect.Bottom)
                    $g.DrawLine($pen, $rect.X, $rect.Bottom, $rect.Right, $rect.Bottom)
                    $g.DrawLine($pen, $rect.Right, $rect.Y, $rect.Right, $rect.Bottom)
                    $pen.Dispose()
                })
            }
        }
        'Button' {
            $Control.BackColor = $script:ThemeColors.ButtonBack
            $Control.ForeColor = $script:ThemeColors.ButtonFore
            if ($Control.Tag -eq 'ActionBtn') {
                $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                $Control.FlatAppearance.BorderColor = $script:ThemeColors.BannerFore
                $Control.FlatAppearance.BorderSize  = 1
                if ($script:IsDarkMode) {
                    $Control.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
                    $Control.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
                } else {
                    $Control.FlatAppearance.MouseOverBackColor = [System.Drawing.SystemColors]::ControlLight
                    $Control.FlatAppearance.MouseDownBackColor = [System.Drawing.SystemColors]::ControlDark
                }
            } else {
                if ($script:IsDarkMode) {
                    $Control.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
                    $Control.FlatAppearance.BorderColor       = $script:ThemeColors.ButtonFlat
                    $Control.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
                    $Control.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
                }
            }
        }
        'TextBox' {
            $Control.BackColor = $script:ThemeColors.TextBoxBack
            $Control.ForeColor = $script:ThemeColors.TextBoxFore
            if ($script:IsDarkMode) {
                $Control.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
            }
        }
        'RichTextBox' {
            $Control.BackColor = $script:ThemeColors.LogBack
            $Control.ForeColor = $script:ThemeColors.LogFore
        }
        'CheckBox' {
            $Control.ForeColor = $script:ThemeColors.CheckBoxFore
            $Control.BackColor = $script:ThemeColors.FormBack
        }
        'Label' {
            $Control.BackColor = $script:ThemeColors.FormBack
        }
        'TableLayoutPanel' {
            $Control.BackColor = $script:ThemeColors.FormBack
        }
        'FlowLayoutPanel' {
            $Control.BackColor = $script:ThemeColors.FormBack
        }
    }
    foreach ($child in $Control.Controls) {
        Set-ControlTheme -Control $child
    }
}

$tipProvider = New-Object System.Windows.Forms.ToolTip
$tipProvider.AutoPopDelay = 15000
$tipProvider.InitialDelay = 400

$launch_progressBar.Value = 15

# --- Main Form ---
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.SuspendLayout()
$mainForm.AutoScaleDimensions = New-Object System.Drawing.SizeF(96, 96)
$mainForm.AutoScaleMode   = [System.Windows.Forms.AutoScaleMode]::Dpi
$mainForm.ShowInTaskbar    = $true
$mainForm.Text             = 'PowerShell 2.0 Patcher'
$mainForm.Size             = New-Object System.Drawing.Size(620, 560)
$mainForm.MinimumSize      = New-Object System.Drawing.Size(500, 400)
$mainForm.FormBorderStyle  = 'FixedSingle'
$mainForm.MaximizeBox      = $false
$mainForm.StartPosition    = [System.Windows.Forms.FormStartPosition]::CenterScreen
$mainForm.Font             = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point)
$mainForm.Icon = [System.Drawing.Icon]::FromHandle($iconHandle)
$mainForm.Add_Load({ [DwmTheme]::SetTitleBarDarkMode($mainForm.Handle, $script:IsDarkMode) })

$launch_progressBar.Value = 30

# ============================================================================
# MAIN TABLE LAYOUT (single column, N rows depending on mode)
# ============================================================================

$mainTable = New-Object System.Windows.Forms.TableLayoutPanel
$mainTable.Dock        = [System.Windows.Forms.DockStyle]::Fill
$mainTable.ColumnCount = 1
$mainTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100))) | Out-Null
$mainTable.Padding     = New-Object System.Windows.Forms.Padding(10)
$nextRow = 0

# ============================================================================
# ROW : MODE BANNER
# ============================================================================

$bannerTable = New-Object System.Windows.Forms.TableLayoutPanel
$bannerTable.Dock        = [System.Windows.Forms.DockStyle]::Fill
$bannerTable.AutoSize    = $true
$bannerTable.ColumnCount = 2
$bannerTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100))) | Out-Null
$bannerTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))     | Out-Null
$bannerTable.RowCount = 1
$bannerTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize'))) | Out-Null
$bannerTable.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)
$modeLabel = New-Object System.Windows.Forms.Label
$modeLabel.AutoSize  = $true
$modeLabel.Dock      = [System.Windows.Forms.DockStyle]::Fill
$modeLabel.ForeColor = $script:ThemeColors.BannerFore
$modeLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
if ($script:FeatureAvailable) {
    $modeLabel.Text = 'PS 2.0 is available as a Windows feature on this system.'
    $tipProvider.SetToolTip($modeLabel, "Your Windows build still has the PowerShell 2.0 optional feature.`nJust enable it alongside .NET 3.5 using the buttons below.")
} else {
    $modeLabel.Text = 'PS 2.0 feature not available on this OS. Binary patching is required.'
    $tipProvider.SetToolTip($modeLabel, "Microsoft removed the PS 2.0 feature from this build (KB5063878).`nThis tool patches powershell.exe to restore CLR2 engine activation.`nPlace ps2DLC.zip (Microsoft mitigation package) next to this script.")
}
$bannerTable.Controls.Add($modeLabel, 0, 0)
$btnShowLog = New-Object System.Windows.Forms.Button
$btnShowLog.Text     = 'Show Log'
$btnShowLog.AutoSize = $true
$tipProvider.SetToolTip($btnShowLog, "Opens the log folder and selects the most recent log file.")
$btnShowLog.Add_Click({
    $logFiles = [System.IO.Directory]::GetFiles($script:LogDir, '*.log')
    if ($logFiles.Count -eq 0) {
        Write-Log 'No log file found.' 'Warning'
        return
    }
    # Find the most recent log file by last write time
    $newestLog = $null
    $newestTime = [datetime]::MinValue
    foreach ($f in $logFiles) {
        $fi = [System.IO.FileInfo]::new($f)
        if ($fi.LastWriteTimeUtc -gt $newestTime) {
            $newestTime = $fi.LastWriteTimeUtc
            $newestLog = $fi.FullName
        }
    }
    Start-Process 'explorer.exe' -ArgumentList "/select,`"$newestLog`""
})
$bannerTable.Controls.Add($btnShowLog, 1, 0)
$mainTable.Controls.Add($bannerTable, 0, $nextRow++)

$launch_progressBar.Value = 40

# ============================================================================
# ROW : PREREQUISITES
# ============================================================================

$prereqGroup = New-Object System.Windows.Forms.GroupBox
$prereqGroup.Text         = 'Prerequisites'
$prereqGroup.Dock         = [System.Windows.Forms.DockStyle]::Fill
$prereqGroup.AutoSize     = $true
$prereqGroup.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$prereqTable = New-Object System.Windows.Forms.TableLayoutPanel
$prereqTable.Dock        = [System.Windows.Forms.DockStyle]::Fill
$prereqTable.AutoSize    = $true
$prereqTable.ColumnCount = 4
$prereqTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))    | Out-Null
$prereqTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100)))| Out-Null
$prereqTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))    | Out-Null
$prereqTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))    | Out-Null
$prereqTable.RowCount = 2
$prereqTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize'))) | Out-Null
$prereqTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize'))) | Out-Null
# --- .NET 3.5 row ---
$lblNetFx3Icon = New-Object System.Windows.Forms.Label
$lblNetFx3Icon.AutoSize  = $true
$lblNetFx3Icon.Dock      = [System.Windows.Forms.DockStyle]::Fill
$lblNetFx3Icon.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$prereqTable.Controls.Add($lblNetFx3Icon, 0, 0)
$lblNetFx3Text = New-Object System.Windows.Forms.Label
$lblNetFx3Text.Text      = '.NET Framework 3.5 (includes CLR 2.0 runtime)'
$lblNetFx3Text.AutoSize  = $true
$lblNetFx3Text.Dock      = [System.Windows.Forms.DockStyle]::Fill
$lblNetFx3Text.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$tipProvider.SetToolTip($lblNetFx3Text, ".NET 3.5 provides the CLR 2.0 runtime (v2.0.50727)`nthat PowerShell 2.0 requires to run.")
$prereqTable.Controls.Add($lblNetFx3Text, 1, 0)
$btnNetFx3 = New-Object System.Windows.Forms.Button
$btnNetFx3.Text     = 'Enable'
$btnNetFx3.AutoSize = $true
$btnNetFx3.Enabled  = $false
$prereqTable.Controls.Add($btnNetFx3, 3, 0)
# --- PS2 engine row ---
$lblPs2Icon = New-Object System.Windows.Forms.Label
$lblPs2Icon.AutoSize  = $true
$lblPs2Icon.Dock      = [System.Windows.Forms.DockStyle]::Fill
$lblPs2Icon.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$prereqTable.Controls.Add($lblPs2Icon, 0, 1)
$lblPs2Text = New-Object System.Windows.Forms.Label
$lblPs2Text.AutoSize  = $true
$lblPs2Text.Dock      = [System.Windows.Forms.DockStyle]::Fill
$lblPs2Text.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$prereqTable.Controls.Add($lblPs2Text, 1, 1)
$btnPs2Download = New-Object System.Windows.Forms.Button
$btnPs2Download.Text     = 'Download'
$btnPs2Download.AutoSize = $true
$btnPs2Download.Enabled  = $false
$tipProvider.SetToolTip($btnPs2Download, "Downloads ps2DLC.zip from Microsoft`nand saves it next to this script.")
$prereqTable.Controls.Add($btnPs2Download, 2, 1)
$btnPs2 = New-Object System.Windows.Forms.Button
$btnPs2.Text     = 'Install'
$btnPs2.AutoSize = $true
$btnPs2.Enabled  = $false
$prereqTable.Controls.Add($btnPs2, 3, 1)
$prereqGroup.Controls.Add($prereqTable)
$mainTable.Controls.Add($prereqGroup, 0, $nextRow++)

$launch_progressBar.Value = 50

# ============================================================================
# ROW : TARGET ARCHITECTURE (feature-removed only)
# ============================================================================

$chkX64 = $null; $chkX86 = $null
if (-not $script:FeatureAvailable) {
    $archGroup = New-Object System.Windows.Forms.GroupBox
    $archGroup.Text         = 'Target Architecture'
    $archGroup.Dock         = [System.Windows.Forms.DockStyle]::Fill
    $archGroup.AutoSize     = $true
    $archGroup.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $archFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $archFlow.Dock         = [System.Windows.Forms.DockStyle]::Fill
    $archFlow.AutoSize     = $true
    $archFlow.WrapContents = $false
    $chkX64 = New-Object System.Windows.Forms.CheckBox
    $chkX64.Text     = 'PowerShell x64 (System32)'
    $chkX64.Checked  = $true
    $chkX64.AutoSize = $true
    $chkX64.Margin   = New-Object System.Windows.Forms.Padding(3, 3, 20, 3)
    $archFlow.Controls.Add($chkX64)
    $tipProvider.SetToolTip($chkX64, "Patch the 64-bit powershell.exe in`n$pathSystem32Ps")
    if ($isOs64Bit -and [System.IO.Directory]::Exists($pathSysWOW64Ps)) {
        $chkX86 = New-Object System.Windows.Forms.CheckBox
        $chkX86.Text     = 'PowerShell x86 (SysWOW64)'
        $chkX86.Checked  = $true
        $chkX86.AutoSize = $true
        $archFlow.Controls.Add($chkX86)
        $tipProvider.SetToolTip($chkX86, "Patch the 32-bit powershell.exe in`n$pathSysWOW64Ps")
    }
    $archGroup.Controls.Add($archFlow)
    $mainTable.Controls.Add($archGroup, 0, $nextRow++)
}

# ============================================================================
# ROW : SHORTCUT CREATION
# ============================================================================

$shortcutGroup = New-Object System.Windows.Forms.GroupBox
$shortcutGroup.Text         = 'Create shortcut(s) to launch PS 2.0 as admin'
$shortcutGroup.Dock         = [System.Windows.Forms.DockStyle]::Fill
$shortcutGroup.AutoSize     = $true
$shortcutGroup.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$shortcutTable = New-Object System.Windows.Forms.TableLayoutPanel
$shortcutTable.Dock        = [System.Windows.Forms.DockStyle]::Fill
$shortcutTable.AutoSize    = $true
$shortcutTable.ColumnCount = 3
$shortcutTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100))) | Out-Null
$shortcutTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))     | Out-Null
$shortcutTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))     | Out-Null
$shortcutTable.RowCount = 1
$shortcutTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize'))) | Out-Null
$txtShortcutDir = New-Object System.Windows.Forms.TextBox
$txtShortcutDir.Text = $defaultShortcutDir
$txtShortcutDir.Dock = [System.Windows.Forms.DockStyle]::Fill
$tipProvider.SetToolTip($txtShortcutDir, "Folder where shortcut(s) will be created.")
$shortcutTable.Controls.Add($txtShortcutDir, 0, 0)
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text     = 'Browse...'
$btnBrowse.AutoSize = $true
$shortcutTable.Controls.Add($btnBrowse, 1, 0)
$btnShortcutCreate = New-Object System.Windows.Forms.Button
$btnShortcutCreate.Text     = 'Create'
$btnShortcutCreate.AutoSize = $true
$tipProvider.SetToolTip($btnShortcutCreate, "Creates .lnk shortcut(s) with Run as administrator enabled.`nOn 64-bit systems : one shortcut per selected architecture.")
$shortcutTable.Controls.Add($btnShortcutCreate, 2, 0)
$shortcutGroup.Controls.Add($shortcutTable)
$mainTable.Controls.Add($shortcutGroup, 0, $nextRow++)

$launch_progressBar.Value = 60

# ============================================================================
# ROW : LOG (fills all remaining vertical space)
# ============================================================================

$logRowIndex = $nextRow
$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Dock        = [System.Windows.Forms.DockStyle]::Fill
$logBox.ReadOnly    = $true
$logBox.BackColor   = [System.Drawing.Color]::FromArgb(30, 30, 30)
$logBox.ForeColor   = [System.Drawing.Color]::FromArgb(204, 204, 204)
$logBox.Font        = New-Object System.Drawing.Font('Consolas', 8.5)
$logBox.BorderStyle = 'None'
$mainTable.Controls.Add($logBox, 0, $nextRow++)
$script:LogRichTextBox = $logBox

# ============================================================================
# ROW : ACTION BUTTONS (left group + right group) — bottom of the form
# ============================================================================

$btnRowTable = New-Object System.Windows.Forms.TableLayoutPanel
$btnRowTable.Dock        = [System.Windows.Forms.DockStyle]::Fill
$btnRowTable.AutoSize    = $true
$btnRowTable.ColumnCount = 2
$btnRowTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('Percent', 100))) | Out-Null
$btnRowTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle('AutoSize')))     | Out-Null
$btnRowTable.RowCount = 1
$btnRowTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize'))) | Out-Null
$btnRowTable.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
# Left-aligned action buttons
$leftFlow = New-Object System.Windows.Forms.FlowLayoutPanel
$leftFlow.Dock         = [System.Windows.Forms.DockStyle]::Fill
$leftFlow.AutoSize     = $true
$leftFlow.WrapContents = $false
$leftFlow.Margin       = New-Object System.Windows.Forms.Padding(0)
$btnDuplicate = $null; $btnReplace = $null; $btnOpenPs2 = $null
if (-not $script:FeatureAvailable) {
    $btnDuplicate = New-Object System.Windows.Forms.Button
    $btnDuplicate.Text     = 'Duplicate Powershell'
    $btnDuplicate.AutoSize = $true
    $btnDuplicate.Tag      = 'ActionBtn'
    $btnDuplicate.Enabled  = $false
    $leftFlow.Controls.Add($btnDuplicate)
    $tipProvider.SetToolTip($btnDuplicate, "Creates a patched copy named powershell2.exe.`nOriginal powershell.exe is never modified.`nSurvives Windows Update.")
    $btnReplace = New-Object System.Windows.Forms.Button
    $btnReplace.Text     = 'Patch existing Powershell'
    $btnReplace.AutoSize = $true
    $btnReplace.Tag      = 'ActionBtn'
    $btnReplace.Enabled  = $false
    $leftFlow.Controls.Add($btnReplace)
    $tipProvider.SetToolTip($btnReplace, "Patches the original powershell.exe in-place.`nBackup (.bak) created, ACL saved and restored.`nReverted by any Cumulative Update.")
}
$btnRowTable.Controls.Add($leftFlow, 0, 0)
# Right-aligned utility buttons
$rightFlow = New-Object System.Windows.Forms.FlowLayoutPanel
$rightFlow.AutoSize     = $true
$rightFlow.WrapContents = $false
$rightFlow.Margin       = New-Object System.Windows.Forms.Padding(0)
$btnUninstall = New-Object System.Windows.Forms.Button
$btnUninstall.Text     = 'Uninstall Patch'
$btnUninstall.AutoSize = $true
$btnUninstall.Enabled  = $false
$rightFlow.Controls.Add($btnUninstall)
$tipProvider.SetToolTip($btnUninstall, "Removes patched files, restores backups, deletes shortcuts.")
$btnOpenPs2 = New-Object System.Windows.Forms.Button
$btnOpenPs2.Text     = 'Open PS 2.0'
$btnOpenPs2.AutoSize = $true
$btnOpenPs2.Enabled  = $false
$rightFlow.Controls.Add($btnOpenPs2)
$tipProvider.SetToolTip($btnOpenPs2, "Opens a PowerShell console and launches PS 2.0.`nPriority : x64 Replace > x64 Duplicate > x86 Replace > x86 Duplicate.")
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text     = 'Refresh'
$btnRefresh.AutoSize = $true
$rightFlow.Controls.Add($btnRefresh)
$btnRowTable.Controls.Add($rightFlow, 1, 0)
$mainTable.Controls.Add($btnRowTable, 0, $nextRow++)

$launch_progressBar.Value = 70

# ============================================================================
# FINALIZE MAIN TABLE ROW STYLES
# ============================================================================
# Log row gets Percent 100% (fills remaining space), all others are AutoSize.

$mainTable.RowCount = $nextRow
for ($i = 0; $i -lt $nextRow; $i++) {
    if ($i -eq $logRowIndex) {
        $mainTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('Percent', 100))) | Out-Null
    } else {
        $mainTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle('AutoSize'))) | Out-Null
    }
}

# Fill-docked control added FIRST to the form
$mainForm.Controls.Add($mainTable)

# Apply theme to all controls recursively
Set-ControlTheme -Control $mainForm

# ============================================================================
# STATUS REFRESH
# ============================================================================

function Update-ButtonStates {
    # Recalculates action button enabled states without re-querying prereqs.
    # Uses $script:AllPrereqsMet and $script:AnyInstalled cached by Update-AllStatus.
    if (-not $script:FeatureAvailable) {
        $stX64 = Get-ArchPatchState -PsDir $pathSystem32Ps
        $stX86 = if ($isOs64Bit) { Get-ArchPatchState -PsDir $pathSysWOW64Ps }
                 else { @{ HasDuplicate=$false; HasInPlace=$false; HasOrphanPatch=$false } }
        $anyDup     = ($stX64.HasDuplicate -or $stX86.HasDuplicate)
        $anyInPlace = ($stX64.HasInPlace -or $stX86.HasInPlace -or $stX64.HasOrphanPatch -or $stX86.HasOrphanPatch)
        $script:AnyInstalled = ($anyDup -or $anyInPlace)
        if ($null -ne $chkX64) { $chkX64.Enabled = $script:AllPrereqsMet }
        if ($null -ne $chkX86) { $chkX86.Enabled = $script:AllPrereqsMet }
        if ($null -ne $btnDuplicate) { $btnDuplicate.Enabled = ($script:AllPrereqsMet -and (-not $anyDup)) }
        if ($null -ne $btnReplace)   { $btnReplace.Enabled   = ($script:AllPrereqsMet -and (-not $anyInPlace)) }
        if ($null -ne $btnOpenPs2)   { $btnOpenPs2.Enabled   = ($script:AllPrereqsMet -and $script:AnyInstalled) }
        $anyArchChecked = (($null -ne $chkX64) -and $chkX64.Checked) -or (($null -ne $chkX86) -and $chkX86.Checked)
        $anyCheckedInstalled = $false
        if (($null -ne $chkX64) -and $chkX64.Checked) {
            if ($stX64.HasDuplicate -or $stX64.HasInPlace -or $stX64.HasOrphanPatch) { $anyCheckedInstalled = $true }
        }
        if (($null -ne $chkX86) -and $chkX86.Checked) {
            if ($stX86.HasDuplicate -or $stX86.HasInPlace -or $stX86.HasOrphanPatch) { $anyCheckedInstalled = $true }
        }
        $btnUninstall.Enabled = ($anyArchChecked -and $anyCheckedInstalled)
    } else {
        if ($null -ne $btnOpenPs2) { $btnOpenPs2.Enabled = $script:AllPrereqsMet }
        $btnUninstall.Enabled = $script:AnyInstalled
    }
    $btnShortcutCreate.Enabled = ($script:AllPrereqsMet -and $script:AnyInstalled)
    $btnBrowse.Enabled         = $btnShortcutCreate.Enabled
    $txtShortcutDir.Enabled    = $btnShortcutCreate.Enabled
    $btnRefresh.Enabled = $true
}

function Update-AllStatus {
    # Refreshes all prerequisite indicators then delegates button states to Update-ButtonStates.
    $allMet = $true
    $anyInstalled = $false
    Set-PrereqIcon $lblNetFx3Icon 'Busy'
    Set-PrereqIcon $lblPs2Icon 'Busy'
    $mainForm.Refresh()
    Reset-DismModule
    # --- .NET 3.5 ---
    $netFx3State = Get-WindowsOptionalFeature -Online -FeatureName $featureNameNetFx3 -ErrorAction SilentlyContinue
    $netFx3Ok = ($netFx3State -and $netFx3State.State -eq 'Enabled')
    if ($netFx3Ok) {
        Set-PrereqIcon $lblNetFx3Icon 'OK'; $btnNetFx3.Enabled = $false
    } else {
        Set-PrereqIcon $lblNetFx3Icon 'Fail'; $btnNetFx3.Enabled = $true; $allMet = $false
    }
    $lblNetFx3Icon.Refresh()
    # --- PS2 engine ---
    if ($script:FeatureAvailable) {
        $lblPs2Text.Text = 'PowerShell 2.0 Engine (Windows Feature)'
        $tipProvider.SetToolTip($lblPs2Text, "The PS 2.0 optional feature.`nEnable alongside .NET 3.5 to use powershell -Version 2.")
        $ps2State = Get-WindowsOptionalFeature -Online -FeatureName $featureNameRoot -ErrorAction SilentlyContinue
        $ps2Ok = ($ps2State -and $ps2State.State -eq 'Enabled')
        if ($ps2Ok) {
            Set-PrereqIcon $lblPs2Icon 'OK'; $btnPs2.Enabled = $false; $anyInstalled = $true
        } else {
            Set-PrereqIcon $lblPs2Icon 'Fail'; $btnPs2.Enabled = $netFx3Ok; $allMet = $false
        }
    } else {
        $zipPresent = [System.IO.File]::Exists($ps2DlcZipPath)
        $gacPresent = [System.IO.Directory]::Exists([System.IO.Path]::Combine($env:SystemRoot, 'assembly', 'GAC_MSIL', 'System.Management.Automation'))
        $regPresent = $null -ne (Get-ItemProperty -Path $registryKeyPath -ErrorAction SilentlyContinue)
        $dlcOk = ($gacPresent -and $regPresent)
        if ($dlcOk) {
            Set-PrereqIcon $lblPs2Icon 'OK'; $btnPs2.Enabled = $false; $btnPs2Download.Enabled = $false
            $lblPs2Text.Text = 'PS 2.0 Engine (ps2DLC installed)'
            $tipProvider.SetToolTip($lblPs2Text, "ps2DLC assemblies are in the legacy GAC and`nPowerShell\1 registry key is present.")
        } else {
            Set-PrereqIcon $lblPs2Icon 'Fail'; $btnPs2.Enabled = $zipPresent; $allMet = $false
            $btnPs2Download.Enabled = (-not $zipPresent)
            $lblPs2Text.Text = if ($zipPresent) { 'PS 2.0 Engine (ps2DLC ready)' } else { 'PS 2.0 Engine (ps2DLC.zip NOT FOUND!)' }
            $tipProvider.SetToolTip($lblPs2Text, $(if ($zipPresent) {
                "Click Install to deploy ps2DLC assemblies into the GAC."
            } else {
                "Place ps2DLC.zip next to this script :`n$scriptDirectory`nDownload from Microsoft KB5065506."
            }))
        }
    }
    $lblPs2Icon.Refresh()
    # Cache prereq results for Update-ButtonStates
    $script:AllPrereqsMet = $allMet
    $script:AnyInstalled  = $anyInstalled
    Update-ButtonStates
}

$launch_progressBar.Value = 80

# ============================================================================
# EVENT HANDLERS
# ============================================================================

if ($null -ne $chkX64) { $chkX64.Add_CheckedChanged({ Update-ButtonStates }) }
if ($null -ne $chkX86) { $chkX86.Add_CheckedChanged({ Update-ButtonStates }) }

$btnPs2Download.Add_Click({
    Disable-AllActionButtons
    $btnPs2Download.Enabled = $false
    $downloadUrl = 'https://download.microsoft.com/download/2b37839b-e146-465a-a78c-c9066609c553/ps2DLC.zip'
    $tempDownloadPath = [System.IO.Path]::Combine($scriptDirectory, 'ps2DLC.zip.downloading')
    Write-Log "Downloading ps2DLC.zip from Microsoft..."
    Write-Log "  URL : $downloadUrl" 'Debug'
    Write-Log "  Destination : $ps2DlcZipPath" 'Debug'
    $webClient = New-Object System.Net.WebClient
    $script:downloadDone  = $false
    $script:downloadOk    = $false
    $script:downloadError = $null
    $script:downloadProgressIdx = -1
    # Async progress handler (fires on ThreadPool, only update script vars)
    $webClient.Add_DownloadProgressChanged({
        param($s, $e)
        $script:downloadPercent = $e.ProgressPercentage
    })
    # Async completion handler
    $webClient.Add_DownloadFileCompleted({
        param($s, $e)
        if ($e.Cancelled) {
            $script:downloadError = 'Download cancelled.'
        } elseif ($e.Error) {
            $script:downloadError = $e.Error.Message
            if ($e.Error.InnerException) { $script:downloadError += " : $($e.Error.InnerException.Message)" }
        } else {
            $script:downloadOk = $true
        }
        $script:downloadDone = $true
    })
    $script:downloadPercent = 0
    # UI timer to poll progress and update the log
    $dlTimer = New-Object System.Windows.Forms.Timer
    $dlTimer.Interval = 200
    $dlTimer.Add_Tick({
        $pct = $script:downloadPercent
        $progressText = "  Downloading : $pct%"
        if ($script:downloadProgressIdx -ge 0) {
            $script:LogRichTextBox.Select($script:downloadProgressIdx, $script:LogRichTextBox.TextLength - $script:downloadProgressIdx)
            $script:LogRichTextBox.SelectedText = $progressText
        } else {
            $script:downloadProgressIdx = $script:LogRichTextBox.TextLength
            $script:LogRichTextBox.AppendText($progressText)
        }
        $script:LogRichTextBox.Select($script:LogRichTextBox.TextLength, 0)
        $script:LogRichTextBox.ScrollToCaret()
    })
    # Start async download to temp file
    try {
        $webClient.DownloadFileAsync([Uri]$downloadUrl, $tempDownloadPath)
    } catch {
        Write-Log "ERROR starting download : $($_.Exception.Message)" 'Error'
        $webClient.Dispose()
        Update-ButtonStates
        return
    }
    $dlTimer.Start()
    # Pump message loop while waiting
    while (-not $script:downloadDone) {
        [System.Windows.Forms.Application]::DoEvents()
        [System.Threading.Thread]::Sleep(15)
    }
    $dlTimer.Stop()
    $dlTimer.Dispose()
    $webClient.Dispose()
    # Finalize the progress line in the log
    if ($script:downloadProgressIdx -ge 0) {
        $script:LogRichTextBox.AppendText("`r`n")
        $script:downloadProgressIdx = -1
    }
    if ($script:downloadOk) {
        # Rename temp file to final name
        try {
            if ([System.IO.File]::Exists($ps2DlcZipPath)) {
                [System.IO.File]::Delete($ps2DlcZipPath)
            }
            [System.IO.File]::Move($tempDownloadPath, $ps2DlcZipPath)
            Write-Log "ps2DLC.zip downloaded to $ps2DlcZipPath"
        } catch {
            Write-Log "ERROR moving downloaded file : $($_.Exception.Message)" 'Error'
        }
    } else {
        Write-Log "Download failed : $($script:downloadError)" 'Error'
        # Cleanup partial temp file
        if ([System.IO.File]::Exists($tempDownloadPath)) {
            try { [System.IO.File]::Delete($tempDownloadPath) } catch { }
        }
    }
    Write-LogSeparator
    Invoke-NotifySound
    Update-AllStatus
})

$btnBrowse.Add_Click({
    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog
    $fbd.Description = 'Select folder for shortcut(s)'; $fbd.SelectedPath = $txtShortcutDir.Text
    if ($fbd.ShowDialog() -eq 'OK') { $txtShortcutDir.Text = $fbd.SelectedPath }
})

$btnShortcutCreate.Add_Click({
    Disable-AllActionButtons
    $shortcutDir = $txtShortcutDir.Text
    if (-not [System.IO.Directory]::Exists($shortcutDir)) {
        try { [System.IO.Directory]::CreateDirectory($shortcutDir) | Out-Null }
        catch { Write-Log "Cannot create directory : $shortcutDir" 'Error'; Update-ButtonStates; return }
    }
    $created = 0
    if ($script:FeatureAvailable) {
        $archEntries = @( @{ Dir=$pathSystem32Ps; Suffix=if($isOs64Bit){' (x64)'}else{''} } )
        if ($isOs64Bit -and [System.IO.Directory]::Exists($pathSysWOW64Ps)) {
            $archEntries += @{ Dir=$pathSysWOW64Ps; Suffix=' (x86)' }
        }
        foreach ($entry in $archEntries) {
            $lnkName = "PowerShell 2.0$($entry.Suffix).lnk"
            $lnkPath = [System.IO.Path]::Combine($shortcutDir, $lnkName)
            $target  = [System.IO.Path]::Combine($entry.Dir, 'powershell.exe')
            try {
                New-RunAsShortcut -Path $lnkPath -Target $target -Arguments '-Version 2 -NoExit' -WorkDir '%USERPROFILE%' -Desc "PS 2.0$($entry.Suffix)" -Icon "$target,0"
                Write-Log "Shortcut : $lnkName"
                $created++
            } catch { Write-Log "Shortcut error : $($_.Exception.Message)" 'Error' }
        }
    } else {
        $archEntries = @()
        $bothChecked = (($null -ne $chkX64) -and $chkX64.Checked) -and (($null -ne $chkX86) -and $chkX86.Checked)
        if (($null -ne $chkX64) -and $chkX64.Checked) {
            $archEntries += @{ Dir=$pathSystem32Ps; Label='x64'; Suffix=if($bothChecked){' (x64)'}else{''} }
        }
        if (($null -ne $chkX86) -and $chkX86.Checked) {
            $archEntries += @{ Dir=$pathSysWOW64Ps; Label='x86'; Suffix=if($bothChecked){' (x86)'}else{''} }
        }
        foreach ($entry in $archEntries) {
            $st = Get-ArchPatchState -PsDir $entry.Dir
            $targetExe = $null; $targetName = $null
            if ($st.HasDuplicate) {
                $targetExe = [System.IO.Path]::Combine($entry.Dir, $patchedExeName)
                $targetName = $patchedExeName
            } elseif ($st.HasInPlace -or $st.HasOrphanPatch) {
                $targetExe = [System.IO.Path]::Combine($entry.Dir, 'powershell.exe')
                $targetName = 'powershell.exe'
            } else {
                Write-Log "  No patch detected for $($entry.Label), skipping shortcut." 'Warning'
                continue
            }
            $lnkName = "PowerShell 2.0$($entry.Suffix).lnk"
            $lnkPath = [System.IO.Path]::Combine($shortcutDir, $lnkName)
            $icon = [System.IO.Path]::Combine($entry.Dir, 'powershell.exe')
            $innerName = [System.IO.Path]::GetFileNameWithoutExtension($targetName)
            $launchArgs = "-NoLogo -NoProfile -ExecutionPolicy Bypass -NoExit -Command Write-Host '$innerName -Version 2' -ForegroundColor Yellow; $innerName -Version 2"
            try {
                New-RunAsShortcut -Path $lnkPath -Target $targetExe -Arguments $launchArgs -WorkDir '%USERPROFILE%' -Desc "PS 2.0 $($entry.Label)" -Icon "$icon,0"
                Write-Log "Shortcut ($($entry.Label)) : $lnkName -> $targetName"
                $created++
            } catch { Write-Log "Shortcut error : $($_.Exception.Message)" 'Error' }
        }
    }
    if ($created -gt 0) {
        $btnShortcutCreate.Text = 'OK !'
        $script:shortcutFeedbackTimer = New-Object System.Windows.Forms.Timer
        $script:shortcutFeedbackTimer.Interval = 1000
        $script:shortcutFeedbackTimer.Add_Tick({
            $btnShortcutCreate.Text = 'Create'
            $script:shortcutFeedbackTimer.Stop()
            $script:shortcutFeedbackTimer.Dispose()
            $script:shortcutFeedbackTimer = $null
        })
        $script:shortcutFeedbackTimer.Start()
    }
    Update-ButtonStates
})

$btnNetFx3.Add_Click({
    Disable-AllActionButtons
    Set-PrereqIcon $lblNetFx3Icon 'Busy'; $lblNetFx3Icon.Refresh()
    Write-Log '=== .NET Framework 3.5 ==='
    Install-NetFx3Feature | Out-Null
    Write-LogSeparator
    Invoke-NotifySound
    Update-AllStatus
})

$btnPs2.Add_Click({
    Disable-AllActionButtons
    Set-PrereqIcon $lblPs2Icon 'Busy'; $lblPs2Icon.Refresh()
    if ($script:FeatureAvailable) {
        Write-Log '=== PS 2.0 Feature ==='
        Install-Ps2Feature | Out-Null
    } else {
        Write-Log '=== ps2DLC ==='
        Install-Ps2DlcPackage | Out-Null
    }
    Write-LogSeparator
    Invoke-NotifySound
    Update-AllStatus
})

$btnRefresh.Add_Click({ Update-AllStatus; Write-Log 'Status refreshed.' })

if ($null -ne $btnDuplicate) {
    $btnDuplicate.Add_Click({
        Disable-AllActionButtons
        Write-Log '=== DUPLICATE INSTALL ==='
        $archTargets = @()
        $bothChecked = (($null -ne $chkX64) -and $chkX64.Checked) -and (($null -ne $chkX86) -and $chkX86.Checked)
        if (($null -ne $chkX64) -and $chkX64.Checked) {
            $archTargets += @{ Dir=$pathSystem32Ps; Label='x64'; Suffix=if($bothChecked){' (x64)'}else{''} }
        }
        if (($null -ne $chkX86) -and $chkX86.Checked) {
            $archTargets += @{ Dir=$pathSysWOW64Ps; Label='x86'; Suffix=if($bothChecked){' (x86)'}else{''} }
        }
        if ($archTargets.Count -eq 0) { Write-Log 'No architecture selected.' 'Warning'; Update-ButtonStates; return }
        foreach ($t in $archTargets) {
            Write-Log "--- $($t.Label) ---"
            Install-BinaryPatch -PsDir $t.Dir -ArchLabel $t.Label -InPlace $false | Out-Null
        }
        Write-Log 'Done.'
        Write-LogSeparator
        Invoke-NotifySound
        Update-ButtonStates
    })
}

if ($null -ne $btnReplace) {
    $btnReplace.Add_Click({
        Disable-AllActionButtons
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "This will modify the original powershell.exe (backup created).`nAny Cumulative Update will revert the patch.`n`nContinue?",
            'Confirm In-Place Patch', 'YesNo', 'Warning')
        if ($confirm -ne 'Yes') { Update-ButtonStates; return }
        Write-Log '=== IN-PLACE INSTALL ==='
        $archTargets = @()
        $bothChecked = (($null -ne $chkX64) -and $chkX64.Checked) -and (($null -ne $chkX86) -and $chkX86.Checked)
        if (($null -ne $chkX64) -and $chkX64.Checked) {
            $archTargets += @{ Dir=$pathSystem32Ps; Label='x64'; Suffix=if($bothChecked){' (x64)'}else{''} }
        }
        if (($null -ne $chkX86) -and $chkX86.Checked) {
            $archTargets += @{ Dir=$pathSysWOW64Ps; Label='x86'; Suffix=if($bothChecked){' (x86)'}else{''} }
        }
        if ($archTargets.Count -eq 0) { Write-Log 'No architecture selected.' 'Warning'; Update-ButtonStates; return }
        foreach ($t in $archTargets) {
            Write-Log "--- $($t.Label) ---"
            Install-BinaryPatch -PsDir $t.Dir -ArchLabel $t.Label -InPlace $true | Out-Null
        }
        Write-Log 'Done.'
        Write-LogSeparator
        Invoke-NotifySound
        Update-ButtonStates
    })
}

if ($null -ne $btnOpenPs2) {
    $btnOpenPs2.Add_Click({
        $launchInfo = Get-Ps2LaunchInfo
        if ($null -eq $launchInfo) {
            Write-Log 'No suitable PowerShell 2.0 executable found.' 'Error'
            return
        }
        $exePath = $launchInfo.ExePath
        $innerExeName = $launchInfo.InnerExeName
        Write-Log "Launching PS 2.0 via $exePath" 'Debug'
        $innerCmd = "$innerExeName -Version 2"
        try {
            Start-Process $exePath -ArgumentList "-NoLogo -NoProfile -ExecutionPolicy Bypass -NoExit -Command `"Write-Host '$innerCmd' -ForegroundColor Yellow; $innerCmd`""
        } catch {
            Write-Log "ERROR launching PS 2.0 : $($_.Exception.Message)" 'Error'
        }
    })
}

$btnUninstall.Add_Click({
    Disable-AllActionButtons
    if ($script:FeatureAvailable) {
        Write-Log '=== UNINSTALL (Feature) ==='
        try {
            Disable-WindowsOptionalFeature -Online -FeatureName $featureNameRoot -NoRestart -ErrorAction Stop | Out-Null
            Write-Log 'PS 2.0 feature disabled.'
        } catch { Write-Log "ERROR : $($_.Exception.Message)" 'Error' }
    } else {
        Write-Log '=== UNINSTALL (Patch) ==='
        $archTargets = @()
        if (($null -ne $chkX64) -and $chkX64.Checked) {
            $archTargets += @{ Dir=$pathSystem32Ps; Label='x64' }
        }
        if (($null -ne $chkX86) -and $chkX86.Checked) {
            $archTargets += @{ Dir=$pathSysWOW64Ps; Label='x86' }
        }
        foreach ($entry in $archTargets) {
            $st = Get-ArchPatchState -PsDir $entry.Dir
            if ($st.HasDuplicate -or $st.HasInPlace -or $st.HasOrphanPatch) {
                Write-Log "--- $($entry.Label) ---"
                Uninstall-BinaryPatch -PsDir $entry.Dir -ArchLabel $entry.Label
            }
        }
    }
    Uninstall-Shortcuts -Dir $txtShortcutDir.Text
    Write-Log 'Done.'
    Write-LogSeparator
    Invoke-NotifySound
    Update-AllStatus
})

# ============================================================================
# LAUNCH
# ============================================================================

$launch_progressBar.Value = 90

Update-AllStatus
$mainForm.ResumeLayout($true)
$loadingForm.Close()
$loadingForm.Dispose()
[System.Windows.Forms.Application]::Run($mainForm)
