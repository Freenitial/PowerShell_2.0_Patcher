# PowerShell 2.0 Patcher for Windows 10/11

Restores the PowerShell 2.0 engine on modern Windows builds where Microsoft has removed it, by patching the native `powershell.exe` stub.

On older builds where PS 2.0 is still available as a Windows Optional Feature, the tool simply enables it.

-----

## IF YOU ARE LOOKING TO INSTALL POWERSHELL 2.0 ON XP/VISTA : 
https://github.com/Freenitial/PowerShell-2.0_.NET-4_for_Windows_XP-2003-Vista-2008

-----

<img width="606" height="553" alt="image" src="https://github.com/user-attachments/assets/3925fced-e6c8-4de8-8f09-9c39b0d6a70d" />

## Quick Start

1. *(Optional)* Download [ps2DLC.zip](https://download.microsoft.com/download/2b37839b-e146-465a-a78c-c9066609c553/ps2DLC.zip) from [Microsoft Support (KB5065506)](https://support.microsoft.com/en-us/topic/powershell-2-0-removal-from-windows-fe6d1edc-2ed2-4c33-b297-afe82a64200a) (or from this repo)
2. *(Optional)* Place `ps2DLC.zip` next to the script.
3. Double-click `PowerShell_2.0_Patcher.bat` (auto-elevates to admin).
4. Click **Enable** next to ".NET Framework 3.5" if not already installed.
5. Click **Install** next to "PS 2.0 Engine" to deploy ps2DLC assemblies.
6. Click **Duplicate Powershell** or **Patch existing PowerShell** to create a patched `powershell2.exe`.
7. Click **Open PS 2.0** or create shortcuts from interface.

-----

### For fully automated installation:

Download `ps2DLC.zip` from [Microsoft](https://download.microsoft.com/download/2b37839b-e146-465a-a78c-c9066609c553/ps2DLC.zip) (or from this repo)

Place `ps2DLC.zip` next to the script, then run :
```
powershell -ExecutionPolicy Bypass -File Install-PowerShell2Patch.ps1 -Unattended
```

## How It Works

### Mode Detection

At startup, the tool queries DISM for the `MicrosoftWindowsPowerShellV2Root` optional feature:

- **Feature exists** (enabled or disabled) → the tool enables it via DISM. No binary patching needed.
- **Feature missing** (purged from CBS) → binary patching is required.

This detection method is more reliable than hardcoding a Windows build number, as it accounts for Insider rings, servicing branches, and SKU variations.

### The Deprecation Block

Starting with KB5063878 (August 2025, build 26100.4946), Microsoft removed the PS 2.0 optional feature from Windows 11 24H2 and later. The native `powershell.exe` stub now contains a hardcoded check:

```
if (requestedVersion <= 2) {
    displayWarning(MUI_RESOURCE_40);   // "PowerShell 2.0 is deprecated..."
    requestedVersion = 3;              // force PowerShell\3 = CLR4 = PS 5.1
}
```

The version number is then formatted into a registry path:

```
SOFTWARE\Microsoft\PowerShell\%d\PowerShellEngine
```

Two registry keys exist:

| Key | `RuntimeVersion` | Engine |
|-----|-------------------|--------|
| `PowerShell\1` | `v2.0.50727` | CLR2 → PS 2.0 |
| `PowerShell\3` | `v4.0.30319` | CLR4 → PS 5.1 |

`PowerShell\2` does **not exist** — version numbering jumped from 1 to 3 historically.

### The Patch

The tool remaps the version override from `3` to `1`, causing the stub to open `PowerShell\1\PowerShellEngine`, read `RuntimeVersion = v2.0.50727`, and activate CLR2 via `mscoree.dll CorBindToRuntimeEx`. The warning display call is replaced with NOPs.

#### x64 (PE32+) — Single Deprecation Block

```asm
; ORIGINAL (28 bytes, pattern-matched with 6 wildcard bytes)
lea     r8d, [rdx+0x28]           ; MUI resource string #40
mov     rax, [rax+8]              ; vtable slot
call    <display_warning>         ; relative call (wildcarded)
mov     dword ptr [r12], 3        ; override version → 3
mov     dword ptr [rbp+??], 3     ; override copy → 3 (offset wildcarded)

; PATCHED
lea     r8d, [rdx+0x28]           ; (unchanged)
mov     rax, [rax+8]              ; (unchanged)
nop (x5)                          ; suppress warning
mov     dword ptr [r12], 1        ; override version → 1
mov     dword ptr [rbp+??], 1     ; override copy → 1
```

Net effect: 3 bytes changed (`E8` → `90` at call site, two `03` → `01`).

#### x86 (PE32) — Two Deprecation Blocks

The 32-bit `powershell.exe` in `SysWOW64` uses a completely different instruction encoding. There are **two** independent deprecation blocks, each sharing a common 20-byte call sequence:

```asm
push    0x28                      ; MUI resource #40
push    0                         ; flags
push    eax                       ; resource handle
mov     ecx, [eax]                ; vtable
mov     esi, [ecx+4]              ; display function pointer
mov     ecx, esi                  ; thiscall self
call    [import_addr]             ; indirect call (4 bytes wildcarded)
call    esi                       ; actual warning display
```

Each block is patched by:
1. NOPing `call [import_addr]` (6 bytes) and `call esi` (2 bytes) = 8 NOP bytes total
2. Forward-scanning up to 50 bytes for `mov [reg], 3` instructions (`C7` opcode with ModRM decoding for `mod=00` and `mod=01`) and changing the immediate from `03` to `01`

The version overrides in the x86 binary use varying ModRM encodings across blocks (`C7 07` = `[edi]`, `C7 45 xx` = `[ebp+disp8]`, `C7 00` = `[eax]`), which is why a forward-scan with ModRM decoding is used instead of a fixed pattern.

### Patching Strategies

| Strategy | Action | Survives CU | Requires Ownership |
|----------|--------|-------------|-------------------|
| **Duplicate** | Creates `powershell2.exe` alongside the original | Yes | No |
| **Replace** | Patches the original `powershell.exe` in-place | No | Yes (TrustedInstaller) |

#### In-Place Patching and the Running-EXE Lock

`powershell.exe` cannot be overwritten while it is running (the current process is loaded from it). The tool uses an NTFS rename trick:

1. Save the full ACL (owner + DACL) via `File.GetAccessControl()`
2. Take ownership via `takeown.exe` + `icacls.exe` (handles `SeTakeOwnershipPrivilege` internally)
3. Fallback: apply ownership via .NET `SetAccessControl()` with SID `S-1-5-32-544`
4. **Rename** the locked original to `.bak` (NTFS allows renaming mapped executables)
5. Write patched bytes to a **new file** with the original name (filename is now free)
6. Restore the saved ACL (TrustedInstaller ownership) on the new file

This approach avoids the "file in use" error entirely. The `.bak` backup is used for restoration during uninstall.

### Pattern-Based Detection

The tool never uses hardcoded file offsets. All patching is done by scanning the binary for byte patterns with wildcard masks. This ensures compatibility across different Windows builds where the compiler may generate different:

- Relative call displacements (x64)
- Absolute import addresses (x86)
- Stack frame offsets (`rbp+??` / `ebp+??`)

If a future Windows build restructures the `powershell.exe` stub significantly enough that the patterns no longer match, the tool reports the failure and makes no changes.

PE type detection (`PE32` vs `PE32+`) is automatic via the magic field at the PE optional header.

## Prerequisites

### .NET Framework 3.5

Provides the CLR 2.0 runtime (`v2.0.50727`) that PowerShell 2.0 runs on. Installed via DISM with real-time progress display. The tool handles common DISM errors:

| Exit Code | Meaning |
|-----------|---------|
| `0x800F081F` | Source files not found — internet connection required |
| `0x800F0954` | Download blocked by WSUS policy |
| `0x80070490` | Component not found in store — image may be corrupted |
| `3010` | Success but restart required |

### ps2DLC Package (Feature-Removed Systems Only)

Microsoft's own mitigation package ([KB5065506](https://support.microsoft.com/en-us/topic/powershell-2-0-removal-from-windows-fe6d1edc-2ed2-4c33-b297-afe82a64200a)) that restores PS 2.0 assemblies. The tool can download it automatically via the **Download** button or it can be placed manually next to the script.

The package contains:

- **9 DLLs per architecture** (amd64 + x86) installed into the legacy GAC via `System.EnterpriseServices.Internal.Publish.GacInstall()`:
  - `System.Management.Automation.dll` (v1.0.0.0, CLR2)
  - `Microsoft.PowerShell.ConsoleHost.dll`
  - `Microsoft.PowerShell.Commands.Management.dll`
  - `Microsoft.PowerShell.Commands.Utility.dll`
  - `Microsoft.PowerShell.Commands.Diagnostics.dll`
  - `Microsoft.PowerShell.Security.dll`
  - `Microsoft.WSMan.Management.dll`
  - `Microsoft.WSMan.Runtime.dll`
  - `pspluginwkr.dll`

- **Localized resource DLLs** for 9 languages (de-DE, en-US, es-ES, fr-FR, it-IT, ja-JP, ko-KR, pt-BR, ru-RU, zh-CN, zh-TW). The tool installs the matching system locale with en-US as fallback.

- **2 registry files** recreating `HKLM\SOFTWARE\Microsoft\PowerShell\1\` with:
  - `Install = 1`
  - `RuntimeVersion = v2.0.50727`
  - `ConsoleHostAssemblyName = Microsoft.PowerShell.ConsoleHost, Version=1.0.0.0, ...`
  - `ShellIds\Microsoft.PowerShell\ExecutionPolicy = Bypass`

## GUI Features

### Polyglot BAT/PS1 Launcher

The script is a polyglot: the `.bat` extension makes it double-clickable, the batch header checks Windows version (displays an HTA dialog if < Windows 10), auto-elevates to admin, then executes the embedded PowerShell code.

### Async DISM Progress

.NET Framework 3.5 installation uses `dism.exe` with async output capture. A compiled C# class (`ProcessOutputCollector`) handles `OutputDataReceived` events on the ThreadPool thread.

A `ConcurrentQueue<string>` bridges to the UI thread where a WinForms Timer drains lines. Progress bars (`[===50.0%===]`) overwrite each other on a single line via `RichTextBox.Select()` + `SelectedText`.

### Async ps2DLC Download

The **Download** button fetches `ps2DLC.zip` from Microsoft's CDN using `WebClient.DownloadFileAsync` with progress percentage displayed inline.

### Logging

All operations are logged simultaneously to:
- Console (color-coded by severity)
- Log file (`%TEMP%\PowerShell 2.0 Patcher\PS2Patcher_YYYYMMDD.log`)
- GUI RichTextBox

## Verification

After patching, launch PS 2.0 via the **Open PS 2.0** button or shortcut and verify:

```powershell
$PSVersionTable
# PSVersion        2.0
# CLRVersion       2.0.50727.xxxx

[System.Environment]::Version
# Major  Minor  Build  Revision
# 2      0      50727  xxxx

# PS 3.0+ features should fail:
[ordered]@{ a = 1 }          # TypeNotFound
class Test { }               # ReservedKeywordNotAllowed
@(1,2,3).Where({$_ -gt 1})   # MethodNotFound
Get-CimInstance              # CommandNotFoundException
```

## Uninstall

Click **Uninstall Patch** to remove all artifacts:

- **Duplicate mode**: deletes `powershell2.exe`
- **In-place mode**: renames the patched exe out of the way (NTFS rename trick), restores the `.bak` backup, restores TrustedInstaller ACL
- **Both present**: handles the combination (duplicate removed, then backup restored)
- **Shortcuts**: removes all known shortcut names from the specified directory

On feature-available systems, uninstall disables the `MicrosoftWindowsPowerShellV2Root` optional feature via DISM.

## Limitations

- **Windows Update**: Cumulative Updates can replace `powershell.exe`, reverting in-place patches. The Duplicate strategy is immune to this.
- **Unsigned binary**: the patched copy has no valid Microsoft signature. SmartScreen or AppLocker may flag it.
- **Pattern compatibility**: if a future build restructures the native stub, the byte patterns may not match. The tool detects this and aborts without making changes.
- **CLR2 removal**: if Microsoft removes the CLR2 runtime itself (not just the PS 2.0 feature), the patch will no longer work.
- **x86 on 32-bit OS**: the x86 pattern has been verified on 64-bit WoW64 builds. 32-bit native Windows 10 builds have not been tested.

## Technical Reference

### Key Strings in the Binary

| Offset (x64) | String | Purpose |
|--------------|--------|---------|
| `.rdata` | `SOFTWARE\Microsoft\PowerShell\%1!ls!\PowerShellEngine` | Registry path format string |
| `.rdata` | `RuntimeVersion` | Registry value for CLR version selection |
| `.rdata` | `PowerShellVersion` | Registry value for engine version |
| `.rdata` | `version` | CLI argument name |
| `.rdata` | `NetFrameworkV4IsInstalled` | .NET 4 detection flag |
| `.text` (ASCII) | `Powershell_2.0_console_start` | ETW tracing marker |
| `.text` (ASCII) | `PSMajorVersion` | ETW field name |

### API Imports

The native stub imports from 6 DLLs:

| DLL | Relevant APIs |
|-----|---------------|
| `KERNEL32.dll` | `GetVersionExW`, `VerifyVersionInfoW` |
| `ADVAPI32.dll` | `RegQueryValueExW` |
| `mscoree.dll` | CLR hosting (`CorBindToRuntimeEx`) |
| `USER32.dll` | Message display |
| `OLE32.dll` / `OLEAUT32.dll` | COM initialization |

No CBS, DISM, or Feature Management APIs are imported.

### Localization

The deprecation warning is MUI resource string #40 (0x28), stored in satellite files:
- `...\en-US\powershell.exe.mui`
- `...\fr-FR\powershell.exe.mui`
- (other locales)

### Execution Flow After Patch

```
powershell2.exe -Version 2
  │
  ├─ Native stub parses arguments → version = 2
  ├─ Deprecation check: version ≤ 2 → enters block
  ├─ [PATCHED] Warning call: NOP'd
  ├─ [PATCHED] Version override: 2 → 1 (was 2 → 3)
  ├─ Registry path: SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine
  ├─ RegQueryValueEx: RuntimeVersion = "v2.0.50727"
  ├─ mscoree.dll CorBindToRuntimeEx("v2.0.50727") → CLR2 activated
  ├─ Loads System.Management.Automation 1.0 from legacy GAC
  └─ PowerShell 2.0 engine running
```

## References

- [KB5065506 — PowerShell 2.0 removal from Windows](https://support.microsoft.com/en-us/topic/powershell-2-0-removal-from-windows-fe6d1edc-2ed2-4c33-b297-afe82a64200a) (Microsoft Support, August 2025)
- [KB5063878 — First CU to remove PS2 feature](https://support.microsoft.com/en-us/topic/august-12-2025-kb5063878-os-build-26100-4946-e4b87262-75c8-4fef-9df7-4a18099ee294) (Build 26100.4946)
- [Windows PowerShell 2.0 Deprecation](https://devblogs.microsoft.com/powershell/windows-powershell-2-0-deprecation/) (PowerShell Team Blog, 2017)

-----

This tool and documentation are provided for educational and research purposes. Use at your own risk. PowerShell and Windows are trademarks of Microsoft Corporation.
