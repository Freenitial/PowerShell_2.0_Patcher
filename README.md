# PowerShell 2.0 Patcher for Windows 10/11

Restores the PowerShell 2.0 engine on modern Windows builds where Microsoft has removed it, by patching the native `powershell.exe` stub.

2 ways are possible :
- By duplicating powershell : Creates powershell2.exe. To use PS 2.0, call `powershell2 -version 2`
- By patching powershell : Replace existing powershell.exe. To use PS 2.0, call `powershell -version 2`

On older builds where PS 2.0 is still available as a Windows Optional Feature, the tool simply enables it.

-----

## IF YOU ARE LOOKING TO INSTALL POWERSHELL 2.0 for XP OR VISTA : https://github.com/Freenitial/PowerShell-2.0_.NET-4_for_Windows_XP-2003-Vista-2008
 

-----

<img width="606" height="553" alt="image" src="https://github.com/user-attachments/assets/3925fced-e6c8-4de8-8f09-9c39b0d6a70d" />

## Quick Start

1. *(Optional)* Download [ps2DLC.zip](https://download.microsoft.com/download/2b37839b-e146-465a-a78c-c9066609c553/ps2DLC.zip) from [Microsoft Support (KB5065506)](https://support.microsoft.com/en-us/topic/powershell-2-0-removal-from-windows-fe6d1edc-2ed2-4c33-b297-afe82a64200a) (or from this repo)
2. *(Optional)* Place `ps2DLC.zip` next to the script.
3. Double-click `PowerShell_2.0_Patcher.bat` (auto-elevates to admin).
4. Click **Enable** next to ".NET Framework 3.5" if not already installed.
5. Click **Install** next to "PS 2.0 Engine" to deploy ps2DLC assemblies.
6. Click **Duplicate Powershell** or **Patch existing PowerShell**.
7. Click **Open PS 2.0** or create shortcuts from interface.

## Verification
After patching, launch PS 2.0 via the **Open PS 2.0** button or shortcut and verify `$PSVersionTable`

-----

### For fully automated installation (only duplicate, not patching existing powershell) :

Download `ps2DLC.zip` from [Microsoft](https://download.microsoft.com/download/2b37839b-e146-465a-a78c-c9066609c553/ps2DLC.zip) (or from this repo)

Place `ps2DLC.zip` next to the script, then run :
```
powershell -ExecutionPolicy Bypass -File PowerShell_2.0_Patcher.bat -Unattended
```

-----

## How It Works : 
See documentation at https://github.com/Freenitial/PowerShell_2.0_Patcher/blob/main/Documentation.md

## Limitations

- **Windows Update**: Cumulative Updates can replace `powershell.exe`, reverting in-place patches. The Duplicate strategy is immune to this.
- **Unsigned binary**: the patched copy has no valid Microsoft signature. SmartScreen or AppLocker may flag it.
- **Pattern compatibility**: if a future build restructures the native stub, the byte patterns may not match. The tool detects this and aborts without making changes.
- **CLR2 removal**: if Microsoft removes the CLR2 runtime itself (not just the PS 2.0 feature), the patch will no longer work.
- **x86 on 32-bit OS**: the x86 pattern has been verified on 64-bit WoW64 builds. 32-bit native Windows 10 builds have not been tested.

## References

- [KB5065506 — PowerShell 2.0 removal from Windows](https://support.microsoft.com/en-us/topic/powershell-2-0-removal-from-windows-fe6d1edc-2ed2-4c33-b297-afe82a64200a) (Microsoft Support, August 2025)
- [KB5063878 — First CU to remove PS2 feature](https://support.microsoft.com/en-us/topic/august-12-2025-kb5063878-os-build-26100-4946-e4b87262-75c8-4fef-9df7-4a18099ee294) (Build 26100.4946)
- [Windows PowerShell 2.0 Deprecation](https://devblogs.microsoft.com/powershell/windows-powershell-2-0-deprecation/) (PowerShell Team Blog, 2017)

-----

This tool and documentation are provided for educational and research purposes. Use at your own risk. PowerShell and Windows are trademarks of Microsoft Corporation.
