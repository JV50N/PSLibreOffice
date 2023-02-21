IF EXIST %WINDIR%\SysNative\WindowsPowerShell\v1.0 (SET PowerShellDir=%WINDIR%\SysNative\WindowsPowerShell\v1.0) ELSE (SET PowerShellDir=%WINDIR%\System32\WindowsPowerShell\v1.0)

"%PowerShellDir%\powershell.exe" -ExecutionPolicy ByPass -File "%~dp0\Install-LibOff.ps1"