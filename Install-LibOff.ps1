<# Script Info for Install-LibOff.ps1.ps1

  .VERSION: 1.0
  
  .INITIAL DATE: 02/21/2023

  .AUTHOR: Jason Goncalves
	
  .DESCRIPTION: Powershell script to locate previous versions of Libre Office on workstation and install new Libre Office Package.
  
#>

# Initilisations
# Set Error Actions to Silently Continue...
$ErrorActionPreference = "SilentlyContinue"

# Declarations
# Script Version
$ScriptVer = "1.0"

#$x64-System = ""
#$x86-System = ""

# Log File Info
$LogPath = "$env:SystemDrive\Temp"
$LogName = "Install_LibreOffice.log"
$TimeStamp = (Get-Date).toString("MM/dd/yyyy HH:mm:ss")

if (Test-Path $LogPath){
	Write-Host "["$TimeStamp"] " " Temp directory exists... " -ForegroundColor Yellow
} else {
	New-Item -Path $LogPath -ItemType Directory
	Write-Host "["$TimeStamp"] " " Creating Temp directory... " -ForegroundColor Green
}

$LogFile = "$LogPath\$LogName"

# If using PowerShell 3 or Greater
if ($PSVersionTable.PSVersion.Major -gt 3) {
	$ScriptPath = $PSScriptRoot
# If using PowerShell 2 or lower
} else {
	$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
}

$x64_Installer = "$ScriptPath\7.5.0\LibreOffice_7.5.0_Win_x86-64.msi"
$x64_HelpPackInstaller = "$ScriptPath\7.5.0\LibreOffice_7.5.0_Win_x86-64_helppack_en-US.msi"

$x86_Installer = "$ScriptPath\7.5.0\LibreOffice_7.5.0_Win_x86.msi"
$x86_HelpPackInstaller = "$ScriptPath\7.5.0\LibreOffice_7.5.0_Win_x86_helppack_en-US.msi"

# Functions
Function Start-Install {
	begin{
		Start-Transcript -Path $LogFile
		Write-Host "["$TimeStamp"] " " Running Libre Install Script version: $ScriptVer"
	}
	process{
		try{
			# See if previous versions of Libre Office are installed, if so, remove...
			Get-CimInstance -ClassName Win32_Product -Filter "Name LIKE 'OpenOffice%%' OR Name LIKE 'LibreOffice%%'" | Invoke-CimMethod -MethodName Uninstall
			
			# Check System Architecture and then install appropriate packages...
			if ((Get-WmiObject Win32_OperatingSystem | Select osarchitecture).osarchitecture -eq "64-bit"){
				# Install Libre Office x64
				Write-Host "["$TimeStamp"] "  " Installing Libre Office 7.5.0 for 64 Bit systems..."
				Start-Process "$SysPath\msiexec.exe" -ArgumentList "i/ `"$x64_Installer`" /quiet /passive /norestart" -Wait
				# Install Offline Help Pack x64
				Write-Host "["$TimeStamp"] " " Installing Libre Office Offline Help Pack for 64 Bit systems..."
				Start-Process "$SysPath\msiexec.exe" -ArgumentList "i/ `"$x64_HelpPackInstaller`" /quiet /passive /norestart" -Wait
			} elseif ((Get-WmiObject Win32_OperatingSystem | Select osarchitecture).osarchitecture -eq "32-bit") {
				# Install Libre Office x86
				Write-Host "["$TimeStamp"] " " Installing Libre Office 7.5.0 for 32 Bit systems..."
				Start-Process "$SysPath\msiexec.exe" -ArgumentList "i/ `"$x86_Installer`" /quiet /passive /norestart" -Wait
				# Install Offline Help Pack x86
				Write-Host "["$TimeStamp"] " " Installing Libre Office Offline Help Pack for 32 Bit Systems..."
				Start-Process "$SysPath\msiexec.exe" -ArgumentList "i/ `"$x86_HelpPackInstaller`" /quiet /passive /norestart" -Wait
			} else {
				Write-Host "["$TimeStamp"] " " This is not a typical windows architecture. Please Contact your System Admin..." - ForegroundColor Red
			}
			# Run Vuln Scan after Libre Office Install w/o UI
			Write-Host "["$TimeStamp"] " " Running Ivanti Security Scan w/o UI..."
			Start-Process "C:\Program Files (x86)\LANDesk\LDClient\vulscan.exe" -ArgumentList "/showui=false"
			
			# Run Inventory Scan w/o UI
			Write-Host "["$TimeStamp"] " " Running Ivanti Inventory Scan w/o UI..."
			Start-Process "C:\Program Files (x86)\LANDesk\LDClient\LDISCN32.exe" -ArgumentList "/NOUI"
		}
		catch{
			Write-Host "["$TimeStamp"] "  " Catastrophic error has occured! Please contact your System Admin..."`n -ForegroundColor Red
			Write-Host "["$TimeStamp"] "  " Message: [$($_.Exception.Message)"] -ForegroundColor Red -BackgroundColor DarkBlue
		}
	}
	end{
		Write-Host "["$TimeStamp"] "  " Install Libre Office Script completed successfully" -ForegroundColor Green
		Stop-Transcript
	}
}

# Execute
Start-Install
