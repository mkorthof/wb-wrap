<# :
@echo off
SETLOCAL & SET "PS_BAT_ARGS=%~dp0 %*"
IF DEFINED PS_BAT_ARGS SET "PS_BAT_ARGS=%PS_BAT_ARGS:"="""%"
ENDLOCAL & powershell.exe -NoLogo -NoProfile -Command "&{ [ScriptBlock]::Create( ( Get-Content \"%~f0\" ) -join [char]10 ).Invoke( @( &{ $Args } %PS_BAT_ARGS% ) ) }"
GOTO :EOF
#>

<# -------------------------------------------------------------------------- #>
<# 20180718 MK: Wrapper for Windows 7 Backup                                  #>
<#            - Allows creating multiple system images using dated dirs       #>
<#            - Deletes oldest backup folder(s) until there is enough space   #>
<#            - Runs normal backup (using sdclt.exe)                          #>
<#            - Logging to "wb.log", located in same dir as script            #>
<# -------------------------------------------------------------------------- #>
<# For easy execution this PowerShell script is embedded in a Batch .cmd file #>
<# using a "polyglot wrapper". It can also be renamed to wb.ps1. More info:    #>
<#  https://blogs.msdn.microsoft.com/jaybaz_ms/2007/04/26/powershell-polyglot #>
<#  https://stackoverflow.com/questions/29645                                 #>
<# -------------------------------------------------------------------------- #>

<# -------------------------------------------------------------------------- #>
<# CONFIG                                                                     #>
<# -------------------------------------------------------------------------- #>

<# Either set $backupDrive to a static drive letter (e.g. "D") or a UUID:     #>
<# "\\?\Volume{c00d6b5a-f734-48e6-b321-029977e5169f". List UUIDS: 'mountvol'. #>
$backupDrive = "\\?\Volume{c00d6b5a-f734-48e6-b321-029977e5169f}\"

<# Set Volume Label of Backup Drive #>
$backupLabel = "My Backups"

<# Is the Backup Drive encrypted with BitLocker [0/1] #>
$blDrive = 1

<# Disk space options: % free space needed, bytes, max num folders to remove. #>
$minPercentFree = 12.5
$minBytesFree = 375809638400
$maxRemove = 3
<# Both minPercentFree and minBytesFree (actually "OR") are used to check     #>
<# against free disk space, as long as either one checks out no folders will  #>
<# be removed. You could look at your current backup to calculate these.      #>
<# ----------------------------------------------------------END-OF-CONFIG--- #>

$log = 1; $admin = 1; $debug = 0; $force = 0; $hibernate = 0; $test = 0
$removeDrive = 0

If ($Args -iMatch "[-/](\?|h$|help)") {
	Write-Host -ForegroundColor White "`r`nWindows Backup Wrapper`r`n"
	Write-Host "  Normally you should not need any of these, just configure options in wb.cmd"
	Write-Host "  and run it without arguments`r`n"
	Write-Host "USAGE (DEBUG) : wb.cmd -[d|f|a0|a1|b0|b1|h0|h1|l0|l1|r0|r1|t0|t1]" 
	Write-Host "DEBUG OPTIONS : d=debug, f=force, a=admin [0/1], b=bitlocker [0/1]"
	Write-Host "                h=hibernate [0/1], l=log [0/1], r=removedrive [0/1]"
	Write-Host "                t=test [0/1]"
	Write-Host "     DEFAULTS : log=$log debug=$debug force=$force admin=$admin bitlocker=$blDrive hibernate=$hibernate"
	Write-Host "                removeDrive=$removeDrive test=$test"
	Write-Host "     EXAMPLE  : wb.cmd -d -a1 -b0  (use debug, admin prompt on, bitlocker off)"
	Exit 0
}
If ($Args -iMatch "[-/]d")   { $debug = 1 }
If ($Args -iMatch "[-/]f")   { $force = 1 }
If ($Args -iMatch "[-/]a0")  { $admin = 0 }       ElseIf ($Args -iMatch "[-/]a1") { $admin = 1 }
If ($Args -iMatch "[-/]b0")  { $blDrive = 0 }     ElseIf ($Args -iMatch "[-/]b1") { $blDrive = 1 }
If ($Args -iMatch "[-/]h0")  { $hibernate = 0 }   ElseIf ($Args -iMatch "[-/]h1") { $hibernate = 1 }
If ($Args -iMatch "[-/]r0")  { $removeDrive = 0 } ElseIf ($Args -iMatch "[-/]r1") { $removeDrive = 1 }
If ($Args -iMatch "[-/]l0")  { $log = 0 }         ElseIf ($Args -iMatch "[-/]l1") { $log = 1 }
If ($Args -iMatch "[-/]t0")  { $test = 0 }        ElseIf ($Args -iMatch "[-/]t1") { $test = 1 }

$scriptDir = $Args[0]
$rm = ($Args[0]); $Args = ($Args) | Where { $_ -ne $rm }
$logFile = $scriptDir + "\wb.log"
If ($log -eq 1) {
	If ( $(Try { (Test-Path variable:local:logFile) -And (-Not [string]::IsNullOrWhiteSpace($logFile)) } Catch { $False }) ) {
		Write-Host "[INFO] -- Logging to: `"$logFile`"`r`n"
	} Else {
		$log = 0
		Write-Host "[WARN] -- Unable to open logfile, output to console only`r`n"
	}
}

<# -------------------------------------------------------------------------- #>
<# Test ("fake" settings for testing stuff) #>
<# -------------------------------------------------------------------------- #>
If ($test -eq 1)  { 
	#$minPercentFree = 3; $minBytesFree = 114000000000
	#$dirDate = Get-Date -Format yyMMdd
	#$backupDrive = "D"
	#$backupLabel = ""
	Write-Host "[DEBUG] minPercentFree=$minPercentFree minBytesFree=$minBytesFree"
	Write-Host "[DEBUG] test=$test backupDrive=$backupDrive backupLabel=$backupLabel"
}
<# -------------------------------------------------------------------------- #>

<# Function to log $msg to disk #>
Function Write-Log($msg) {
	If ($log -eq 1) { Add-Content $logFile -Value (((Get-Date).toString("yyyy-MM-dd HH:mm:ss")) + " $msg") }
}

# Function to display $msg and error
Function ErrorMessage($msg) {
	Write-Log "[ERROR] $msg $($_.Exception.Message) $($_.ErrorDetails)"
  	Write-Host
	<#Write-Host -ForegroundColor Red -NoNewLine ("[ERROR]"); Write-Host -NoNewLine (" -- $msg`r`n $(" "*9) $($_.Exception.Message)`r`n`r`n") #>
	Write-Host -ForegroundColor Red -NoNewLine ("[ERROR]"); Write-Host -NoNewLine (" -- $msg`r`n {0} {1} `r`n {0} {2} `r`n" -f $(" "*9), $_.Exception.Message, $_.ErrorDetails )
	<#Write-Error ("`r`n")#>
	If ($debug) { TIMEOUT /T 1 } Else { TIMEOUT /T 15 }
	Exit 1
}

<# Function to display or log disk space message, $target can be CONSOLE, INFO or DEBUG #>
Function DiskMessage($target) {
	If ($drive.TotalFreeSpace) {
		$backupDriveFree = ($drive.TotalFreeSpace/1024/1024/1024)
		$backupDriveFreePct = ($drive.TotalFreeSpace / $drive.TotalSize * 100)
	}
	$backupDriveMinFree = ($minBytesFree/1024/1024/1024)
	If ($target -eq "LOGFILE") {
		Write-Log "[INFO] Diskspace $backupVolName Currently free: $([int]($backupDriveFree))GB (need: $([int]($backupDriveMinFree))GB or $([int]($minPercentFree))%)"
	} Else {
		Switch ($target) {
			"CONSOLE" { $level = "INFO"; Break }
			"DEBUG" {   $level = "DEBUG"; Break }
		}
		Write-Host ("[{0}] -- Diskspace {1} Free: {2:#}GB or {3:g2}% (Need: {4:#}GB or {5:g2}%)" `
			-f $level, $backupVolName, $backupDriveFree, $backupDriveFreePct, $backupDriveMinFree, $minPercentFree)
	}
}

If ($admin) {
	If (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
		Write-Host; Write-Warning "You do not have Administrator rights to run this script!`nPlease re-run this script as an Administrator!"
		TIMEOUT /T 15
		Break
	}
}

Write-Log "[INFO] Start pid:$pid name:$($(Get-PSHostProcessInfo|where ProcessId -eq $pid).ProcessName)"

<# Get Drive Volume from Letter or UUID #>
If ( $(Try { (Test-Path variable:local:backupDrive) -And (-Not [string]::IsNullOrWhiteSpace($backupDrive)) } Catch { $False }) ) {
	If ($backupDrive -Match "^[a-zA-Z]$") {
		$backupVolume = (Get-Volume -DriveLetter $backupDrive)
	}	
	ElseIf ($backupDrive -Match "^\\\\\?\\Volume\{[a-fA-F0-9-]{36}\}\\") {
		$backupVolume = (Get-Volume -UniqueId "$backupDrive")
	}
	Else {
		ErrorMessage("No valid Drive Letter or UUID set"); Exit 1
	}
	If ((-Not (Test-Path variable:local:backupVolume)) -Or ( [string]::IsNullOrWhiteSpace($backupVolume))) {
		ErrorMessage("Could not find Backup Drive"); Exit 1
	}
	Else {
		$backupVolName = $($backupVolume).DriveLetter + ":"
		$Path = "${backupVolName}\WindowsImageBackup\$env:COMPUTERNAME"
		$drive = [System.IO.DriveInfo]::getdrives() | Where-Object {($_.Name -eq "$($backupVolName)\") -And ($_.VolumeLabel -eq "$($backupVolume.FileSystemLabel)")}
	}
} Else {
	ErrorMessage("backupDrive not set"); Exit 1
}

<# If Backup Drive uses Bitlocker make sure it's unlocked #>
If ($blDrive) {
	Try { $bitLocker = (Get-BitLockerVolume -ErrorAction Stop -WarningAction Stop -MountPoint "$backupVolName") } Catch { ErrorMessage("Could not get BitLocker Volume status") }
	If (-Not ($bitLocker.LockStatus -eq 'UnLocked')) {
		Write-Log "[INFO] Drive $backupVolName needs to be Unlocked by BitLocker first"
		Write-Host; Write-Host -NoNewLine "[INFO] -- "; Write-Host -ForeGroundColor White -NoNewLine "Drive $backupVolName needs to be Unlocked by BitLocker first"; Write-Host
		Unlock-BitLocker -ErrorAction Stop -WarningAction Stop -MountPoint "$backupVolName" -Confirm
		Try { $bitLocker = (Get-BitLockerVolume -ErrorAction Stop -WarningAction Stop -MountPoint "$backupVolName") } Catch { ErrorMessage("Could not Unlock BitLocker Volume") }
	}	
}

<# Check Disk Label #>
If (-Not ($drive.VolumeLabel -eq $backupLabel)) {
	ErrorMessage "Wrong disk! Drive `"$($drive.Name)`" has label `"$($drive.VolumeLabel)`" instead of `"$backupLabel`", exiting..."
	TIMEOUT /T 15
	Exit 1
}

<# Rename Backup folder from "WindowsImageBackup\COMPUTERNAME" to "COMPUTERNAME_yyMMdd" #>
If ((&Test-Path $Path) -Or ($force)) {
	If ($force -eq 0) { $dirDate = (Get-Item -Path $Path).LastWriteTime.ToString("yyMMdd") }
	If (-Not (&Test-Path ${Path}_$dirDate)) {
		If ($debug) { 
			Write-Host "[DEBUG] -- Rename-Item $Path ${Path}_$dirDate"
		} Else {
			If ($force -eq 0) {
				Try { Rename-Item -ErrorAction Stop -WarningAction Stop $Path ${Path}_$dirDate } Catch { ErrorMessage("Could not rename $Path") };
			}
		}
		$i = 0
		<# Make sure there is enough free disk space to backup #>
		If ($debug) {
			DiskMessage "DEBUG"
		}
		Else {
			DiskMessage "LOGFILE"
			DiskMessage "CONSOLE"
		}
		While (($($drive.TotalFreeSpace) -lt $minBytesFree) -Or (($drive.TotalFreeSpace / $drive.TotalSize * 100) -lt $minPercentFree)) {
			If ($i -ge $maxRemove) {
				$removeWait = 1
				Write-Log     "[WARN] Maximum of $maxRemove folders removed, exiting..."
				Write-Warning "[WARN] Maximum of $maxRemove folders removed, exiting..."
				TIMEOUT /T 15
				Exit 1
			} Else {
				$rmDir = Get-ChildItem -Path "$backupVolName\WindowsImageBackup" -Directory | Sort-Object CreationTime | Select-Object -First 1
				If ($debug) {
					DiskMessage "DEBUG"
					Write-Host  "[DEBUG] -- Remove-Item: $($rmDir.FullName)"
				} Else {
					Try { Remove-Item -ErrorAction Stop -WarningAction Stop -Recurse $($rmDir.FullName) } Catch { ErrorMessage("Could not remove $rmDir") }
					DiskMessage "LOGFILE"
					Write-Log   "[INFO] Removed oldest folder: $($rmDir.FullName)"

					DiskMessage "CONSOLE"
					Write-Host  "[INFO] -- Removed oldest folder: $($rmDir.FullName)"
				}
			}
			$i++
		}
		<# If we removed dirs wait 5 sec so user can read disk space msgs #>
		If ($removeWait) { TIMEOUT /T 5 }
		If ($debug) {
			Write-Host "DEBUG: Start-Process $env:SystemRoot\system32\sdclt.exe (using calc.exe as test)"
			Start-Process -FilePath "$env:SystemRoot\system32\calc.exe" -ArgumentList "/UIMODE /SHOW" -Wait -NoNewWindow -PassThru
			Sleep 5
			$i = 1; While ( @(Get-Process | Where {$_.Name -eq "Calculator"}).Count -gt 0 ) {
				Write-Host "DEBUG: Sleep 1 (i: $i)"; Sleep 1; $i++
			}
			Write-Host ("[INFO] -- Total backup duration was: {0} seconds (or {1:n0} minutes)" -f $i, ($i/60))
		} Else {
			<#
				Start Windows Backup :
				- Show progress window by starting sdclt.exe again with /UIMODE
				- Sleep until sdclt.exe exits and ask for key press so user can read total duration msg
				- Display error msg with ExitCode if first sdclt.exe cmd fails
			#>
			$p = Start-Process -FilePath "$env:SystemRoot\system32\sdclt.exe" -ArgumentList "/KICKOFFJOB" -Wait -NoNewWindow -PassThru
			If ($p.ExitCode -eq 0) {
				If ($hibernate -eq 0) {
					Start-Process -FilePath "$env:SystemRoot\system32\sdclt.exe" -ArgumentList "/UIMODE /SHOW" -NoNewWindow
				}
			} Else {
				ErrorMessage "Could not start Backup ($p.ExitCode)"
				TIMEOUT /T 15
				Exit 1
			}	
			Sleep 5
			$i = 1; While ( @(Get-Process | Where {$_.Path -eq "$env:SystemRoot\system32\sdclt.exe" }).Count -gt 0 ) {
				Sleep 1; $i++
			}
			Write-Log "[INFO] Total backup duration was: $i seconds (or $([int]($i/60)) minutes)"
			Write-Host ("[INFO] -- Total backup duration was: {0} seconds (or {1:n0} minutes)" -f $i, ($i/60))
			<# TODO: #>
			If ($hibernate -eq 0) {
				Write-Host; Write-Host -ForeGroundColor Yellow "Press any key to close window..."; Write-Host; [void][System.Console]::ReadKey($true)
			}
		}
		<# TODO: 
		If ($debug) { 
			If ($removeDrive) {
				If ($blDrive) { Write-Host "[DEBUG] -- ( first lock bitlocker? )" }
				Write-Host "[DEBUG] -- remove disk"
			}
		}
 		#>
		If ($hibernate) {
			Write-Log  "[INFO] Hibernating in 1 minute"
			Write-Host "[INFO] -- Hibernating in 1 minute..."
			TIMEOUT /T 60
			If ($debug) { 
				Write-Host "DEBUG: Invoke-Expression $env:SystemRoot\system32\shutdown.exe /h"
			} Else {
				Invoke-Expression "$env:SystemRoot\system32\shutdown.exe /h"
			}
		}
	} Else {
		ErrorMessage "${Path}_$dirDate) already exists"
		TIMEOUT /T 15
		Exit 1
	}
} Else {
	ErrorMessage "$Path not found"
	TIMEOUT /T 15
	Exit 1
}