##******************************************************************
## Revision date: 2024.01.25
##
##		2023.06.22: Proof of concept / Initial release
##		2023.08.10:	Cleanup ;-)
##		2024.01.14: Allow system disk size reduction following KB5034439 (Windows Server)
##					and KB5034441 (Windows Workstation) January 2024 Cumulative updates.
##		2024.01.15: Add warning/explanation at startup.
##		2024.01.24: Allow creation of the recovery partition from the recovery environment
##					installed on the system partition. Also, allow replacement of a damaged
##					recovery partition.
##		2024.01.25:	Send output from bcdedit /create to Null: subsequent commands will fail
##					with the real issue if "$WindowsRELoaderOptions" cannot be updated.
##
## Copyright (c) 2023-2024 PC-Évolution enr.
## This code is licensed under the GNU General Public License (GPL).
##
## THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
## ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
## IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
## PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
##
##******************************************************************

param (
	[Parameter(mandatory=$true, HelpMessage='Set the drive letter to use for the Recovery partition')]
	[char]$UseLetter,
	[parameter()]
	[switch]$Log,
	[parameter()]
	[switch]$Details,
	[parameter()]
	[string]$SourcesDir,
	[parameter()]
	[int]$ExtendedSize
)
$Usage = @"

Usage: $($MyInvocation.MyCommand.Name) -UseLetter R -ExtendedSize <size> -Log -Details -SourcesDir <Path>

where:
`t-UseLetter`tis the drive letter that will be assigned to the recovery partition.
`t`t`tThis letter is only used during excution of the script. There is no default value.

`t-ExtendedSize`tis the size of the new recovery partition. If <size> is less than 1MB, <size>
`t`t`tis multiplied by 1MB, e.g. 600 implies 600MB.

`t-Log`t`tCreate a transcript log on the user's desktop.

`t-Details`tDisplay detailed information throughout execution of this script.

`t-SourcesDir`tis the directory of the Windows Installation Media containing Install.win.
`t`t`tThis is only used if this script finds no Recovery Environment on this system.

"@

Function RestartVDS {
	## "Disk Management Services" can interfere in many ways if something is trying to display/manipulate partition data.
	## Stop all Microsoft Consoles in which "Disk Management Services" are used
	Try {
		# Stop Microsoft Management Consoles that may use Disk Management Services
		$Process = (Get-Process -ProcessName mmc -ErrorAction Stop) `
					| Where-Object { $_.Modules.Description -match "Disk Management Snap-in" } `
					| Select-Object -Property ID | Stop-Process
		# ... and insist if the user did not cooperate.
		$Process = (Get-Process -ProcessName mmc -ErrorAction Stop) `
					| Where-Object { $_.Modules.Description -match "Disk Management Snap-in" } `
					| Select-Object -Property ID | Stop-Process -Force
	 }
	Catch {
		# Do nothing!
	 }
	Finally {
		## Restart the "Virtual Dik Service"
		Restart-Service -Name VDS
	 }
}

$ResizeSystemPartition = {
	Param ($FreeSizeRequired)

	<#
		Shrink/Expand the system partition, not knowing what will happen, and reserve
		enough space to create a Recovery partition.
	#>

	If ($ExtendedSize -lt 1MB) {$ExtendedSize = $ExtendedSize * 1MB }
	$NewRESize =[math]::Max( $ExtendedSize, $FreeSizeRequired )
	
	## For Windows 11, the recommended recovery partition size is 990MB
	## See https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/hard-drives-and-partitions?view=windows-11
	If ($NewRESize -lt 990MB) {
		If ($(Read-Host "Enter 'Yes' to use the recommended minimum partition size of 990MB for Windows RE, anything else to continue").tolower().StartsWith('yes')) `
		  {	$NewRESize = 990MB }
	  }

	Write-Host ""
	Write-Host "Computing new System partition size. Please be patient ..."
	$SupportedSize = (Get-PartitionSupportedSize -DiskId $SystemPartition.DiskID -Offset $SystemPartition.Offset).SizeMax
	$NewSize = $SupportedSize - $NewRESize
	
	If ($SystemPartition.Size -ne $NewSize) {
		Try {
			Resize-Partition -DiskId $SystemPartition.DiskID -Offset $SystemPartition.Offset -Size $NewSize -ErrorAction Stop
		  }
		Catch {
			Write-Warning "Error during System partition resize. You should keep the backup of the Recovery partition."
			Write-Warning "An attempt to restore a Recovery partition follows..."
		 }
	  }

}

$CreateRecoveryPartition = {
	Param ($UseLetter)
<#
	Create a new Recovery partition in the poper format, fix BCD entries and restore contents from backup.
	
	Note: there is no notion of disk geometry here. All free space on the system disk is alloated to this
	new partition: there is no way of computing an exact size based on the available information.
	
#>

	Write-Host ""
	Write-Host "Creating new Windows RE partition ..."

	RestartVDS
	If ($Disks.PartitionStyle -eq "GPT") {
		## "Microsoft Recovery" partitions are created "hidden" and even DISKPART cannot clear this attribute.
		## Create a basic data partition and change its type
		$NewTarget = New-Partition -DiskPath $Disks.Path -DriveLetter $UseLetter -GptType "{ebd0a0a2-b9e5-4433-87c0-68b6b72699c7}" -UseMaximumSize
		## Note: "-IsHidden" is superbly ignored on Windows 11/Server 2022 ;-)
		$NewTarget | Set-Partition -GptType "{de94bba4-06d1-4d40-a16a-bfd50179d6ac}" -IsHidden $True -NoDefaultDriveLetter $True
	<#
		This sets the "Recovery" partition type and the GPT_BASIC_DATA_ATTRIBUTE_NO_DRIVE_LETTER (0x8000000000000000) attribute.
			Diskpart will display this partition as:
			Type    : de94bba4-06d1-4d40-a16a-bfd50179d6ac
			Hidden  : Yes
			Required: No
			Attrib  : 0X8000000000000000
		Windows (at least Disk Management console) will not display this as a Recovery partition unless the GPT_ATTRIBUTE_PLATFORM_REQUIRED (0x0000000000000001) is also set. There is no PowerShell equivalent and we must resort to Diskpart to do so:
	#>
	## Must resort to an Here-String to run DiskPart
@"
select disk $($Disks.Number)
select partition $($NewTarget.PartitionNumber)
gpt attributes=0x8000000000000001
detail partition
exit
"@ | Diskpart
	}
	else {
		## PowerShell has no MBRtype for "Microsoft Recovery" partitions
		## Create a basic data partition and change its type
		$NewTarget = New-Partition -DiskPath $Disks.Path -DriveLetter $UseLetter -MBRType IFS -UseMaximumSize
		$NewTarget | Set-Partition -MBRType 0x27
	  }

	Write-Host ""
	Write-Host "Restoring the Windows RE partition ..."

	$NewVolume = $NewTarget | Format-Volume -FileSystem NTFS -NewFileSystemLabel "Recovery"
		# | Set-Volume -DriveLetter $UseLetter is implicit when the partition is created.

	## Create a skeleton directory structure
	New-Item -Path "$UseLetter`:\" -Name "Recovery" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null
	New-Item -Path "$UseLetter`:\Recovery" -Name "WindowsRE" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null
	New-Item -Path "$UseLetter`:\Recovery" -Name "Logs" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null

	## Display the Disk Management Consoles
	diskmgmt.msc

	Return $NewTarget
}

$LocateWindowsREImage = {
	Param ($RELocation)
<#
	REagentC stores a copy of the Windows RE image file in $([Environment]::SystemDirectory)\Recovery\
	
	If no such file exists, restore a copy from the Windows installation media (install.wim).
	
	Note that Winre.wim may have been updated : direct the user to "KB5034957: Updating the WinRE partition on deployed devices"
	
#>
	If ((Get-ChildItem -Path "$([Environment]::SystemDirectory)\Recovery\Winre.wim" -Force -ErrorAction SilentlyContinue).Exists) {
		Write-Warning "The Windows RE image can be enabled from the system directory ..."
		## Do NOT enable Windows RE at this time since the state of the Recovery partition is unknown
		## and "REagentC /Enable" might create the Recovery environment on the system drive.
		$RELocation = "$([Environment]::SystemDirectory)\Recovery"
	  }
	else {
			## Attempt to restore the Windows RE image from a source media.
			If (Test-Path -Path "$SourcesDir\Install.wim" -IsValid) {
				If (Test-Path -Path "$SourcesDir\Install.wim") {
					Write-Warning "Restoring Winre.wim from the Windows install source. This is a long process, be patient ..."
					$Decoy = $(New-Item -Type Directory -Path $env:TEMP -Name ([System.Guid]::NewGuid())).FullName
					dism /Mount-Wim /ReadOnly /WimFile:"$SourcesDir\install.wim" /Index:1 /MountDir:"$Decoy"
					## PowerShell has some limitations with system files!
					Try {
						Remove-Item -Path "$([Environment]::SystemDirectory)\Recovery\Winre.wim" -force -Confirm:$False `
							-ErrorAction Stop
					  }
					Catch {
						# Do Nothing!
					  }
					Finally {
						Copy-Item  -LiteralPath "$Decoy\Windows\System32\Recovery\Winre.wim" `
							-Destination "$([Environment]::SystemDirectory)\Recovery\Winre.wim" -force -Confirm:$False `
							-ErrorAction SilentlyContinue
					  }
					dism /Unmount-Wim /Discard /MountDir:"$Decoy"
					Remove-Item -LiteralPath "$Decoy" -Force -Recurse
					# Sanity check: did we get the image file?
					If ((Get-ChildItem -Path "$([Environment]::SystemDirectory)\Recovery\Winre.wim" -Force -ErrorAction SilentlyContinue).Exists) {
						Write-Warning "The Windows RE image was successfully restored from the installation media."
						Write-Warning "Please see KB5034957: Updating the WinRE partition on deployed devices."
						$RELocation = "$([Environment]::SystemDirectory)\Recovery"
					  }
					else {
						Write-Warning "Cannot locate a Windows RE image file. Continuing witout a known Windows RE image ..."
					  }
				  }
				else {
						# Source installation image not found
						Write-Warning "Cannot locate Install.wim. Continuing witout a known Windows RE image ..."
				  }
			  }
			else {
					# No valid path to a Windows source installation image
					Write-Warning "$SourcesDir is an invalid path. Continuing witout a known Windows RE image ..."
			  }
	  }

	Return $RELocation
}

$PopulateRecoveryPartition = {
	Param ($UseLetter)
<#
	We need to ensure a measure of health before anything else.

	REagentC /Disable may actually removes the contents of the Recovery partition.
	Here are the details:
	
		PS C:\Users\Administrator> dir R:\Recovery\ -Force -Recurse

			Directory: R:\Recovery

		Mode                 LastWriteTime         Length Name
		----                 -------------         ------ ----
		d--hs-          6/5/2023   1:58 PM                Logs
		d--hs-          6/5/2023   8:20 AM                WindowsRE

			Directory: R:\Recovery\Logs

		Mode                 LastWriteTime         Length Name
		----                 -------------         ------ ----
		-a----          6/5/2023  12:20 PM           1354 Reload.xml

			Directory: R:\Recovery\WindowsRE

		Mode                 LastWriteTime         Length Name
		----                 -------------         ------ ----
		---hs-          5/8/2021   4:14 AM        3170304 boot.sdi
		---hs-          6/5/2023   8:20 AM           1109 ReAgent.xml
		---hs-          8/6/2021   8:26 PM      440718104 Winre.wim

		PS C:\Users\Administrator> REagentC /Disable
		REAGENTC.EXE: Operation Successful.

		PS C:\Users\Administrator> dir R:\Recovery\ -Force -Recurse

			Directory: R:\Recovery

		Mode                 LastWriteTime         Length Name
		----                 -------------         ------ ----
		d--hs-          6/5/2023   1:58 PM                Logs

			Directory: R:\Recovery\Logs

		Mode                 LastWriteTime         Length Name
		----                 -------------         ------ ----
		-a----          6/5/2023  12:20 PM           1354 Reload.xml

	Winre.wim is tucked away in [Environment]::SystemDirectory)\Recovery when the Recovery environment
	is disabled:

		PS C:\Users\Administrator> dir "$([Environment]::SystemDirectory)\Recovery" -Force -Recurse

			Directory: C:\Windows\system32\Recovery

		Mode                 LastWriteTime         Length Name
		----                 -------------         ------ ----
		-a----         4/25/2023   2:05 PM            875 $PBR_Diskpart.txt
		-a----         4/25/2023   2:05 PM            468 $PBR_ResetConfig.xml
		-a----         6/17/2023  10:44 AM           1079 ReAgent.xml
		-a-hs-          8/6/2021   8:26 PM      440718104 Winre.wim

#>
		## Create / Repair the contents of the Recovery partition
		$GlobalVol = $Target.AccessPaths.Count - 1
		if ( [string]::IsNullOrEmpty($Target.DriveLetter) -or ($Target.DriveLetter -eq "`0") )
		  { ## Assign the drive letter
			$R = $Target | Set-Partition -NewDriveLetter $UseLetter
			Write-Host "Assigned drive letter $UseLetter to", $Target.AccessPaths[$GlobalVol]
		  }
		elseif ($Target.DriveLetter -ne $UseLetter)
		  { ## Override the user parameter with the assigned driver letter
			[char] $UseLetter = $($Target.DriveLetter)
			Write-Warning "Parameter override: using $UseLetter for the Recovery partition."
		  }
		New-Item -Path "$UseLetter`:\" -Name "Recovery" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null
		New-Item -Path "$UseLetter`:\Recovery" -Name "WindowsRE" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null
		New-Item -Path "$UseLetter`:\Recovery" -Name "Logs" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null
		
		## Install a copy of the Recovery image in this location
		Copy-Item -Path "$([Environment]::SystemDirectory)\Recovery\Winre.wim" `
					-Destination "$UseLetter`:\Recovery\WindowsRE\Winre.wim" -force -Confirm:$False `
					-ErrorAction SilentlyContinue
		
		## Rebuild the structures required to reboot into the Recovery environment
		REagentC /SetREImage /Path "$UseLetter`:\Recovery\WindowsRE"
		
		## This does NOT enable the Recovery environment.
		
		Return $UseLetter
}

$BackupRecoveryPartition = {
	Param ($Description)
	
<#

	Backup a recovery partition : we don't know the kind of customization or the up-to-date status
	of its contents. This is preferable to any attempt to rebuild this data.
	
	Note: Backup is NOT always relative to ...\Recovery

#>
	If ($CaptureDir -ne $Null) {
		Write-Host ""
		Write-Host "Testing $CaptureDir ($Description). Please be patient ..."
		
		# $CaptureDir can be specified using
		#	- a local path (e.g. C:\Windows\system32\Recovery)
		#	- a global path (e.g. \\?\GLOBALROOT\device\harddisk0\partition4\Recovery)
		#	- a unique partition ID (e.g. \\?\Volume{ce912212-c124-4492-8609-adb7a8e7a6ac}\)
		# dism will process all three but most Windows apps won't.

		# Get a hash of the WinRE image file 
		If ($CaptureDir.StartsWith('\\?\Volume{', 'CurrentCultureIgnoreCase')) {
			#BUG	One added frustration: Get-Volume returns $Null in this Context
			#BUG	For additionnal comments, see https://github.com/PowerShell/PowerShell/issues/16688
			#BUG		$ThisVolume = Get-Volume -UniqueID $CaptureDir	# Note: do not use (any) quotes, the \ character is not escaped in this string
			#BUG		if ( [string]::IsNullOrEmpty($ThisVolume.DriveLetter) -or ($ThisVolume.DriveLetter -eq " 0") )
			#BUG		  { $ThisVolume | Set-Volume -DriveLetter $UseLetter
			#BUG			Write-Host "Assigned drive letter $UseLetter to", $CaptureDir
			#BUG		  }

			# Get the drive letter from the partition instead of the volume
			$TargetGuid = $CaptureDir -replace "^.+(\{.*\}).*", '$1'
			$TargetPartition = Get-Partition | Where { $_.Guid -eq $TargetGuid }
			If ( [string]::IsNullOrEmpty($TargetPartition.DriveLetter) -or ($TargetPartition.DriveLetter -eq "`0") )
			  { $TargetPartition | Set-Partition -NewDriveLetter $UseLetter
				# Poor man's -Passthru (Yes! there are simpler ways, but I want the system's view of what is going on).
				$TargetPartition = Get-Partition | Where { $_.Guid -eq $TargetGuid }
				Write-Host "Assigned drive letter $UseLetter to", $CaptureDir
			  }
			$RELocation = $TargetPartition.DriveLetter + ":\Recovery"
		  }
		elseif ($CaptureDir.StartsWith('\\?\GLOBALROOT\device\', 'CurrentCultureIgnoreCase')) {
				$DiskNumber = $CaptureDir -replace "^\\\\\?\\GLOBALROOT\\device\\harddisk(\d+)\\partition\d+\\.*$", '$1'
				$PartitionNumber = $CaptureDir -replace "^\\\\\?\\GLOBALROOT\\device\\harddisk\d+\\partition(\d+)\\.*$", '$1'
				$TargetDir = $CaptureDir -replace "^\\\\\?\\GLOBALROOT\\device\\harddisk\d+\\partition\d+(\\.*)$", '$1'

				$TargetPartition = Get-Partition | Where { ($_.DiskNumber -eq $DiskNumber) -and ($_.PartitionNumber -eq $PartitionNumber) }
				If ( [string]::IsNullOrEmpty($TargetPartition.DriveLetter) -or ($TargetPartition.DriveLetter -eq "`0") )
				  { $TargetPartition | Set-Partition -NewDriveLetter $UseLetter
					# Poor man's -Passthru (Yes! there are simpler ways, but I want the system's view of what is going on).
					$TargetPartition = Get-Partition | Where { $_.Guid -eq $TargetGuid }
					Write-Host "Assigned drive letter $UseLetter to", $CaptureDir
				  }
				$RELocation = $TargetPartition.DriveLetter + $TargetDir
		  }
		else {
			$RELocation = $CaptureDir
		  }

		# Get-FileHash can access the image fle but does not error on missing files!
		# The directory structure is not consistent between the enabled and disabled states :
		Try {
			# a) Try the default location ...
			$OriginalWimSignature = (Get-FileHash -Algorithm SHA256 -LiteralPath "$RELocation\WindowsRE\Winre.wim" -ErrorAction Stop).Hash
		  }
		Catch {
			# b) Try the backup location
			$OriginalWimSignature = (Get-FileHash -Algorithm SHA256 -LiteralPath "$RELocation\Winre.wim" -ErrorAction SilentlyContinue).Hash
		  }

		If ($OriginalWimSignature -ne $Null) {

			Write-Host ""
			Write-Host "Creating a backup of the existing $Description. Please be patient ..."

			$Capture = dism /Capture-Image /ImageFile:$($Env:Temp+"\Recovery.wim") /CaptureDir:"$CaptureDir" /Name:Recovery /Verify /Quiet
			$CaptureContents = dism /List-Image /ImageFile:$($Env:Temp+"\Recovery.wim") /Index:1
			If ($Verbose) {
				Write-Host $($Capture | Out-String)
				Write-Host $($CaptureContents | Out-String)
				<#
					Here is a typical output: dism does NOT create a disk image:
					
						Deployment Image Servicing and Management tool
						Version: 10.0.19041.3636

						\
						\WindowsRE\
						\WindowsRE\boot.sdi
						\WindowsRE\ReAgent.xml
						\WindowsRE\Winre.wim
						The operation completed successfully.
				#>
			  }
			Write-Host "Windows RE SHA256 signature:", $OriginalWimSignature
			Write-Host ""
		  }
	  }
	Return $OriginalWimSignature # which may be unchanged
}

## Are we logging this ?
[Boolean]$Logging = $Log.IsPresent

If ($Logging) { Start-Transcript -Path "$($env:USERPROFILE)\Desktop\RecoveryPartitionMaintenance.txt" -Append }

[char]$UseLetter = "$UseLetter".ToUpper() # Traditionnal ;-)

## Determine how much will be displayed
[Boolean]$Verbose = $Details.IsPresent

## Create a separator for verbose output
$Separator = "-" * $Host.UI.RawUI.WindowSize.Width

## Warn user and display the current Windows Build
Write-Host ""
Write-Warning "-> Drive letter $UseLetter is used to manipulate the Recovey partition. This may conflict"
Write-Warning "-> with your current drive assignments."
Write-Host ""
Write-Warning "This script attempts to relocate the Recovery partition contiguous to the end of the"
Write-Warning "System partition. This may extend the System partition by allocating all free space"
Write-Warning "available following a disk size increase. This may shrink the system partition to"
Write-Warning "increase the size of the recovery partition (see MS KB5034439/KB5034441)."
Write-Warning ""
Write-Warning "The script will also repair a disabled recovery partition."
Write-Warning "Optionally, the script may create a recovery partition using the Windows installation"
Write-Warning "media."
Write-Warning ""
Write-Warning "The Recovery Environment can be brought up-to-date using Microsoft's KB5034957."
Write-Warning ""
Write-Warning "In any case: USE AT YOUR OWN RISK!"
Write-Warning ""
If (-not $(Read-Host "Enter 'Yes' to continue, anything else to exit").tolower().StartsWith('yes')) { Exit }
Write-Host $Usage
Write-Host ""

#
Write-Host ""
Write-Host $([System.Environment]::OSVersion.VersionString)
Write-Host ""

<#

	Avoid BitLocker encrypted disk objects.

#>

$SystemDrive =  (Get-WmiObject Win32_OperatingSystem).SystemDrive
$BitLocker = Get-WmiObject -Namespace "Root\cimv2\Security\MicrosoftVolumeEncryption" -Class "Win32_EncryptableVolume" -Filter "DriveLetter = '$SystemDrive'"
If (-not $BitLocker) {
	Write-Host "No BitLocker protection on drive $SystemDrive."
  }
elseif ( $BitLocker.GetProtectionStatus().protectionStatus -ne "0" ) {
	Write-Warning "BitLocker is enabled on this system disk. Aborting ..."
	If ($Logging) { Stop-Transcript }
	Exit 911
 }
else { Write-Host "BitLocker is off on this system." } 
Write-Host ""

<#

	Locate the System partition and get the corresponding disk object.

#>

$SystemPath =  $SystemDrive + "\"
$SystemPartition = $Null
ForEach ($Partition in $(Get-Partition)) {
	If ($Partition.AccessPaths.Count -gt 0) {
		If ($Partition.AccessPaths.Contains($SystemPath)) {
			$SystemPartition = $Partition
		  }
	  }
  }
If ($SystemPartition -eq $Null) {
	Write-Warning "Could not find the System partition!!!"
	If ($Logging) { Stop-Transcript }
	Exit 911
  }

$Disks = Get-Disk -Number $SystemPartition.DiskNumber

## Assume no WindowsRE partition will be selected
$Target = $Null

<#

	Ask the user to confirm the apparent recovery partition, if any.

	See the discussion in scriptblock PopulateRecoveryPartition regarding the behavior
	of "REagentC". We need a valid image file to relocate the Recovery partition.
	This script may use the -SourcesDir parameter to locate such a file.
	
	For now, locate and backup the Windows RE image file.
#>

$RELocation = $($(REagentC /Info | Select-String -Pattern 'GLOBALROOT') -replace '^.*\s', '')
If ($RELocation -eq $Null) {
	Write-Warning "The Recovery environment is currently disabled."
	$CaptureDir = &$LocateWindowsREImage ($RELocation)
	# Presume there is a valid WindowsRE image file available
	$RELocation = $CaptureDir
  }
else { # The Recovery environment is enabled and accessible through GLOBALROOT
	$CaptureDir = $RELocation -replace "(^*)\\WindowsRE", $1
}
$OriginalWimSignature = &$BackupRecoveryPartition("Recovey Environment")
Write-Host ""

If ($Verbose -and $RELocation -ne $Null) {
	Write-Host $Separator -ForeGroundColor Green
	Write-Host "Windows RE image attributes:"
	Write-Host ""
	Dism /Get-ImageInfo /ImageFile:"$RELocation\winre.wim" /index:1
	Write-Host $Separator -ForeGroundColor Green
  }

ForEach ($Disk in $Disks) {
	$Partitions = Get-Partition -DiskNumber $SystemPartition.DiskNumber
	## Caution: Get-PartitionSupportedSize has the side effect of starting the defragmentation service ("Optimixe Drive")
	## and returns an array of maximum sizes that is not always in the same order as the partition table!
	$MaxSize = $Disk.Size - $Disk.AllocatedSize
	If ($MaxSize -gt 1GB) {
		## Announce possible gains
		Write-Host "Approximately", $([int] [Math]::floor($($MaxSize / 1GB))), "GB can be allocated on this drive" -ForegroundColor Green
		Write-Host "presume all free space is contiguous at the end of this disk."  -ForegroundColor Green
		Write-Host
		If ($Verbose) {
			## Provide a short description of this disk
			$Disk | fl PartitionStyle, FriendlyName, GUID
			## Dump the partition table: don't rely on the order of the display
			$Partitions | ft Type, IsSystem, AccessPaths, GptType, MBRType, Size
		}
	}
	If ($Partitions.Count -gt 1) {
		ForEach ($Partition in $Partitions) {
			If ($Partition.AccessPaths.Count -gt 0) {
				If ( (($Disk.PartitionStyle -eq "MBR" -and $Partition.MBRType -eq 0x27) `
						-or ($Disk.PartitionStyle -eq "GPT" -and $Partition.GptType -eq "{de94bba4-06d1-4d40-a16a-bfd50179d6ac}") ) `
					-and $Partition.Offset -gt $SystemPartition.Offset ) {
					$GlobalVol = $Partition.AccessPaths.Count - 1
					If ($Verbose) { Write-Host "Testing:", $($Partition.AccessPaths[$GlobalVol]) }
					Try {
							## Note: Test-Path will not return a value for hidden/system objects.
							## In the default name space (drive letter), a suffix "\..\" is required to find the directory, which does not seem coherent
							## We have direct access to the directory using the global name space...
							## ... and this may point to a healthy Recovery partition which has a number of subdirectories
							$WinRE = Get-ChildItem -Attributes Directory,Directory+Hidden -Path $($Partition.AccessPaths[$GlobalVol]+"Recovery\") `
										-Directory -Depth 0 -Exclude "System Volume Information" -Force -ErrorAction Stop
							If ($WinRe.Count -ge 1) {
								Write-Host $($Partition.AccessPaths[$GlobalVol]), "is an apparent Recovery partition" -ForegroundColor Green
								If ($(Read-Host "Enter 'Yes' to select this partition as the target recovery partition, anything else to continue").tolower().StartsWith('yes')) {
									## Note: the above manipulations do not effect the local variables
									$Target = $Partition
								 }
							  }
					  }
					Catch {
						If ($Verbose) { Write-Host "-> Ignored" }
					  }
					 }
			  }
		  }
	  }
  }

<#
	At this point, we may have a recovery partition, we may have a Windows RE image file, and/or we may have some free space.

	This says nothing about the BCD entries and support structures required to boot the Recovery
	environment.
	
	Discussion continues below ...
#>
If ($RELocation -eq "$([Environment]::SystemDirectory)\Recovery") {
	If ($Target -eq $Null) {

		# Presume a minimum partition size (see MS KB5034439/KB5034441)
		&$ResizeSystemPartition ( (Get-Item -Path "$([Environment]::SystemDirectory)\Recovery\Winre.wim" -Force).Length + 300MB )

		$Target = &$CreateRecoveryPartition ($UseLetter)
	}

	<#

		Refresh the contents of the Recovery partion _AND_ BCD entries

	#>
	$UseLetter = &$PopulateRecoveryPartition ($UseLetter)
}

## Force enable the Recovery environment (which we may have just repaired ;-)
## This has no impact if it is already enabled.
ReagentC /enable

## Test if the Recovery environment is available on this system
Try {
	$ReagentC = ReagentC /Info
	If ( $($($ReagentC | Select-String -Pattern 'GLOBALROOT') -replace '^.*\s', '') -eq $Null ) {
		Write-Warning "A recovery partition was created but the Recovery environment is not enabled."
		Write-Warning ""
		If ($Logging) { Stop-Transcript }
		Exit 911
	  }

	$WindowsRELoader = "{" + $( $($ReagentC | Select-String -Pattern 'BCD') -replace "^.*:\s*", "" ) + "}"
	$WindowsRELoaderOptions = $(bcdedit /enum $WindowsRELoader | Select-String -Pattern '^device') -replace "^.*{", "{"
	$SanityCheck = $WindowsRELoaderOptions -eq $($(bcdedit /enum $WindowsRELoader | Select-String -Pattern '^osdevice') -replace "^.*{", "{")
	If ( !$SanityCheck) { Write-Warning "The device and osdevice parameters in BCD entry $WindowsRELoader should be identical!" }

	## Fix the BCD
	bcdedit /create "$WindowsRELoaderOptions" /Device | Out-Null
	bcdedit /set "$WindowsRELoaderOptions" ramdisksdipath \Recovery\WindowsRE\boot.sdi
	bcdedit /set "$WindowsRELoaderOptions" ramdisksdidevice partition="$UseLetter`:"
	bcdedit /set "$WindowsRELoader" device ramdisk=["$UseLetter`:"]\Recovery\WindowsRE\Winre.wim,"$WindowsRELoaderOptions"
	bcdedit /set "$WindowsRELoader" osdevice ramdisk=["$UseLetter`:"]\Recovery\WindowsRE\Winre.wim,"$WindowsRELoaderOptions"

  }
Catch {
	Write-Warning "The following error occured trying to get the current Recovery partition setup:"
	Write-Warning $_
	If ($Logging) { Stop-Transcript }
	Exit 911
 }


If ($Verbose) {	
	Write-Host $Separator -ForeGroundColor Green
	Write-Host $($ReagentC | Select-String -Pattern '.*Win.*RE.*' | Out-String ) -NoNewLine
	Write-Host $($ReagentC | Select-String -Pattern 'BCD' | Out-String ) -NoNewLine
	Write-Host ""
	Write-Host $Separator -ForeGroundColor Green

	## Note to self: someday, learn to use scopes in PowerShell and create script blocks ;-)
	## Display the Recovery Environment status on this system
	## Note: the display will reflect the partition letter assigned to the recovery partition.
	##       BCDEdit wil use the [\Device\HarddiskVolume...] name space otherwise which cannot
	##       be related easily to drive and partition numbers.
	if ($Target -ne $Null) {
		Write-Host $Separator -ForeGroundColor Green
		Write-Host ""
		Write-Host "Current Recovery partition BCD entries:" -ForegroundColor Green
		if ( [string]::IsNullOrEmpty($Target.DriveLetter) -or ($Target.DriveLetter -eq "`0") )
		  { ## Should we assign a drive letter here ?
			If ($(Read-Host "Enter 'Yes' to assign drive letter $UseLetter to the EXISTING recovery partition, anything else to continue").tolower().StartsWith('yes')) {
				$R = $Target | Set-Partition -NewDriveLetter $UseLetter
				  ## Hopefully, there is no side effect to this lack of "-Passthru"
				Write-Host "Assigned drive letter $UseLetter to", $Target.AccessPaths[$GlobalVol]
			  }
		  }
		else { ## 
			If ($(Read-Host "Enter 'Yes' to remove drive letter $($Target.DriveLetter) from the EXISTING recovery partition, anything else to continue").tolower().StartsWith('yes')) {
				## Remove the volume from he default name space
				$R = Remove-PartitionAccessPath -DiskNumber $Disk.Number -PartitionNumber $Target.PartitionNumber -AccessPath $Partition.AccessPaths[0]
			  }
			else {
				[char] $UseLetter = $($Target.DriveLetter)
				Write-Warning "Parameter override: using $UseLetter for the Recovery partition."
			  }
		  }
		Write-Host $(bcdedit /enum $WindowsRELoader | Out-String) -NoNewLine
		Write-Host $(bcdedit /enum $WindowsRELoaderOptions | Out-String) -NoNewLine
		Write-Host ""
		Write-Host $Separator -ForeGroundColor Green
	}
  }

	$SanityCheck = $SystemPartition.Guid -ne $Target.Guid
	If ( !$SanityCheck ) {
		Write-Error "Internal programming error: quit while you're ahead."
		If ($Logging) { Stop-Transcript }
		Exit 911

	  }
##
<#
	There is no simple activation sequence for restoring the recovery partition.
	
	Follow the discussion below tracking the contents of the partition as we go ...
#>
##
	  
if ($Target -ne $Null) {

	# Note: we don't know that the backup we took on entry is from the partition
	# selected by the user.
	$CaptureDir = $Target.AccessPaths[$Target.AccessPaths.Count - 1]
	$TargetSignature = &$BackupRecoveryPartition("Recovery Partition")
	If ($TargetSignature -ne $Null) { $OriginalWimSignature = $TargetSignature }
	
	$PreviousPartitionSize = $Target.Size

	Write-Host "Removing active WindowsRE partition..."
	Remove-Partition -DiskId $Target.DiskID -Offset $Target.Offset -Confirm:$False
<#
	This is an undocumented Windows feature: the Recovery image is also deleted
	from [Environment]::SystemDirectory)\Recovery:

		PS C:\Users\Administrator> dir "$([Environment]::SystemDirectory)\Recovery" -Force -Recurse

			Directory: C:\Windows\system32\Recovery

		Mode                 LastWriteTime         Length Name
		----                 -------------         ------ ----
		-a----         4/25/2023   2:05 PM            875 $PBR_Diskpart.txt
		-a----         4/25/2023   2:05 PM            468 $PBR_ResetConfig.xml
		-a----         6/17/2023  11:00 AM           1109 ReAgent.xml
	
	Discussion continues below ...
#>

	# Presume a minimum partition size : there is no implicit reduction of the existing WinRE partition size.
	&$ResizeSystemPartition( $PreviousPartitionSize )

	$NewTarget = &$CreateRecoveryPartition ($UseLetter)
	## Fix the BCD
	bcdedit /create "$WindowsRELoaderOptions" /Device | Out-Null
	bcdedit /set "$WindowsRELoaderOptions" ramdisksdipath \Recovery\WindowsRE\boot.sdi
	bcdedit /set "$WindowsRELoaderOptions" ramdisksdidevice partition="$UseLetter`:"
	bcdedit /set "$WindowsRELoader" device ramdisk=["$UseLetter`:"]\Recovery\WindowsRE\Winre.wim,"$WindowsRELoaderOptions"
	bcdedit /set "$WindowsRELoader" osdevice ramdisk=["$UseLetter`:"]\Recovery\WindowsRE\Winre.wim,"$WindowsRELoaderOptions"


<#
	Since a new Recovery partition is initialized, "REagentC /Enable" fails with the error:
		REAGENTC.EXE: The Windows RE image was not found.
	
	The recovery partition is restored from backup.

#>
	Try {
		Write-Host ""
		Write-Host "There is a backup of the Recovery partition in $($Env:Temp)."

		## Populate the Recovery partition from backup
		DISM /Apply-Image /ImageFile:$($Env:Temp+"\Recovery.wim") /Index:1 /ApplyDir:"$UseLetter`:" /Verify

		## Verify the restored environment
		$RestoredWimSignature = (Get-FileHash -Algorithm SHA256 -LiteralPath "$UseLetter`:\Recovery\WindowsRE\Winre.wim" -ErrorAction Stop).Hash
		Write-Host "Windows RE SHA256 signature:", $RestoredWimSignature

		If ($RestoredWimSignature -ne $OriginalWimSignature) {
			Write-Warning "The signature of the restored Recovery environment does not match the original"
			Write-Warning "Leaving the Windows RE status as is."
		  }
		else {
			## Again, force enable the Recovery environment
			REagentC /Enable
		  }
	  }
	Catch {
		## In case of any and all errors, exit and leave the backup.
		Write-Warning "The following error occured trying to restore the Recovery partition setup:"
		Write-Warning $_
		If ($Logging) { Stop-Transcript }
		Exit 911
	}
	
	## Display the Recovery Environment status on this system
	## Note: the display will reflect the partition letter assigned to the recovery partition.
	##       BCDEdit wil use the [\Device\HarddiskVolume...] name space otherwise which cannot
	##       be related easily to drive and partition numbers.
	If ($Verbose) {
		$ReagentC = ReagentC /Info
		Write-Host $Separator -ForeGroundColor Green
		Write-Host $($ReagentC | Select-String -Pattern '.*Win.*RE.*' | Out-String ) -NoNewLine
		Write-Host $($ReagentC | Select-String -Pattern 'BCD' | Out-String ) -NoNewLine
		Write-Host $(bcdedit /enum "$WindowsRELoader" | Out-String) -NoNewLine
		Write-Host $(bcdedit /enum "$WindowsRELoaderOptions" | Out-String) -NoNewLine
		If ( !$SanityCheck) { Write-Warning "The device and osdevice parameters should be identical!" }
		Write-Host $Separator -ForeGroundColor Green
	  }
##
	Write-Host "Removing drive letter $UseLetter..."
	$R = Remove-PartitionAccessPath -DiskNumber $SystemPartition.DiskNumber -PartitionNumber $NewTarget.PartitionNumber -AccessPath $NewTarget.AccessPaths[0]
	Write-Host ""
	If ($(Read-Host "Enter 'Yes' to delete the backup of the recovery partition in"$($Env:Temp)", anything else to continue").tolower().StartsWith('yes')) {
				Remove-Item -Path $($Env:Temp+"\Recovery.wim")
			  }
	Write-Host "Done!" -ForegroundColor Green
 }
else {
	## Dump all Recovery partitions accessible on this system
	## Remember: the "partiton" object has both MBR and UEFI type fields.
	$RecoveryPartitions = @( $(( Get-Partition | Where-Object { ($_.MBRType -eq 0x27) `
		-or ($_.GptType -eq "{de94bba4-06d1-4d40-a16a-bfd50179d6ac}") })) )
	If ($RecoveryPartitions -ne $Null) {
		Write-Host ""
		Write-Host "Recovery Partition(s) available on this system:"
		Write-Host ""
		 ForEach ($Partition in $RecoveryPartitions) { 
			Write-Host "     Disk number:", $Partition.DiskNumber
			Write-Host "Partition number:", $Partition.PartitionNumber
			Write-Host ( ($Partition.AccessPaths -replace "^","     Access path: ") -join "`n" )
			Write-Host ""
			}
	  }
	else { Write-Warning "There are no Recovery Partitions on this system." }
}
Write-Host ""

<#
	Display version and service pack level of current Recovery environment
#>

$WIM = $($(REagentC /Info | Select-String -Pattern 'GLOBALROOT') -replace '^.*\s', '')+"\Winre.wim"
If ($WIM -ne $Null) {
	Write-Host "Active WinRE version information"
	$WIMAttributes = Dism /Get-ImageInfo /ImageFile:$WIM /Index:1
	Write-Host $($WIMAttributes | Select-String -Pattern 'Version:' | Out-String ) -NoNewLine
	Write-Host $($WIMAttributes | Select-String -Pattern 'ServicePack' | Out-String ) -NoNewLine
	Write-Host ""
  }

If ($Logging) { Stop-Transcript }
Exit 0
