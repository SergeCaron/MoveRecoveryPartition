##******************************************************************
##
## Revision date: 2024.05.12
##
## Copyright (c) 2023-2024 PC-Évolution enr.
## This code is licensed under the GNU General Public License (GPL).
##
## THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF
## ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY
## IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR
## PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.
##
## Revision date: 2024.02.19
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
##		2024.02.09:	- Add a three seconds delay after creating a new partition to allow
##					  for some asynchronous system activity.
##					- Force explicit type CHAR return value in PopulateRecoveryPartition
##					- Add missing ":" when converting a \\?\GLOBALROOT\... path in BackupRecoveryPartition
##					- Correct some typos
##					- Do not assume the BitLocker name space is always present
##					- Display (hopefully) language independant WinRE version information
##		2024.02.11:	- Resolve type ambiguity of AccessPaths in partition CimInstance
##					- Leave a 1MB buffer when computing system partition size to prevent
##					  crash on some M2 devices.
##		2024.02.13:	- Fix use of $_.AccessPaths ... again!
##		2024.02.14:	- Do not resize if gain is less than 1MB
##		2024.02.16:	Allow another data partition as upper boundary to the system partition
##					(instead of disk size) and recoup free space between these two Partitions
##					if Recovery partition is located anywhere "above" this data partition.
##		2024.02.17:	- Identify Recovery Partition usng drive letter and path.
##					- Cleanup phantom drives, if at all possible.
##		2024.02.19:	- Identify phantom drives and cleanup related warnings.
##					- Add administrator privilegs check
##		2024.05.12: Verify that all disks are healthy. Resize-Partition will not Process
##					a dirty partition and the user may end up with a recover partition the
##					size of all remaining free space on the disk.
##		2025.02.18:	Force [UInt64] on all partition size computations
##		2025.02.19:	Formatter insert extra space in bcdedit /set command lines: rewrite
##
##******************************************************************

param (
	[Parameter(mandatory = $true, HelpMessage = 'Set the drive letter to use for the Recovery partition')]
	[char]$UseLetter,
	[parameter()]
	[switch]$Log,
	[parameter()]
	[switch]$Details,
	[parameter()]
	[string]$SourcesDir,
	[parameter()]
	[UInt64]$ExtendedSize
)
$Usage = @"

Usage: $($MyInvocation.MyCommand.Name) -UseLetter R -ExtendedSize <size> -Log -Details -SourcesDir <Path>

where:
`t-UseLetter`tis the drive letter that could be assigned to the recovery partition.
`t`t`tThis letter is only used during excution of the script. There is no default value.

`t-ExtendedSize`tis the size of the new recovery partition. If <size> is less than 1MB, <size>
`t`t`tis multiplied by 1MB, e.g. 600 implies 600MB.

`t-Log`t`tCreate a transcript log on the user's desktop.

`t-Details`tDisplay detailed information throughout execution of this script.

`t-SourcesDir`tis the directory of the Windows Installation Media containing Install.win.
`t`t`tThis is only used if this script finds no Recovery Environment on this system.

"@
<#	RestartVDS:
	---------------------------------------------------------------------------------------------------
	"Disk Management Services" can interfere in many ways if something is trying to display/manipulate partition data.
	Stop all Microsoft Consoles in which "Disk Management Services" are used
#>

Function RestartVDS {
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

<#	ResizeSystemPartition
	---------------------------------------------------------------------------------------------------
	Shrink/Expand the system partition, not knowing what will happen, and reserve enough space to 
	create a Recovery partition.
	
	Ask the user to confirm a minimum partition size.
#>

$ResizeSystemPartition = {
	Param ($FreeSizeRequired)

	# Compute a Recovery partition size, if it is going to be contiguous to the system partition.
	If ($FreeSizeRequired -gt 0) {
		If ($ExtendedSize -lt 1MB) { $ExtendedSize = $ExtendedSize * 1MB }
		$NewRESize = [UInt64] [math]::Max( [UInt64] $ExtendedSize, [UInt64] $FreeSizeRequired )
		
		## For Windows 11, the recommended recovery partition size is 990MB
		## See https://learn.microsoft.com/en-us/windows-hardware/manufacture/desktop/hard-drives-and-partitions?view=windows-11
		If ($NewRESize -lt 990MB) {
			If ($(Read-Host "Enter 'Yes' to use the recommended minimum partition size of 990MB for Windows RE, anything else to continue").tolower().StartsWith('yes')) `
			{	$NewRESize = [UInt64] 990MB }
		}
	}
 else { $NewRESize = [UInt64] 0 }

	Write-Host ""
	Write-Host "Computing new System partition size. Please be patient ..."
	$SupportedSize = (Get-PartitionSupportedSize -DiskId $SystemPartition.DiskID -Offset $SystemPartition.Offset).SizeMax
	$NewSize = $SupportedSize - $NewRESize - 1MB
	
	If ([Math]::Abs($SystemPartition.Size - $NewSize) -gt 1MB) {
		Try {
			Resize-Partition -DiskId $SystemPartition.DiskID -Offset $SystemPartition.Offset -Size $NewSize -ErrorAction Stop
		}
		Catch {
			Write-Warning "Error during System partition resize. You should keep the backup of the Recovery partition."
			Write-Warning "An attempt to restore a Recovery partition follows..."
		}
	}

}

<#	CreateRecoveryPartition
	---------------------------------------------------------------------------------------------------
	Create a new Recovery partition in the poper format, create a skeleton directory structure and
	display the Disk Management Console.
	
	Note: there is no notion of disk geometry here. The largest free space area on the system disk
	is alloated to this new partition: there is no way of computing an exact size based on the 
	available information.
#>

$CreateRecoveryPartition = {
	Param ($UseLetter)

	Write-Host ""
	Write-Host "Creating new Windows RE partition ..."

	RestartVDS

	If ($Disks.PartitionStyle -eq "GPT") {
		## "Microsoft Recovery" partitions are created "hidden" and even DISKPART cannot clear this attribute. The
		## message "Virtual Disk Service error: The object is not found." is displayed when the user attempts to
		## display the volume attributes.

		## In order to avoid Windows Explorer from offering to format the new partition, it is created hidden and later exposed.
		## There is quite a bit of grief trying to get a consistent behavior across all Windows versions when a drive letter
		## is specified. (On Windows 11/Server 2022 the volume is accessible using two drive letters, try the command
		## 				get-psdrive -PSProvider FileSystem		# ;-)
		## Create a basic data partition 
		## Create the partition "Hidden" to prevent darn thing ...

		Try {
			$NewTarget = New-Partition -DiskPath $Disks.Path -AssignDriveLetter -GptType "{ebd0a0a2-b9e5-4433-87c0-68b6b72699c7}" -UseMaximumSize -IsHidden -ErrorAction Stop
			## ... and expose it!
			Set-Partition -InputObject $NewTarget -IsHidden $False 
			$NewVolume = Format-Volume -Partition $NewTarget -Force -FileSystem NTFS -NewFileSystemLabel "Recovery"
			# There is an explicit -DriveLetter assignment when the partition is created.
  }
		Catch {
			Write-Warning $PSItem.Exception.Message
			If ($Logging) { Stop-Transcript }
			Exit 911
  }
		<#
        The instruction

            $NewTarget | Set-Partition -GptType "{de94bba4-06d1-4d40-a16a-bfd50179d6ac}" -IsHidden $True -NoDefaultDriveLetter $True

		would set the "Recovery" partition type and the GPT_BASIC_DATA_ATTRIBUTE_NO_DRIVE_LETTER (0x8000000000000000) attribute.
		Diskpart will display this partition as:

			Type    : de94bba4-06d1-4d40-a16a-bfd50179d6ac
			Hidden  : Yes
			Required: No
			Attrib  : 0X8000000000000000

		Windows (at least Disk Management console) will not display this as a Recovery partition unless the
        GPT_ATTRIBUTE_PLATFORM_REQUIRED (0x0000000000000001) is also set. There is no PowerShell equivalent
        and we must resort to Diskpart to do so.

	    Run DiskPart using an Here-String : the result is

            Type    : de94bba4-06d1-4d40-a16a-bfd50179d6ac
            Hidden  : Yes
            Required: Yes
            Attrib  : 0X8000000000000001
            Offset in Bytes: 93384081408

              Volume ###  Ltr  Label        Fs     Type        Size     Status     Info
              ----------  ---  -----------  -----  ----------  -------  ---------  --------
            * Volume 3     E   Recovery     NTFS   Partition    990 MB  Healthy    Hidden

        Although the partition above was created using drive letter X, diskpart displays some other drive
        letter on Windows 11 23H2 or no letter at all on earlier versions. Expect a phantom drive until
        reboot.
	#>
		$DiskpartLog = $(@"
select disk $($Disks.Number)
select partition $($NewTarget.PartitionNumber)
set id="de94bba4-06d1-4d40-a16a-bfd50179d6ac"
gpt attributes=0x8000000000000001
detail partition
exit
"@ | Diskpart)

		If ($Verbose) {
			Write-Host $($Capture | Out-String)
			Write-Host $DiskpartLog | Out-String
			Write-Host $($Capture | Out-String)
		}
	}
	else {
		## PowerShell has no MBRtype for "Microsoft Recovery" partitions
		## Create a basic data partition and change its type
		$NewTarget = New-Partition -DiskPath $Disks.Path -DriveLetter $UseLetter -MbrType IFS -UseMaximumSize
		$NewVolume = $NewTarget | Format-Volume -FileSystem NTFS -NewFileSystemLabel "Recovery"
		# | Set-Volume -DriveLetter $UseLetter is implicit when the partition is created.
		$NewTarget | Set-Partition -MbrType 0x27
	}

	Write-Host ""
	Write-Host "Restoring the Windows RE partition ..."

	## Do not override the user's assignment
	$RemoveTemporaryAccess = $False
	If ( [string]::IsNullOrEmpty($NewTarget.DriveLetter) -or ($NewTarget.DriveLetter -eq "`0") ) {
		$LocalTarget = $NewTarget | Add-PartitionAccessPath -AccessPath $($UseLetter + "`:") -PassThru
		$LocalLetter = $UseLetter
		Write-Host "Assigned drive letter $LocalLetter to the Recovery partition."
		$RemoveTemporaryAccess = $True
	}
	else {
		$LocalLetter = $NewTarget.DriveLetter
		Write-Host "Using drive letter $LocalLetter to access  the Recovery partition."
	}

	## Create a skeleton directory structure
	New-Item -Path "$LocalLetter`:\" -Name "Recovery" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null
	New-Item -Path "$LocalLetter`:\Recovery" -Name "WindowsRE" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null
	New-Item -Path "$LocalLetter`:\Recovery" -Name "Logs" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null

	# Drop the temporary assignment
	If ($RemoveTemporaryAccess) {
		$LocalTarget = $LocalTarget | Remove-PartitionAccessPath -AccessPath $($UseLetter + "`:") -PassThru
	}

	Return $NewTarget
}

<#	LocateWindowsREImage
	---------------------------------------------------------------------------------------------------
	Presume the Recovery Environment is disabled.
	
	REagentC stores a copy of the Windows RE image file in $([Environment]::SystemDirectory)\Recovery\
	
	If no such file exists, restore a copy from the Windows installation media (install.wim).
	
	Note that Winre.wim may have been updated : direct the user to "KB5034957: Updating the WinRE partition on deployed devices"
	
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

$LocateWindowsREImage = {
	[OutputType([String])]
	Param ($RELocation)

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
					Remove-Item -Path "$([Environment]::SystemDirectory)\Recovery\Winre.wim" -Force -Confirm:$False `
						-ErrorAction Stop
				}
				Catch {
					# Do Nothing!
				}
				Finally {
					Copy-Item  -LiteralPath "$Decoy\Windows\System32\Recovery\Winre.wim" `
						-Destination "$([Environment]::SystemDirectory)\Recovery\Winre.wim" -Force -Confirm:$False `
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

	Return [String] $RELocation
}

<#	PopulateRecoveryPartition
	---------------------------------------------------------------------------------------------------
	Mimic the behavior of ReagentC /enable while relocating the Recovery environment.

#>

$PopulateRecoveryPartition = {
	[OutputType([CHAR])]
	Param ($UseLetter)
	<#
	We need to ensure a measure of health before anything else.

#>
	## Do not override the user's assignment
	$RemoveTemporaryAccess = $False
	If ( [string]::IsNullOrEmpty($Target.DriveLetter) -or ($Target.DriveLetter -eq "`0") ) {
		$LocalTarget = $Target | Add-PartitionAccessPath -AccessPath $($UseLetter + "`:") -PassThru
		$RELetter = $UseLetter
		Write-Host "Assigned drive letter $RELetter to the Recovery partition."
		$RemoveTemporaryAccess = $True
	}
	else {
		$RELetter = $Target.DriveLetter
		Write-Host "Using drive letter $RELetter to access  the Recovery partition."
	}

	## Create / Repair the contents of the Recovery partition
	New-Item -Path "$RELetter`:\" -Name "Recovery" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null
	New-Item -Path "$RELetter`:\Recovery" -Name "WindowsRE" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null
	New-Item -Path "$RELetter`:\Recovery" -Name "Logs" -ItemType "directory" -ErrorAction SilentlyContinue | Out-Null
	
	## Install a copy of the Recovery image in this location
	Copy-Item -Path "$([Environment]::SystemDirectory)\Recovery\Winre.wim" `
		-Destination "$RELetter`:\Recovery\WindowsRE\Winre.wim" -Force -Confirm:$False `
		-ErrorAction SilentlyContinue
	
	## Drop the Recovery entry
	bcdedit /deletevalue "{current}" recoverysequence

	## Rebuild the structures required to reboot into the Recovery environment
	REagentC /SetREImage /Path "$RELetter`:\Recovery\WindowsRE"
	
	## This does NOT enable the Recovery environment.
	
	If ($RemoveTemporaryAccess) {
		$LocalTarget = $LocalTarget | Remove-PartitionAccessPath -AccessPath $($UseLetter + "`:") -PassThru
	}

	Return [CHAR] $UseLetter
}

<#	BackupRecoveryPartition
	---------------------------------------------------------------------------------------------------

	Backup a recovery partition : we don't know the kind of customization or the up-to-date status
	of its contents. This is preferable to any attempt to rebuild this data.
	
	Note: Backup is NOT always relative to ...\Recovery

#>

$BackupRecoveryPartition = {
	Param ($Description)
	
	If ($CaptureDir -ne $Null) {
		Write-Host ""
		Write-Host "Testing $CaptureDir ($Description). Please be patient ..."
		
		$RemoveTemporaryAccess = $False
		
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

			# Get the drive letter from the partition instead of the volume
			$TargetGuid = $CaptureDir -replace "^.+(\{.*\}).*", '$1'
			$TargetPartition = Get-Partition | Where-Object { $_.Guid -eq $TargetGuid }

			If ( [string]::IsNullOrEmpty($TargetPartition.DriveLetter) -or ($TargetPartition.DriveLetter -eq "`0") ) {
				$TargetPartition = $TargetPartition | Add-PartitionAccessPath -AccessPath $($UseLetter + "`:") -PassThru
				$RELocation = $UseLetter
				Write-Host "Assigned drive letter $UseLetter to", $CaptureDir
				$RemoveTemporaryAccess = $True
			}
			else {
				$RELocation = $TargetPartition.DriveLetter
				Write-Host "Using drive letter $RELocation to access", $CaptureDir
			}
			$RELocation = $RELocation + ":"
		}
		elseif ($CaptureDir.StartsWith('\\?\GLOBALROOT\device\', 'CurrentCultureIgnoreCase')) {
			$DiskNumber = $CaptureDir -replace "^\\\\\?\\GLOBALROOT\\device\\harddisk(\d+)\\partition\d+\\.*$", '$1'
			$PartitionNumber = $CaptureDir -replace "^\\\\\?\\GLOBALROOT\\device\\harddisk\d+\\partition(\d+)\\.*$", '$1'
			$TargetDir = $CaptureDir -replace "^\\\\\?\\GLOBALROOT\\device\\harddisk\d+\\partition\d+(\\.*)$", '$1'
			$TargetPartition = Get-Partition | Where-Object { ($_.DiskNumber -eq $DiskNumber) -and ($_.PartitionNumber -eq $PartitionNumber) }

			If ( [string]::IsNullOrEmpty($TargetPartition.DriveLetter) -or ($TargetPartition.DriveLetter -eq "`0") ) {
				$TargetPartition = $TargetPartition | Add-PartitionAccessPath -AccessPath $($UseLetter + "`:") -PassThru
				$RELocation = $UseLetter
				Write-Host "Assigned drive letter $UseLetter to", $CaptureDir
				$RemoveTemporaryAccess = $True
			}
			else {
				$RELocation = $TargetPartition.DriveLetter
				Write-Host "Using drive letter $RELocation to access", $CaptureDir
			}
			$RELocation = $RELocation + ":"
		}
		else {
			$RELocation = $CaptureDir
		}

		# Get-FileHash can access the image fle but does not error on missing files!
		# The directory structure is not consistent between the enabled and disabled states :
		Try {
			# a) Try the default location ...
			$OriginalWimSignature = (Get-FileHash -Algorithm SHA256 -LiteralPath "$RELocation\Recovery\WindowsRE\Winre.wim" -ErrorAction Stop).Hash
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
		If ($RemoveTemporaryAccess) {
			$TargetPartition = $TargetPartition | Remove-PartitionAccessPath -AccessPath $($UseLetter + "`:") -PassThru
		}
		
	}
	Return $OriginalWimSignature # which may be unchanged
}

<#	Main logic
	---------------------------------------------------------------------------------------------------
	
	If the Recovery environment is currently disabled:
	- the user can identify ONE recovery partition located AFTER the system partition and this partition
	  will be removed.
	- if the Recovery environment is not available in System32\Recovery, it can be restored from the
	  Windows distribution media (DVDm ISO, etc.)
	- the system partition is extended to the maximum size permitted less the space required for the
	  recovery partition.
	- a recovery partition is created and the Recovery environment is enabled.
	
	If the Recovery environment is currently enabled:
	- the user can identify ONE recovery partition located AFTER the system partition.
	- if no partition is selected, the script enumerates all Recovery partitions on the system disk.
	- a backup of the selected partition is created: this includes all customizations of the existing
	  partition.
	- the system partition is extended to the maximum size permitted less the space required for the
	  recovery partition.
	- a recovery partition is created and the backup of the existing partition is restored.
	- the BCD entries are updated.
	
	On exit, the version information of the ACTIVE Recovery environment is displayed.
	
	To move the Recovery environment from the system partition to its own partition, simply disable
	the Recovery Environment before invoking this script.
	
	To extend the system partition following a disk restore, simply invoke the script without
	disabling the Recovery Environment.
	
	Caution: the script does not examine the system's disk partition allocation. The typical use case
	is a disk containing a single system partition followed by the Recovery partition, both of which
	should occupy the entire disk. Unexpected results occur when there is free space following a data
	partition on the system disk: user beware.
	
	You have the option to display (and log) (hopefully) all actions of this script.
#>

## Are we logging this ?
[Boolean]$Logging = $Log.IsPresent

If ($Logging) { Start-Transcript -Path "$($env:USERPROFILE)\Desktop\RecoveryPartitionMaintenance.txt" -Append }

[char]$UseLetter = "$UseLetter".ToUpper() # Traditionnal ;-)

## Some sanity checks
If ( $Null -ne $(Get-PSDrive "$UseLetter" -PSProvider FileSystem -ErrorAction SilentlyContinue)) `
{
	Write-Host "Drive $UseLetter is already in use."
	If ($Logging) { Stop-Transcript }
	Exit 911
}

$Concerns = @(Get-Volume | Where-Object { $_.HealthStatus -ne "Healthy" })
If ($Concerns.Count -gt 0) {
	Write-Warning "Repair these drives before relocating the recovery partition:"
	$Concerns | Format-Table DriveLetter, FriendlyName, FileSystemType, DriveType, HealthStatus, OperationalStatus
	If ($Logging) { Stop-Transcript }
	Exit 911
}

# Get the ID and security principal of the current user account
$myWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent()
$myWindowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($myWindowsID)

# Get the security principal for the administrator role
$adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator

# Check to see if we are currently running as an administrator
if (!$myWindowsPrincipal.IsInRole($adminRole)) {
 Write-Host "Administrative privileges are required to run this script."
	If ($Logging) { Stop-Transcript }
	Exit 911
}

## Determine how much will be displayed
[Boolean]$Verbose = $Details.IsPresent

## Create a separator for verbose output
$Separator = "-" * $Host.UI.RawUI.WindowSize.Width

## Warn user and display the current Windows Build
Write-Host ""
Write-Warning "-> Drive letter $UseLetter is used to manipulate the Recovery partition. This may conflict"
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
Write-Warning "See https://en.wikibooks.org/wiki/Windows_10%2B_Recovery_Environment_(RE)_Notes"
Write-Warning "for a general revision of the Recovery Environment since the release of Windows 10."
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

	The « Root\cimv2\Security\MicrosoftVolumeEncryption » Namespace may not exist.
#>

$SystemDrive = (Get-WmiObject Win32_OperatingSystem).SystemDrive
$BitLocker = Get-WmiObject -Namespace "Root\cimv2\Security\MicrosoftVolumeEncryption" -Class "Win32_EncryptableVolume" `
	-Filter "DriveLetter = '$SystemDrive'" -ErrorAction SilentlyContinue
If (-not $BitLocker) {
	Write-Host "No BitLocker protection found for drive $SystemDrive."
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

$SystemPath = $SystemDrive + "\"
$SystemPartition = $Null
ForEach ($Partition in $(Get-Partition)) {
	If ($Partition.AccessPaths.Count -gt 0) {
		If ($Partition.AccessPaths.Contains($SystemPath)) {
			$SystemPartition = $Partition
		}
	}
}
If ($Null -eq $SystemPartition) {
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
If ($Null -eq $RELocation) {
	Write-Warning "The Recovery environment is currently disabled."
	$CaptureDir = &$LocateWindowsREImage ($RELocation)
	# Presume there is a valid WindowsRE image file available
	$RELocation = $CaptureDir
}
else {
	# The Recovery environment may be accessible through GLOBALROOT
	$CaptureDir = $RELocation -replace "(^*)\\WindowsRE", $1
}
$OriginalWimSignature = &$BackupRecoveryPartition("Recovery Environment")
Write-Host ""

If ($Verbose -and ($Null -ne $RELocation)) {
	Write-Host $Separator -ForegroundColor Green
	Write-Host "Windows RE image attributes:"
	Write-Host ""
	Dism /Get-ImageInfo /ImageFile:"$RELocation\winre.wim" /index:1
	Write-Host $Separator -ForegroundColor Green
}

ForEach ($Disk in $Disks) {
	$Partitions = Get-Partition -DiskNumber $SystemPartition.DiskNumber
	## Caution: Get-PartitionSupportedSize has the side effect of starting the defragmentation service ("Optimixe Drive")
	## and returns an array of maximum sizes that is not always in the same order as the partition table!
	$MaxSize = $Disk.Size - $Disk.AllocatedSize
	If ($MaxSize -gt 1GB) {
		## Announce possible gains
		Write-Host "Approximately", $([int] [Math]::floor($($MaxSize / 1GB))), "GB can be allocated on this drive." -ForegroundColor Green
		Write-Host "(Presuming all free space is contiguous at the end of this disk.)"  -ForegroundColor Green
		Write-Host
		If ($Verbose) {
			## Provide a short description of this disk
			$Disk | Format-List PartitionStyle, FriendlyName, GUID
			## Dump the partition table: don't rely on the order of the display
			$Partitions | Format-Table Type, IsSystem, AccessPaths, GptType, MBRType, Size
		}
	}
	If ($Partitions.Count -gt 1) {
		ForEach ($Partition in $Partitions) {
			If ($Partition.AccessPaths.Count -gt 0) {
				If ( (($Disk.PartitionStyle -eq "MBR" -and $Partition.MBRType -eq 0x27) `
							-or ($Disk.PartitionStyle -eq "GPT" -and $Partition.GptType -eq "{de94bba4-06d1-4d40-a16a-bfd50179d6ac}") ) `
						-and $Partition.Offset -gt $SystemPartition.Offset ) {

					## Get-Children cannot access a volume using its \\?\Volume... designation
					$RemoveTemporaryAccess = $False
					If ( [string]::IsNullOrEmpty($Partition.DriveLetter) -or ($Partition.DriveLetter -eq "`0") ) {
						$LocalTarget = $Partition | Add-PartitionAccessPath -AccessPath $($UseLetter + "`:") -PassThru
						$LocalLetter = $UseLetter
						$RemoveTemporaryAccess = $True
					}
					else {
						$LocalLetter = $Partition.DriveLetter
					}

					$ThisRoot = If ($Partition.AccessPaths.Gettype().BaseType.Name -eq "Array") `
					{ $Partition.AccessPaths[$Partition.AccessPaths.Count - 1] }
					else { $Partition.AccessPaths }
					Write-Host "Testing:", $ThisRoot, "Using $LocalLetter`:"
					
					Try {
						## Note: Test-Path will not return a value for hidden/system objects.
						## In the default name space (drive letter), a suffix "\..\" is required to find the directory, which does not seem coherent
						## We have direct access to the directory using the global name space...
						## ... and this may point to a healthy Recovery partition which has a number of subdirectories
						$WinRE = Get-ChildItem -Attributes Directory, Directory+Hidden -Path $($LocalLetter + "`:Recovery\") `
							-Directory -Depth 0 -Exclude "System Volume Information" -Force -ErrorAction Stop
						If ($WinRe.Count -ge 1) {
							Write-Host $ThisRoot, "is an apparent Recovery partition" -ForegroundColor Green
							If ($(Read-Host "Enter 'Yes' to select this partition as the target recovery partition, anything else to continue").tolower().StartsWith('yes')) {
								## Note: the above manipulations do not effect the local variables
								$Target = $Partition
							}
						}
					}
					Catch {
						If ($Verbose) { Write-Host "-> Ignored" }
					}

					## Drop the temporary assignment: there may be more than one Recovery partition
					If ($RemoveTemporaryAccess) {
						$LocalTarget = $LocalTarget | Remove-PartitionAccessPath -AccessPath $($UseLetter + "`:") -PassThru
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
$RELocated = $False
If ($RELocation -eq "$([Environment]::SystemDirectory)\Recovery") {
	If ($Null -ne $Target) {
		## Do not attempt to salvage the user identified Recovery partition
		Write-Host "Removing inactive WindowsRE partition..."
		If (!( [string]::IsNullOrEmpty($Target.DriveLetter) -or ($Target.DriveLetter -eq "`0") )) `
		{ Remove-PartitionAccessPath -DiskNumber $Disk.Number -PartitionNumber $Target.PartitionNumber -AccessPath $($Target.DriveLetter + ":\") }
		Remove-Partition -DiskId $Target.DiskID -Offset $Target.Offset -Confirm:$False
		$Target = $Null
	}

	# Presume a minimum partition size (see MS KB5034439/KB5034441)
	&$ResizeSystemPartition ( [UInt64] (Get-Item -Path "$([Environment]::SystemDirectory)\Recovery\Winre.wim" -Force).Length + 300MB )
	$SystemPartition = $SystemPartition | Get-Partition

	$Target = &$CreateRecoveryPartition ($UseLetter)
	$RELocated = $True

	# This SHOULD create all required BCD entries corresponding to the new location and populate the partition
	ReagentC /enable

	# Get the BCD entry from the ReAgent configuraton file
	$REAgentData = New-Object xml
	$REAgentData.Load( (Convert-Path $([System.Environment]::SystemDirectory + "\Recovery\ReAgent.xml") ) )
	$WindowsRELoader = $REAgentData.WindowsRE.WinreBCD.id

	# Note: ReagentC/Windows manipulate the access path of the new RE partition. Don't trust anything...
	#		The partition can still be accessed using $UseLetter later on
	$CaptureDir = If ($Target.AccessPaths.Gettype().BaseType.Name -eq "Array") `
	{ $Target.AccessPaths[$Target.AccessPaths.Count - 1] }
	else	{ $Target.AccessPaths }
	$TargetSignature = &$BackupRecoveryPartition("Reconstructed Recovery Partition")

	If ($TargetSignature -ne $OriginalWimSignature) {
		Write-Warning "The signature of the restored Recovery environment does not match the original"
		Write-Warning "Leaving the Windows RE status as is."
	}
	
	# Claim free space contiguous to the system partition.
	&$ResizeSystemPartition( [UInt64] 0 )
}
else {
	## Force enable the Recovery environment (which is still untouched ;-)
	## This has no impact if it is already enabled.
	ReagentC /enable

	$REAgentData = New-Object xml
	$REAgentData.Load( (Convert-Path $([System.Environment]::SystemDirectory + "\Recovery\ReAgent.xml") ) )
	$WindowsRELoader = $REAgentData.WindowsRE.WinreBCD.id
}

## Test if the Recovery environment is available on this system
Try {
	If ( $REAgentData.WindowsRE.WinreBCD.id -eq "{00000000-0000-0000-0000-000000000000}" ) {
		Write-Warning "Cannot enable the Recovery environment. Please review the -SourcesDir option."
		Write-Warning ""
		If ($Logging) { Stop-Transcript }
		Exit 911
	}

	$WindowsRELoaderOptions = $(bcdedit /enum $WindowsRELoader | Select-String -Pattern '^device') -replace "^.*{", "{"
	$SanityCheck = $WindowsRELoaderOptions -eq $($(bcdedit /enum $WindowsRELoader | Select-String -Pattern '^osdevice') -replace "^.*{", "{")
	If ( !$SanityCheck) { Write-Warning "The device and osdevice parameters in BCD entry $WindowsRELoader should be identical!" }

}
Catch {
	Write-Warning "The following error occured trying to get the current Recovery partition setup:"
	Write-Warning $_
	If ($Logging) { Stop-Transcript }
	Exit 911
}


If ($Verbose) {
	$ReagentC = ReagentC /Info
	Write-Host $Separator -ForegroundColor Green
	Write-Host $($ReagentC | Select-String -Pattern '.*Win.*RE.*' | Out-String ) -NoNewline
	Write-Host $($ReagentC | Select-String -Pattern 'BCD' | Out-String ) -NoNewline
	Write-Host ""
	Write-Host $Separator -ForegroundColor Green

	## Display the Recovery Environment status on this system
	## Note: the display will reflect the partition letter assigned to the recovery partition.
	##       BCDEdit wil use the [\Device\HarddiskVolume...] name space otherwise which cannot
	##       be related easily to drive and partition numbers.
	if ($Null -ne $Target) {
		Write-Host ""
		Write-Host "Current Recovery partition BCD entries:" -ForegroundColor Green

		## Attempt to beautify the output without overriding the user's assignment
		$RemoveTemporaryAccess = $False
		If ( [string]::IsNullOrEmpty($Target.DriveLetter) -or ($Target.DriveLetter -eq "`0") ) {
			$Target = $Target | Add-PartitionAccessPath -AccessPath $($UseLetter + "`:") -PassThru
			$RELetter = $UseLetter
			Write-Host "Assigned drive letter $RELetter to the EXISTING Recovery partition."
			$RemoveTemporaryAccess = $True
		}
		else {
			$RELetter = $Target.DriveLetter
			Write-Host "Using drive letter $RELetter to access the EXISTING Recovery partition."
		}

		Write-Host $(bcdedit /enum $WindowsRELoader | Out-String) -NoNewline
		Write-Host $(bcdedit /enum $WindowsRELoaderOptions | Out-String) -NoNewline

		If ($RemoveTemporaryAccess) {
			$Target = $Target | Remove-PartitionAccessPath -AccessPath $($UseLetter + "`:") -PassThru
		}
		Write-Host ""
	}
	Write-Host $Separator -ForegroundColor Green
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
	  
if (($Null -ne $Target) -and !$RElocated) {

	# Note: we don't know that the backup we took on entry is from the partition
	# selected by the user.
	$CaptureDir = If ($Target.AccessPaths.Gettype().BaseType.Name -eq "Array") `
	{ $Target.AccessPaths[$Target.AccessPaths.Count - 1] }
	else	{ $Target.AccessPaths }
	$TargetSignature = &$BackupRecoveryPartition("Recovery Partition")
	If ($Null -ne $TargetSignature) { $OriginalWimSignature = $TargetSignature }
	
	$PreviousPartitionSize = [UInt64] $Target.Size

	Write-Host "Removing active WindowsRE partition..."
	If (!( [string]::IsNullOrEmpty($Target.DriveLetter) -or ($Target.DriveLetter -eq "`0") )) `
	{ Remove-PartitionAccessPath -DiskNumber $Disk.Number -PartitionNumber $Target.PartitionNumber -AccessPath $($Target.DriveLetter + ":\") }
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
	$SystemPartition = $SystemPartition | Get-Partition

	$NewTarget = &$CreateRecoveryPartition ($UseLetter)
	## Do not override the user's assignment
	$RemoveTemporaryAccess = $False
	If ( [string]::IsNullOrEmpty($NewTarget.DriveLetter) -or ($NewTarget.DriveLetter -eq "`0") ) {
		$LocalTarget = $NewTarget | Add-PartitionAccessPath -AccessPath $($UseLetter + "`:") -PassThru
		$LocalLetter = $UseLetter
		Write-Host "Assigned drive letter $LocalLetter to the Recovery partition."
		$RemoveTemporaryAccess = $True
	}
	else {
		$LocalLetter = $NewTarget.DriveLetter
		Write-Host "Using drive letter $LocalLetter to access  the Recovery partition."
	}

	## Fix the BCD
	bcdedit /create "$WindowsRELoaderOptions" /Device | Out-Null
	bcdedit /set "$WindowsRELoaderOptions" ramdisksdipath \Recovery\WindowsRE\boot.sdi
	bcdedit /set "$WindowsRELoaderOptions" ramdisksdidevice partition="$LocalLetter`:"
	bcdedit /set "$WindowsRELoader" device ramdisk=["$LocalLetter`:"]\Recovery\WindowsRE\Winre.wim",$WindowsRELoaderOptions"
	bcdedit /set "$WindowsRELoader" osdevice ramdisk=["$LocalLetter`:"]\Recovery\WindowsRE\Winre.wim",$WindowsRELoaderOptions"

	<#
	The recovery partition is restored from backup.

	Since a new Recovery partition is initialized without the need for the ...\System21\Recovery intermediate storage,
	"REagentC /Enable" fails with the error:
		REAGENTC.EXE: The Windows RE image was not found.
	
#>
	Try {
		Write-Host ""
		Write-Host "There is a backup of the Recovery partition in $($Env:Temp)."

		## Populate the Recovery partition from backup
		DISM /Apply-Image /ImageFile:$($Env:Temp+"\Recovery.wim") /Index:1 /ApplyDir:"$LocalLetter`:" /Verify

		## Verify the restored environment
		$RestoredWimSignature = (Get-FileHash -Algorithm SHA256 -LiteralPath "$LocalLetter`:\Recovery\WindowsRE\Winre.wim" -ErrorAction Stop).Hash
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
		Write-Host $Separator -ForegroundColor Green
		Write-Host $($ReagentC | Select-String -Pattern '.*Win.*RE.*' | Out-String ) -NoNewline
		Write-Host $($ReagentC | Select-String -Pattern 'BCD' | Out-String ) -NoNewline
		Write-Host $(bcdedit /enum "$WindowsRELoader" | Out-String) -NoNewline
		Write-Host $(bcdedit /enum "$WindowsRELoaderOptions" | Out-String) -NoNewline
		If ( !$SanityCheck) { Write-Warning "The device and osdevice parameters should be identical!" }
		Write-Host $Separator -ForegroundColor Green
	}
	##
	Write-Host "Removing drive letter $UseLetter..."
	## Do NOT trust the drive letter: ReagentC is acting on its own!
	$Lint = Get-Partition -DiskNumber $SystemPartition.DiskNumber -PartitionNumber $NewTarget.PartitionNumber
	If ( !([string]::IsNullOrEmpty($Lint.DriveLetter) -or ($Lint.DriveLetter -eq "`0")) ) {
		# OK, we have a real drive letter, get rid of it
		$ThisPath = If ($Lint.AccessPaths.Gettype().BaseType.Name -eq "Array") `
		{ $Lint.AccessPaths[0] }
		else	{ $Lint.AccessPaths }
		$R = Remove-PartitionAccessPath -DiskNumber $SystemPartition.DiskNumber -PartitionNumber $Lint.PartitionNumber -AccessPath $ThisPath
	}
	Write-Host ""
	If ($(Read-Host "Enter 'Yes' to delete the backup of the recovery partition in"$($Env:Temp)", anything else to continue").tolower().StartsWith('yes')) {
		Remove-Item -Path $($Env:Temp + "\Recovery.wim")
	}

	# Claim free space contiguous to the system partition.
	&$ResizeSystemPartition( [UInt64] 0 )
}

## Dump all Recovery partitions accessible on this system
## Remember: the "partiton" object has both MBR and UEFI type fields.
$RecoveryPartitions = @( $(( Get-Partition | Where-Object { ($_.MBRType -eq 0x27) `
					-or ($_.GptType -eq "{de94bba4-06d1-4d40-a16a-bfd50179d6ac}") })) )
If ($Null -ne $RecoveryPartitions) {
	Write-Host ""
	Write-Host "Recovery Partition(s) available on this system:"
	Write-Host ""
	ForEach ($Partition in $RecoveryPartitions) { 
		Write-Host "     Disk number:", $Partition.DiskNumber
		Write-Host "Partition number:", $Partition.PartitionNumber
		Write-Host ( ($Partition.AccessPaths -replace "^", "     Access path: ") -join "`n" )
		If ( !([string]::IsNullOrEmpty($Partition.DriveLetter) -or ($Partition.DriveLetter -eq "`0")) ) {
			# OK, we have a real drive letter, get rid of it
			If ($Partition.AccessPaths.Gettype().BaseType.Name -eq "Array") `
			{ $ThisPath = $Partition.AccessPaths[0] }
			else	{ $ThisPath = $Partition.AccessPaths }

			Write-Warning "Removing access path $ThisPath"
			$R = Remove-PartitionAccessPath -DiskNumber $Partition.DiskNumber -PartitionNumber $Partition.PartitionNumber -AccessPath $ThisPath
		}
		Write-Host ""
	}
}
else {
	Write-Warning "There are no Recovery Partitions on this system."
	Write-Host ""
}

## Identify phantom drives
$DriveLetters = (Get-PSDrive -PSProvider FileSystem).Name
$SuggestReboot = $False
ForEach ($DriveLetter in $DriveLetters) {
	If ( $DriveLetter.Length -eq 1 ) {
		If ( "$DriveLetter`:\" -eq (Get-PSDrive $DriveLetter).Root ) {
			Try { $Void = Get-Volume -DriveLetter $Driveletter -ErrorAction Stop }
			Catch {
				Write-Warning "Drive $DriveLetter`:\ is a phantom drive."
				$SuggestReboot = $True 
			}
		}
	}
}
If ($SuggestReboot ) {
	Write-Warning "You should reboot this system to remove phantom drives potentially created by PowerShell during this script."
	Write-Host ""
}

<#
	Display version and service pack level of current Recovery environment
#>

$WIM = $($(REagentC /Info | Select-String -Pattern 'GLOBALROOT') -replace '^.*\s', '') + "\Winre.wim"
If ($Null -ne $WIM) {
	Write-Host "Active WinRE version information" -ForegroundColor Green
	Write-Host $Separator -ForegroundColor Green
	$WIMAttributes = Dism /Get-ImageInfo /ImageFile:$WIM /Index:1
	Write-Host $($WIMAttributes | Select-String -Pattern 'Version' | Out-String ).Trim() -ForegroundColor Green
	Write-Host $($WIMAttributes | Select-String -Pattern 'Pack' | Out-String ).Trim() -ForegroundColor Green
	Write-Host ""
}

## Display the Disk Management Consoles
diskmgmt.msc

Write-Host "Done!" -ForegroundColor Green

If ($Logging) { Stop-Transcript }
Exit 0
