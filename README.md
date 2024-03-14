# MoveRecoveryPartition
 Setup/Repair a Windows Recovery partition.

If the Recovery environment is currently disabled:
- the user can identify ONE recovery partition located AFTER the system partition and this partition
  will be removed.
- if the Recovery environment is not available in System32\Recovery, it can be restored from the
  Windows distribution media (DVD, ISO, etc.)
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
disabling the Recovery Environment: this will protect any customization that may exist on
the restored Recovery partition.

Caution: the script does not examine the system's disk partition allocation. The typical use case
is a disk containing a single system partition followed by the Recovery partition, both of which
should occupy the entire disk. Unexpected results occur when there is free space located beyond a
data partition on the system disk: **user beware**. If that should be the case, the system partition
is made contiguoous to the first data partition.

You have the option to display (and log) (hopefully) all actions of this script.

You have the option to extend the existing / to specify the size of the new Recovery partition.

# Usage:

You must run PowerShell as an administrator. Invoke the script (relative path not specified here).
Script parameters are:

	MoveRecoveryPartition.ps1 -UseLetter R -ExtendedSize <size> -Log -Details -SourcesDir <Path>

where:

- -UseLetter:      is the drive letter that will be assigned to the recovery partition.
				This letter is only used during excution of the script. There is no default value.

- -ExtendedSize:   is the size of the new recovery partition. If `<size`> is less than 1MB, `<size`>
				is multiplied by 1MB, e.g. 600 implies 600MB.

- -Log:            Create a transcript log on the user's desktop. The log is named C:\Users\<LoggedOnUser>\Desktop\RecoveryPartitionMaintenance.txt

- -Details:        Display detailed information throughout execution of this script.

- -SourcesDir:     is the directory of the Windows Installation Media containing Install.win.
				This is only used if this script finds no Recovery Environment on this system.

The Recovery Environment can be brought up-to-date using Microsoft's KB5034957.

See *[Windows 10+ Recovery Environment (RE) Notes](https://en.wikibooks.org/wiki/Windows_10%2B_Recovery_Environment_(RE)_Notes)*
for a general revision of the Recovery Environment since the release of Windows 10.

**In any case: USE AT YOUR OWN RISK!**

## Sample console output

```
PS C:\Users\ThisUser> D:\MoveRecoveryPartition.ps1 -UseLetter R -ExtendedSize 1536

WARNING: -> Drive letter R is used to manipulate the Recovey partition. This may conflict
WARNING: -> with your current drive assignments.

WARNING: This script attempts to relocate the Recovery partition contiguous to the end of the
WARNING: System partition. This may extend the System partition by allocating all free space
WARNING: available following a disk size increase. This may shrink the system partition to
WARNING: increase the size of the recovery partition (see MS KB5034439/KB5034441).
WARNING:
WARNING: The script will also repair a disabled recovery partition.
WARNING: Optionally, the script may create a recovery partition using the Windows installation
WARNING: media.
WARNING:
WARNING: The Recovery Environment can be brought up-to-date using Microsoft's KB5034957.
WARNING:
WARNING: In any case: USE AT YOUR OWN RISK!
WARNING:
Enter 'Yes' to continue, anything else to exit: yes

Usage: MoveRecoveryPartition.ps1 -UseLetter R -ExtendedSize <size> -Log -Details -SourcesDir <Path>

where:
        -UseLetter      is the drive letter that will be assigned to the recovery partition.
                        This letter is only used during excution of the script. There is no default value.

        -ExtendedSize   is the size of the new recovery partition. If <size> is less than 1MB, <size>
                        is multiplied by 1MB, e.g. 600 implies 600MB.

        -Log            Create a transcript log on the user's desktop.

        -Details        Display detailed information throughout execution of this script.

        -SourcesDir     is the directory of the Windows Installation Media containing Install.win.
                        This is only used if this script finds no Recovery Environment on this system.



Microsoft Windows NT 10.0.22631.0

BitLocker is off on this system.                                                                                                                                                                                                                                                                                                                                        Testing \\?\GLOBALROOT\device\harddisk0\partition4\Recovery (Recovey Environment). Please be patient ...                

Assigned drive letter R to \\?\GLOBALROOT\device\harddisk0\partition4\Recovery                                                                                                                                                                  \\?\Volume{961e043d-faa9-4e15-9ad8-8762cd29e195}\ is an apparent Recovery partition                                     
Enter 'Yes' to select this partition as the target recovery partition, anything else to continue: yes                   

REAGENTC.EXE: Operation Successful.                                                                                     
The operation completed successfully.
The operation completed successfully.
The operation completed successfully.
The operation completed successfully.

Testing \\?\Volume{961e043d-faa9-4e15-9ad8-8762cd29e195}\ (Recovery Partition). Please be patient ...

Creating a backup of the existing Recovery Partition. Please be patient ...
Windows RE SHA256 signature: 83FF2CC8D7E17DE27588CF7E373589EE6CA853FBBEF0CB62AB3E93E27A53B46C

Removing active WindowsRE partition...

Computing new System partition size. Please be patient ...

Creating new Windows RE partition ...

Restoring the Windows RE partition ...
The operation completed successfully.
The operation completed successfully.
The operation completed successfully.
The operation completed successfully.

There is a backup of the Recovery partition in C:\Users\ADMINI~1.PCE\AppData\Local\Temp.

Deployment Image Servicing and Management tool
Version: 10.0.22621.2792

Applying image
[==========================100.0%==========================]
The operation completed successfully.
Windows RE SHA256 signature: 83FF2CC8D7E17DE27588CF7E373589EE6CA853FBBEF0CB62AB3E93E27A53B46C
REAGENTC.EXE: Operation Successful.

Removing drive letter R...

Enter 'Yes' to delete the backup of the recovery partition in C:\Users\ADMINI~1.PCE\AppData\Local\Temp , anything else to continue:
Done!

Active WinRE version information

Version: 10.0.22621.2792

ServicePack Build : 3000
ServicePack Level : 0

```
