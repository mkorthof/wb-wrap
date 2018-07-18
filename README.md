# wb.cmd
## Wrapper for Windows Backup and Restore

This wrapper is meant to work with "Windows Backup and Restore".

In Windows 10 this is available as: `"Control Panel\All Control Panel Items\Backup and Restore (Windows 7)"`

This PowerShell script adds the following functionlity:

- Allows creating multiple system images using dated dirs
- Deletes oldest backup folder(s) until there is enough space
- Runs normal backup (using sdclt.exe)
- Logging to "wb.log", located in same dir as script

First some checks are done to make sure you have Admin rights, the Backup Drive is available etc. Then if a Backup folder exists like `"Z:\WindowsImageBackup\COMPUTERNAME"` it will be renamed to `"COMPUTERNAME_yyMMdd"`.
If the new Backup is e.g. 200GB and your Backup Drive only has 100GB free, older Backup folders will be removed until there is enough free space (so 200GB at least).
Now a normal Windows Backup with progress window will be started using the settings you configured using the GUI (via Control Panel). In case anything goes wrong the script will let you know and abort. Any messages go to a separate console window and to `wb.log`.

### Usage

Just start `wb.cmd`

For easy execution this PowerShell script is embedded in a Batch .cmd file using a "polyglot wrapper".
It also can be renamed to wb.ps1. More info:
- https://blogs.msdn.microsoft.com/jaybaz_ms/2007/04/26/powershell-polyglot
- https://stackoverflow.com/questions/29645

### Configuraton

All config is done by editting `wb.cmd`.

Either set `$backupDrive` to a static drive letter (e.g. "D") or a UUID: "\\?\Volume{c00d6b5a-f734-48e6-b321-029977e5169f".
To list UUIDS you can use this CLI command: `mountvol`.

`$backupDrive = "\\?\Volume{c00d6b5a-f734-48e6-b321-029977e5169f}\"`

or:

`$backupDrive = "D"`

Set Volume Label of Backup Drive:

`$backupLabel = "My Backups"`

Is the Backup Drive encrypted with BitLocker [0/1]

`$blDrive = 1`

Disk space options: % free space needed, bytes, max number of folders to remove.

`$minPercentFree = 12.5`

`$minBytesFree = 375809638400` (= 375GB)

`$maxRemove = 3`

Both minPercentFree and minBytesFree (actually "OR") are used to check against current free disk space, as long as either one checks out no folders will be removed. You could look at your current backup folder to calculate these.
