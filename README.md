# Start-MailboxImport.ps1
Import one or more PST files into an exisiting mailbox or a archive.

##Description
This script imports one or more PST files into a user mailbox or a user archive as batch.

PST file names can used as target folder names for import. PST files are renamed to support file name limitations by New-MailboxImportRequest cmdlet.

All files of a given folder will be imported into the user's mailbox.

## Notes

Requirements 
- Windows Server 2012 R2  
- Exchange Server 2013
- GlobalFunctions PowerShell Module, https://www.powershellgallery.com/packages/GlobalFunctions

##Inputs

Identity  
Type: string. Mailbox identity in which the PST files get imported

Archive
Type: switch. Import PST files into the online archive.

FilePath
Type:string. Folder which contains the PST files. Has to be an UNC path.

FilenameAsTargetFolder
Type: switch. Import the PST files into dedicated target folders. The folder name will equal the file name.

BadItemLimit
Type: int32. Standard is set to 0. Don't max it out because the script doesn't add "AcceptLargeDatalost".

ContinueOnError
Type: switch. If set the script continue with the next PST file if a import request failed.

ContinueOnError
Type: int32. Timespan to wait between import request staus checks in seconds. Default: 320

##Outputs
None

##Examples
```
.\Start-MailboxImport.ps1 -Identity testuser -Filepath "\\testserver\share"
```
Import all PST files into the mailbox "testuser"

```
.\Start-MailboxImport.ps1 -Identity testuser -Filepath "\\testserver\share\*" -FilenameAsTargetFolder -SecondsToWait 90
```
Import all PST files into the mailbox "testuser". Use PST file name as target folder name. Wait 90 seconds between each status check.


##TechNet Gallery
Find the script at TechNet Gallery
* https://gallery.technet.microsoft.com/Purge-Exchange-Server-2013-c2e03e72


##Credits
Written by: Thomas Stensitzki

Follow me:

* My Blog: https://www.granikos.eu/en/justcantgetenough
* Archived Blog: http://www.sf-tools.net/
* Twitter:	https://twitter.com/stensitzki
* LinkedIn:	http://de.linkedin.com/in/thomasstensitzki
* Github:	https://github.com/Apoc70

For more Office 365, Cloud Security and Exchange Server stuff checkout services provided by Granikos

* Blog:     http://blog.granikos.eu/
* Website:	https://www.granikos.eu/en/
* Twitter:	https://twitter.com/granikos_de