<#
    .SYNOPSIS
    Import one or more pst files into an exisiting mailbox or a archive
   
   	Thomas Stensitzki
	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	Version 1.2, 2016-06-07

    .DESCRIPTION
	
    This script imports one or more PST files into a user mailbox or a user archive as batch.
    PST file names can used as target folder names for import. PST files are renamed to support
    file name limitations by New-MailboxImportRequest cmdlet.

    All files of a given folder will be imported into the user's mailbox.

    .NOTES 
    Requirements 
    - Windows Server 2012 R2  
    - Exchange Server 2013
    - GlobalFunctions PowerShell Module, https://www.powershellgallery.com/packages/GlobalFunctions

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial release
    1.1     log will now be stored in a subfolder (name equals Identity) 
    1.2     PST file renaming added
	
	.PARAMETER Identity  
    Type: string. Mailbox identity in which the pst files get imported

    .PARAMETER Archive
    Type: switch. Import pst files into the online archive.

    .PARAMETER FilePath
    Type:string. Folder which contains the pst files. Has to be an UNC path.

    .PARAMETER FilenameAsTargetFolder
    Type: switch. Import the PST files into dedicated target folders. The folder name will equal the file name.

    .PARAMETER BadItemLimit
    Type: int32. Standard is set to 0. Don't max it out because the script doesn't add "AcceptLargeDatalost".

    .PARAMETER ContinueOnError
    Type: switch. If set the script continue with the next pst file if a import request failed.

    .PARAMETER ContinueOnError
    Type: int32. Timespan to wait between import request staus checks in seconds. Default: 320

    .EXAMPLE
    Import all PST files into the mailbox "testuser"
    .\Start-MailboxImport.ps1 -Identity testuser -Filepath "\\testserver\share"

    .EXAMPLE
    Import all PST files into the mailbox "testuser". Use PST file name as target folder name. Wait 90 seconds between each status check.
    .\Start-MailboxImport.ps1 -Identity testuser -Filepath "\\testserver\share\*" -FilenameAsTargetFolder -SecondsToWait 90

    #>

Param(
    [parameter(Mandatory=$true,ValueFromPipeline=$false)]
        [string]$Identity,
    [parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [switch]$Archive,
    [parameter(Mandatory=$true,ValueFromPipeline=$false)]
        [string]$FilePath,
    [parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [switch]$FilenameAsTargetFolder,
    [parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [int32]$BadItemLimit = 0,
    [parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [switch]$ContinueOnError,
    [parameter(Mandatory=$false,ValueFromPipeline=$false)]
        [int32]$SecondsToWait = 320
)

    Import-Module ActiveDirectory

    Set-StrictMode -Version Latest

    # IMPORT GLOBAL MODULE
    Import-Module BDRFunctions
    $ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
    $ScriptName = $MyInvocation.MyCommand.Name
    # Create a log folder for each identity
    $logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 14 -LogFolder $Identity
    $logger.Write('Script started')
    $InfoScriptFinished = 'Script finished.'

    <#
        Helpder function to remove invalid chars for New-mailboxImportRequest cmdlet
    #>
    Function Optimize-PstFileName {
    param (
        [string]$PstFilePath
    )
    
        foreach ($pst in (Get-ChildItem -Path $PstFilePath -Include '*.pst')) {
            $newFileName = $pst.Name
            # List of chars, add additional chars as needed
            $chars = @(' ','(',')')
            $chars | % {$newFileName = $newFileName.replace($_,'')} 
            $logger.Write("Renaming PST: Old: $($pst.Name) New: $($newFileName)")
            if($newFileName -ne $pst.Name) {
                # Rename PST file
                $pst | Rename-Item -NewName $newFileName
            }
        }
    }

    # Get all pst files from the share
    if ($FilePath.StartsWith('\\')) { 
        try {
            # Check file path and add wildcard, if required
            If ((!$FilePath.EndsWith('*')) -and (!$FilePath.EndsWith('\'))) {
                $FilePath = $FilePath + '\*'
            } 

            Optimize-PstFileName -PstFilePath $FilePath

            # Fetch all pst files in source folder
            $PstFiles = Get-ChildItem -Path $FilePath -Include '*.pst'

            # Check if there are any files to import
            If (($PstFiles| Measure-Object).Count) {
                $InfoMessage = "Note: Script will wait $($SecondsToWait)s between each status check!"
                Write-Host $InfoMessage
                $logger.Write($InfoMessage)

                # Fetch AD user object from Active Directory
                $Name = Get-ADUser $Identity

                foreach ($PSTFile in $PSTFiles) {
                    $ImportName = $($Name.SamAccountName + '-' + $PstFile.Name)
				    $InfoMessage = "Create New-MailboxImportRequest for user: $($Name.Name) and file: $($PSTFile)"

                    # Built command string
                    $cmd = "New-MailboxImportRequest -Mailbox $($($Name).SamAccountName) -Name $($ImportName) -FilePath ""$($PSTFile)"" -BadItemLimit $($BadItemLimit) -WarningAction SilentlyContinue"
                    if ($Archive) {
                        $cmd = $cmd + ' -IsArchive'
					    $InfoMessage = "$($InfoMessage) into the archive."
                    }
				    else {
					    $InfoMessage = $InfoMessage + '.'
				    }
                    if ($FilenameAsTargetFolder) {
                        [string]$FolderName = $($PSTFile.Name.ToString()).Replace('.pst', '')
                        $cmd = $cmd + " -TargetRootFolder ""$($FolderName)"""
					    $InfoMessage = $InfoMessage + " Targetfolder:""$($FolderName)""."
                    }

				    Write-Host $InfoMessage
                    $logger.Write($InfoMessage)

                    try {
                       Invoke-Expression -Command $cmd | Out-Null
                    } 
                    catch {
                        $ErrorMessage = "Error accessing creating import request for user $($Name.Name). Script aborted."
                        Write-Error $ErrorMessage
                        $logger.Write($ErrorMessage,1)
                        Exit(1)
                    }
                    # Some nice sleep
                    Start-Sleep -Seconds 5

                    [bool]$NotFinished = $true
                    $logger.Write("Waiting for import request $($ImportName) to be completed.")
                    while($NotFinished) {
                       try {
					       $ImportRequest = Get-MailboxImportRequest -Mailbox $($($Name).SamAccountName) -Name $($ImportName) -ErrorAction SilentlyContinue
					       switch ($ImportRequest.Status) {
							    'Completed' {
								    # Remove the ImportRequest so we can't run into the limit
                                    $InfoMessage = "Import request $($ImportName) completed successfully."
								    Write-Host $InfoMessage
								    $logger.Write("$($InfoMessage) Import Request Statistics Report:")

                                    # Fetch Import statistics
                                    $importRequestStatisticsReport = (Get-MailboxImportRequest -Mailbox $($($Name).SamAccountName) -Name $($ImportName) | Get-MailboxImportRequestStatistics -IncludeReport).Report
                                    $logger.Write($importRequestStatisticsReport)

                                    # Delete mailbox import request
								    Get-MailboxImportRequest -Mailbox $($($Name).SamAccountName) -Name $($ImportName) | Remove-MailboxImportRequest -Confirm:$false

                                    $InfoMessage = "Import request $($ImportName) deleted."
								    Write-Host $InfoMessage
								    $logger.Write($InfoMessage)

								    $NotFinished = $false 
							    }
							    'Failed' {
                                    $InfoMessage = "Error: Administrative action is needed. ImportRequest $($ImportName) failed."
								    Write-Error $InfoMessage
								    $logger.Write($InfoMessage,1)
								    if (-not $ContinueOnError) {
									    Write-Host $InfoScriptFinished
									    $logger.Write($InfoScriptFinished)
									    Exit(2)
								    } 
                                    else {
                                        $InfoMessage = 'Info: ContinueonError is set. Continue with next PST file.'
									    Write-Host $InfoMessage
									    $logger.Write($InfoMessage)
									    $NotFinished = $false
								    }
							    }
							    'FailedOther' {
								    Write-Error "Error: Administrative action is needed. ImportRequest $($ImportName) failed."
								    $logger.Write("Error: Administrative action is needed. ImportRequest $($ImportName) failed.",1)
								    if (-not $ContinueOnError) {
									    Write-Host $InfoScriptFinished
									    $logger.Write($InfoScriptFinished)
									    Exit(2)
								    } else {
                                        $InfoMessage = 'Info: ContinueonError is set. Continue with next pst file.'
									    Write-Host $InfoMessage
									    $logger.Write($InfoMessage)
									    $NotFinished = $false
								    }
							    }
							    default { 
								    Write-Host "Waiting for import $($ImportName) to be completed. Status: $($ImportRequest.Status)"
								    Start-Sleep -Seconds $SecondsToWait
							    }
					       }
					    } 
                        catch {
                            $InfoMessage = "Error on getting Mailboximport statistics. Trying again in $($SecondsToWait) seconds."
						    Write-Host $InfoMessage
						    $logger.Write($InfoMessage, 1)
                            Start-Sleep -Seconds $SecondsToWait
					    }
                    }
                }
            } Else {
                $InfoMessage = "No files for import found in $($FilePath)."
                Write-Host $InfoMessage
                $logger.Write($InfoMessage)
            }


        } 
        catch {
            $InfoMessage = 'Error accessing $($FilePath). Script aborted.'
            Write-Error $InfoMessage
            $logger.Write($InfoMessage, 1)
        }
    } 
    Else {
        $InfoMessage = 'Filepath has to be an UNC path. Script aborted.'
        Write-Error $InfoMessage
        $logger.Write($InfoMessage, 1)
}

Write-Host $InfoScriptFinished
$logger.Write($InfoScriptFinished)