<#
.SYNOPSIS
Get-MSOfficeFiles gets the number of MS Office Files saved in a local or remote computer(s).

.DESCRIPTION
Get-MSOfficeFiles retrieves the number of .doc, .docx, .xls, .xlsx, .ppt, and .pptx. files that are saved in the Desktop, Documents, and Downloads folders of logged on user(s) for the past year.

.EXAMPLE
Invoke-Command -ScriptBlock $scriptBlock 

.NOTES
In addition, you could save the script as a powershell script and then run Invoke-Command -ComputerName localhost -ScriptBlock { C:\PowerShellScriptsFolder\Get-MSOfficeFile.ps1} -Credential DomainName\Credential
instead of running the entire script in a file. When saving the script, copy from $searchResultofDocFile = o... to ...Write-Output $searchResultofPptxFile' .pptx files: ' >> $destFolder\$folderName.txt. Name it Get-MSOfficeFile.ps1
 
Author:  Omar Rosa
 
#>


$scriptBlock  = {

$searchResultofDocFile = 0
$searchResultofDocxFile = 0
$searchResultofXlsFile = 0 
$searchResultofXlsxFile = 0
$searchResultofPptFile = 0
$searchResultofPptxFile = 0

$usersFolder = Get-ChildItem -Name -Directory -Path C:\Users\ -Exclude administrator, public, Default.migrated, Default # -Exclude excludes folder(s) that you do not want the script to search
$FolderLocation3D = @("Documents", "Desktop", "Downloads") 

$destFolder = 'C:\destFol'

IF(!(Test-Path $destFolder))
{
    New-Item -Path $destFolder -Type Directory

}


foreach ($folderName in $usersFolder)
{
    foreach ($FolderLocationsD in $FolderLocation3D)
    {    
       $searchResultofDocFile += (Get-ChildItem -Path C:\Users\$folderName\$FolderLocationsD -Recurse -Include "*.doc" -ea 0 | Where {$_.LastWriteTime -gt (Get-Date).AddDays(-365)}).count
       $searchResultofDocxFile += (Get-ChildItem -Path C:\Users\$folderName\$FolderLocationsD -Recurse -Include "*.docx" -ea 0 | Where {$_.LastWriteTime -gt (Get-Date).AddDays(-365)}).count 
       $searchResultofXlsFile += (Get-ChildItem -Path C:\Users\$folderName\$FolderLocationsD -Recurse -Include "*.xls" -ea 0 | Where {$_.LastWriteTime -gt (Get-Date).AddDays(-365)}).count
       $searchResultofXlsxFile += (Get-ChildItem -Path C:\Users\$folderName\$FolderLocationsD -Recurse -Include "*.xlsx" -ea 0 | Where {$_.LastWriteTime -gt (Get-Date).AddDays(-365)}).count
       $searchResultofPptFile += (Get-ChildItem -Path C:\Users\$folderName\$FolderLocationsD -Recurse -Include "*.ppt" -ea 0 | Where {$_.LastWriteTime -gt (Get-Date).AddDays(-365)}).count
       $searchResultofPptxFile += (Get-ChildItem -Path C:\Users\$folderName\$FolderLocationsD -Recurse -Include "*.pptx" -ea 0 | Where {$_.LastWriteTime -gt (Get-Date).AddDays(-365)}).count
       
    }
}


  Write-Output $searchResultofDocFile' .doc files: ' >> $destFolder\$folderName.txt
  Write-Output $searchResultofDocxFile' .docx files: ' >> $destFolder\$folderName.txt
  Write-Output $searchResultofXlsFile' .xls files: ' >> $destFolder\$folderName.txt
  Write-Output $searchResultofXlsxFile' .xlsx files: ' >> $destFolder\$folderName.txt
  Write-Output $searchResultofPptFile' .ppt files: ' >> $destFolder\$folderName.txt
  Write-Output $searchResultofPptxFile' .pptx files: ' >> $destFolder\$folderName.txt


  }


 
 # In order for this command to work in a remote computer, PSRemoting should be enable.
 # Get-Service WinRM verifies whether the WinRM service is running.
 # Enable-PSRemoting enables a computer to receive remote command. 
 # Refer to https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/enable-psremoting?view=powershell-5.1 for a description of Enable-PSRemoting
 
 $remoteComputers = 'localhost'
 Invoke-Command -ComputerName $remoteComputers -ScriptBlock $scriptBlock -Credential DomainName\Credential
  
 # to run the command in the local computer
 Invoke-Command -ScriptBlock $scriptBlock