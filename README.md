# Get-MSOfficeFiles

.SYNOPSIS
Get-MSOfficeFiles gets the number of MS Office Files saved in a local or remote computer(s).

.DESCRIPTION
Get-MSOfficeFiles retrieves the number of .doc, .docx, .xls, .xlsx, .ppt, and .pptx. files that are saved in the Desktop, Documents, and Downloads folders of logged on user(s) for the past year.

.EXAMPLE
Invoke-Command -ScriptBlock $scriptBlock 

.NOTES
In addition, you could save the script as a powershell script and then run Invoke-Command -ComputerName localhost -ScriptBlock { C:\PowerShellScriptsFolder\Get-MSOfficeFile.ps1} -Credential DomainName\Credential instead of running the entire script in a file.
When saving the script, copy from $searchResultofDocFile = o... to ...Write-Output $searchResultofPptxFile' .pptx files: ' >> $destFolder\$folderName.txt. Name it Get-MSOfficeFile.ps1
 
Author:  Omar Rosa
 
