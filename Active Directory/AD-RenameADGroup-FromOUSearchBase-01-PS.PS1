<#
.SYNOPSIS
  TThis script will rename a group in Active Directory inside an OU. 
.DESCRIPTION
  This script will rename a group in Active Directory inside an OU. The example in here will ranem an ou 
  from something like GG-ADD-DIV-OU1-OU2-OU3-FC to GG-ADD-DIV-OU1_OU2_OU3-FC. Just replace the delimeter
  in the script with your delimeters as well as the number of delimiters.
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.0
  Author:         Julian Thibeault
  Creation Date:  2021/04/06
  Purpose/Change: Initial
.LINK
  https://github.com/StrataNorthCo/PowershellPublic/blob/main/Active%20Directory/AD-RenameADGroup-FromOUSearchBase-01-PS.ps1

  #>

<#Set user variables#>

$reportPath = "TYPE PATH HERE" 
$report = @()
$searchbase = "OU=name,dc=domain,dc=com"

<#You can add as many delimeters as you want and 
adjust the script to do what you need for the name
change. #>

<#Delim1 will be used for the split. This is 
how your group is sectioned off in its name#>
$delim1 = "-"

<#Delim2 will replace the character that is 
currently in delim1 where you want#>
$delim2 = "_"

<#Delim3 will be removed if present in the name#>
$delim3 = "OU"

<#Get the groups from the OU#>
$Groups = Get-ADGroup -SearchBase $searchbase -Filter *

<#Loop through the groups in the array and renambe them#>
foreach ($g in $Groups) {
    $oldGroupName += $g.Name
    $oldGroupNameSplit = $OldGroupName.Split($delim1)
    $newGroupName = $oldGroupNameSplit[0] +
    "$($delim1)" + $oldGroupNameSplit[1] +
    "$($delim1)" + $oldGroupNameSplit[2] +
    "$($delim2)" + $oldGroupNameSplit[3] +
    "$($delim2)" + $oldGroupNameSplit[4] +
    "$($delim1)" + $oldGroupNameSplit[5].replace($delim3, "")
    <#Rename the groups#>
    Set-ADGroup -Identity $g -DisplayName $newGroupName
    Set-ADGroup -Identity $g -SamAccountName $newGroupName
    Rename-ADObject -Identity $g.DistinguishedName -NewName $newGroupName
    
    <#Get the new group name from AD to verify it was changed#>
    $newGroupCheck = Get-ADGroup $newGroupName | Select-Object samaccountname, displayname, distinguishedname
    
    <#Create the report to verify the groups were renamed#>
    $obj = New-Object PSOBJECT
  
    $obj | Add-Member -TYPE NoteProperty -NAME OldDisplayNameName -Value $g.displayname
    $obj | Add-Member -TYPE NoteProperty -NAME NewDisplayName -value $newGroupCheck.displayname
    $obj | Add-Member -TYPE NoteProperty -NAME OldSAMName -value $g.samaccountname
    $obj | Add-Member -TYPE NoteProperty -NAME NewSAMName -value $newGroupCheck.samaccountname
  
    $report += $obj 
  
}
  
<#Remove the comment line to export the 
  report to other methods other than the
  screen#>
  
#$report | export-csv $Reportpath\nameofcsv.csv
#$report | Out-GridHtml