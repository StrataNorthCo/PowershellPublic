<#
.SYNOPSIS
  This script will import users from a CSV and then created them in the location you specify.
.DESCRIPTION
  This script will import users from a CSV and then created them in the location you specify.
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.0
  Author:         Julian Thibeault
  Creation Date:  2021/04/05
  Purpose/Change: Initial
.LINK
  https://github.com/StrataNorthCo/PowershellPublic/blob/main/Active%20Directory/AD-NewADUser-FromCSV-01-PS.ps1

  #>

<#User defined variables#>
$searchbase = "OU=People,DC=Domain,DC=Com"
$report = @()
$CSVPath = 'C:\temp\AD-NewUser-FromCSV-01-IMPORT.csv'
$ExportCSVPath = 'C:\temp\AD-NewUser-FromCSV-01-EXPORT.csv'

<#Import the users from the CSV#>
$usersArray = Import-csv -Path $CSVPath

<#Loop through the users array to create the user accounts#>
foreach ($u in $usersArray) {
  
    $gad = Get-ADUser -Filter { displayname -eq '$u.displayname' } -Properties displayname, canonicalname | 
    Select-Object displayname, canonicalname 
                  
    
    IF ($gad) {
        write-host "$($u.name) already created in $($gad.canonicalname), adding to report" -ForegroundColor Yellow
              
        $created = "no"
        <# Prepare the report #>
        $obj = New-Object PSOBJECT
  
        $obj | Add-Member -TYPE NoteProperty -NAME Name -Value $gad.displayname
        $obj | Add-Member -TYPE NoteProperty -NAME Location -Value $gad.canonicalname
        $obj | Add-Member -TYPE NoteProperty -NAME Created -Value $created
  
        $report += $obj
        Clear-Variable -Name created
    }
    ELSE {
        write-host "$($u.displayname) was not found, good to proceed." -ForegroundColor green
  
        New-ADUser -Name $u.name -GivenName $u.firstname -Surname $u.lastname -SamAccountName $u.samaccountname `
            -DisplayName $u.displayname -UserPrincipalName $u.upn -Description $u.description `
            -Path $searchbase -Enabled:$true -AccountPassword (ConvertTo-SecureString -AsPlainText "WelcomeToZoll1!" -Force) -passThru
              
        $created = "Yes"
  
        <# Prepare the report #>
        $obj = New-Object PSOBJECT
  
        $obj | Add-Member -TYPE NoteProperty -NAME Name -Value $u.displayname
        $obj | Add-Member -TYPE NoteProperty -NAME Location -value $searchbase
        $obj | Add-Member -TYPE NoteProperty -NAME Created -value $created
  
        $report += $obj 
  
        Clear-Variable -Name created
  
    }
}
   <#Display the report on the screen#>
   $report
  
<#HTMLREPORT#>
      #$report | out-gridhtml

<#CSVReport#>
      $report | export-csv -path $ExportCSVPath 
  
  
  
   