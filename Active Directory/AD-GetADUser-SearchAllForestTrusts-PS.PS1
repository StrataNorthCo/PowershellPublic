<#
.SYNOPSIS
  This script will search for user accounts in all forest trusts and if found, place it into a report.
.DESCRIPTION
  This script will search for user accounts in all forest trusts and if found, place it into a report. 
  Output will be sen to the screen VIA HTML or CSV. 
  Input will be VIA static array or imported from a CSV.
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.0
  Author:         Julian Thibeault
  Creation Date:  2021/04/01
  Purpose/Change: Initial
.LINK
  https://github.com/StrataNorthCo/PowershellPublic/blob/main/Active%20Directory/AD-GetAdUser-SearchAllForestTrusts-PS.ps1

  #>

  <#
    OPTON1
    $userArray = @("user1","user2")
    #>
<#
    OPTION2
    $csv = Import-csv -Path .\EnterPathToCSV

#>

<#User defined variables#>
$report = @()
$dcs = @()

<#Get the forests and domains in the trusts from the current domain#>
$forests = Get-ADTrust -Filter * | Select-Object name
$domain = Get-ADDomain | Select-Object name
$forests += $domain

<#Loop through the forests and get the doman contollers to query#>
foreach ($f in $forests) {

    $dc = (Get-ADDomainController -Discover -DomainName $f.name).hostname

    $dcs += $dc

}

<#Loop through the #>

foreach ($d in $dcs) {

    foreach ($u in $userArray) {

        try {
            Get-ADUser -Identity $u -Server $d -ErrorAction Continue
            write-host "Found $($u) on $($d)" -ForegroundColor green

            <# Prepare the report #>
            $obj = New-Object PSOBJECT

            $obj | Add-Member -TYPE NoteProperty -NAME Name -Value $u
            $obj | Add-Member -TYPE NoteProperty -NAME Domain -Value $d.Substring($d.IndexOf(".") + 1)

            $report += $obj 
        }
        catch {

            write-host "Did not Find $($u) on $($d)" -ForegroundColor yellow

        }

    }

} 

<#HTMLREPORT 
    $report | out-gridhtml
    #>
<#CSVReport
    $report | export-csv -path .\PlaceExportPathHere 
    #>
