<#
Title: AD-GetAdUser-RemoveAdminCount1-PS
Summary: This script will remove the admincount1 attribute from users that 
            are no longer in protected groups and no longer require it.
Version: 1.0
Date: 2021-03-21
Author: Julian Thibeault
Website: https://StrataNorth.co
GitHub: https://github.com/StrataNorthCo/PowershellPublic
#>

<#
.SYNOPSIS
  Remove the admincount1 attribute from users that 
  are no longer in protected groups and no longer require it. 
.DESCRIPTION
  This script will remove the admincount1 attribute from users that 
  are no longer in protected groups and no longer require it. 
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.0
  Author:         Julian Thibeault
  Creation Date:  2021/03/21
  Purpose/Change: Initial
.LINK
  https://github.com/StrataNorthCo/PowershellPublic/blob/main/Active%20Directory/AD-GetAdUser-RemoveAdminCount1-PS.ps1
#>


<#User defined variables#>
$removeUsers = @()
$report = @()
$object = @()
$SMs = @()
$reportPath = ".\AD-GetAdUser-RemoveAdminCount1-CSV.csv"
$logpath = ".\AD-GetAdUser-RemoveAdminCount1-LOG-$($date).txt"
$date = Get-Date -Format "yyyy_MM_dd_hh_mm_ss"

<#Start logging#>
Start-Transcript -Path $logpath

<#Get all of the domains that are trusted including the domain this is run in
IF YOU ONLY HAVE ONE DOMAIN YOU CAN COMMENT THIS SECTION#>
$Forests = (Get-ADTrust -filter * | Select-Object Target).target
$adDomain = (Get-ADDomain | Select-Object forest).forest

$allDomains = $forests + $adDomain

<#Cycle through the domains and get the schema master
IF YOU ONLY HAVE ONE DOMAIN YOU CAN COMMENT THIS SECTION#>
foreach ($d in $allDomains) {

  $SMs += (Get-ADForest -Identity $d | Select-Object schemaMaster).schemamaster

}

<#Cycle through each schema master to get the admincount for the users 
for each domain#>

<#IF YOU ONLY HAVE ONE DOMAIN YOU CAN REMOVE THE NEXT LINE AND THE 
LAST BRACE#>
foreach ($s in $SMs) {

  <#Gather the domain and forest SIDS from well-known accounts#>
  $rootDomain = (Get-ADForest -Server $s | Select-Object rootdomain).rootdomain
  $rootDomainSID = (Get-ADDomain $rootdomain | Select-Object-Object domainsid).domainsid.value
  $domainSID = (Get-ADDomain -Server $s | Select-Object domainsid).domainsid.value

  <#Set the variables for the well-known SIDS for the domain.#>
  $adm = "$($domainSID)-500"
  $da = "$($domainSID)-512"
  $dc = "$($domainSID)-516"
  $cp = "$($domainSID)-517"
  $sa = "$($rootDomainSID)-518"
  $ea = "$($rootDomainSID)-519"
  $rodc = "$($domainSID)-521"
  $ka = "$($domainSID)-526"
  $eka = "$($rootDomainSID)-527"
  $kr = "$($domainSID)-502"

  <#Create the array for all the well-known SIDs in the forest#>
  $knownSIDS = 'SID -eq "S-1-5-32-544" -or 
SID -eq "S-1-5-32-548" -or
SID -eq "S-1-5-32-549" -or 
SID -eq "S-1-5-32-550" -or 
SID -eq "S-1-5-32-551" -or 
SID -eq "S-1-5-32-552" -or
SID -eq $da -or
SID -eq $dc -or
SID -eq $cp -or
SID -eq $rodc -or
SID -eq $ka -or
SID -eq $sa -or
SID -eq $ea -or
SID -eq $eka'

  <#Get your protected groups based on the SID filters above#>
  $protectedGroups = Get-ADGroup -Server $s -Filter $knownSIDS

  <#Get all of the users in the forest that have an admin count of "1". Filter out 
    the well known accounts (buildtin\admin and KRBTGT) that we don't want to clear.#>
  $adminCountUsers = Get-ADuser -Server $s -filter 'admincount -eq "1" -and (SID -ne $adm -and SID -ne $kr)' -Properties memberof, CanonicalName | Select-Object name, distinguishedname, memberof, CanonicalName

  <#Now loop through the users with admincount of "1" and set a value to "0"
    Then, loop through their individual membership and compare that to the
    protected groups. If the user has a membership that is supposed to be 
    in the protected group, it will add a "1" to the variable. If not, the 
    variable will be a "0" and the user will be added to an array of users
    that no longer need the admincount set to "1". remove the comment below 
    to clear the admincount#>

  foreach ($acu in $adminCountUsers) {
    $admincount = 0
    $acumember = $acu.memberof
    Foreach ($item in $acumember) {
      IF ($protectedGroups.distinguishedname -contains $item) {
        Write-Host "$($acu.name) is a member of $($item) which is a protected group" -ForegroundColor Magenta
        $admincount ++ 
      }
      Else {
        Write-host "$($acu.name) is not a member of a protected group" -ForegroundColor Yellow
      }
    }
    IF ($admincount -eq "0") {
      $domain = $acu.CanonicalName -replace '\/.*$'
      $RemoveUsers += $acu.name
      #Set-ADUser $acu.distinguishedname -clear admincount


      # CHECK THE NEW LOCATION AND REPORT 

      $object = New-Object -TypeName psobject

      $object | Add-Member -MemberType NoteProperty -Name accountName -Value $acu.name
      $object | Add-Member -MemberType NoteProperty -Name DN -Value $acu.distinguishedname
      $object | Add-Member -MemberType NoteProperty -Name Domain -Value $Domain
        
      $report += $object
    }
  }   
}

<#Generate an HTML report as well as a CSV report#>
$report | Out-GridHtml
$report | Export-Csv $reportPath

<#Stop logging#>
Stop-Transcript