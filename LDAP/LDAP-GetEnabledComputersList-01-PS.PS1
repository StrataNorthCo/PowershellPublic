<#
Title: LDAP-GetEnabledComputersList-01-PS
Summary: This script will use LDAP to find the enabled computers in an OU
            inside of ActiveDirectoy.
Version: 1.0
Date: 2021-03-21
Author: Julian Thibeault
Website: https://StrataNorth.co
GitHub: https://github.com/StrataNorthCo/PowershellPublic
#>

<#
.SYNOPSIS
  Use LDAP to find the enabled computers in an OU
        inside of ActiveDirectoy. This is good to 
        use if you do not have a computer to install
        or run the RSAT tools for Get-AD cmdlets.
.DESCRIPTION
  This script uses LDAP bind to identify the computers
        in AD that are enabled.
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.0
  Author:         Julian Thibeault
  Creation Date:  2021/03/28
  Purpose/Change: Initial
.LINK
  https://github.com/StrataNorthCo/PowershellPublic/blob/main/Active%20Directory/AD-GetAdUser-RemoveAdminCount1-PS.ps1
#>

<#User defined variables#>
$computerList = @()
$Root = [ADSI]"LDAP://RootDSE"
$domainRoot = $Root.rootDomainNamingContext
$Filter = "(&(objectCategory=computer)(!userAccountControl:1.2.840.113556.1.4.803:=2))"
$searchOU = "ou=computers,ou=weirs gmc"
$searchbase = "$($searchOU),$($domainRoot)"
$Searcher = New-Object DirectoryServices.DirectorySearcher
$Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($searchbase)")

<#Perform the search using the filter#>
$Searcher.Filter = $Filter
<#Gather the properties of the returned array results#>
$enabledComputers = ($Searcher.FindAll()).properties

<#Loop through the array and convert the properties to objects.
    Display the results in a new array.#>
foreach ($i in $enabledComputers) {

    $object = New-Object -TypeName psobject

    $object | Add-Member -MemberType NoteProperty -Name name -Value $i.name
    $object | Add-Member -MemberType NoteProperty -Name dnshostname -Value $i.dnshostname
    $object | Add-Member -MemberType NoteProperty -Name distinguishedname -Value $i.distinguishedname
        
    $computerList += $object

}

<#Final array of computers#>
$computerList