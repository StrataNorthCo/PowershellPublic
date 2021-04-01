<#
.SYNOPSIS
  This script will gather the LAPS password of an AD Computer object. 
.DESCRIPTION
  This script will gather the LAPS password of an AD Computer object. The input will also access wildcard characters.
  The output will be generated on the screen.
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES

  Title:            AD-GetLPASPassword-WithWildCard-PS
  Version:          1.0
  Author:           Julian Thibeault
  Creation Date:    2021/04/01
  Purpose/Change:   Initial
  Website:          https://StrataNorth.co
  GitHub:           https://github.com/StrataNorthCo/PowershellPublic
.LINK
  https://github.com/StrataNorthCo/PowershellPublic/blob/main/Active%20Directory/AD-GetLPASPassword-WithWildCard-PS.ps1
#>


<#User defined variables#>
$computer = Read-Host "Enter computer name (wildcards are accepted)"
$cred = Get-Credential

Get-ADComputer -Credential $cred -Filter { name -like $computer } -Properties name, ms-mcs-admpwd, ms-mcs-admpwdexpirationtime, canonicalname | Select-Object name, ms-mcs-admpwd, ms-mcs-admpwdexpirationtime, canonicalname
