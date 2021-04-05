<#
.SYNOPSIS
  Remove inheritance from folders in a share. Then, remove the 
  users specified from the ACLs.
.DESCRIPTION
  Remove inheritance from folders in a share. Then, remove the 
  users specified from the ACLs.
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.1
  Author:         Julian Thibeault
  Creation Date:  2021/04/04
  Purpose/Change: Initial
.LINK
  https://github.com/StrataNorthCo/PowershellPublic/blob/main/Systems/SYS-SetACL-RemoveInheritanceandUsers-01-PS.ps1
#>


<#User defined variables#>
$path = "\\Server\folder"
$folders = Get-ChildItem -Path $path -Directory
$Domain = "Dealerships" <#Do not include the top level domain. If your domain is company.com, only enter "company"#>
$admin = (Get-aduser -Filter * | Where-Object { $_.sid -match "500" }).samaccountname
$ACEtoRemove = @("Everyone", "$domain\$admin")
$ACERemovals = @()

foreach ($f in $folders) {
    Write-Host "Getting ACLs for Folder $($path)\$f" -ForegroundColor Green
    $acl = Get-ACL -Path "$($path)\$f" 
    $acl.SetAccessRuleProtection($True, $True)
    Set-Acl -Path "$($path)\$f" -AclObject $acl
    $acl = Get-ACL -Path "$($path)\$f" 
    $ACERemovals += $acl.access | Where-Object { $_.Identityreference -in $ACEtoRemove }
    foreach ($a in $ACERemovals) {
        IF ($ACERemovals) { $acl.RemoveAccessRuleAll($a) }
    }
    Set-Acl -Path "$($Path)\$f" $acl
}