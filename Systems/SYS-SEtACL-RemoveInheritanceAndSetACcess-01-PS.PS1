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
  Version:        1.0
  Author:         Julian Thibeault
  Creation Date:  2021/04/04
  Purpose/Change: Initial
.LINK
  https://github.com/StrataNorthCo/PowershellPublic/blob/main/Systems/SYS-SetACL-RemoveInheritanceAndSetAccess-01-PS.ps1
#>


<#User defined variables#>
$path = "\\WEIRS-FILESERVE\TEST"
$folders = Get-ChildItem -Path $path -Directory
$ACEtoKeep = @("NT AUTHORITY\SYSTEM", "BUILTIN\Administrators")

foreach ($f in $folders) {
    Write-Host "Getting ACLs for Folder $($path)\$f" -ForegroundColor Green
    $acls = Get-ACL -Path "$($path)\$f"
    
    <#check inheritance and remove if inheriting#>
    $folderInherit = $acls.GetAccessRules($false, $True, [System.Security.Principal.NTAccount])
    IF ($folderInherit.count -ge 1) {
        Write-Host "$($path)\$($f) is inheriting"
        $acls.SetAccessRuleProtection($True, $True)
        Set-Acl -Path "$($path)\$f" -AclObject $acls
        $acls = Get-ACL -Path "$($path)\$f"
        Clear-Variable folderInherit
    }
    
    foreach ($acl in $acls.access) {
        IF (($acl.identityreference.value -notlike "*$($f.name)*") -and 
            ($ACEtoKeep -notcontains $acl.identityreference.value)) {
            Write-Host "$($acl.IdentityReference.value) will be removed" -ForegroundColor Green
            $acls.RemoveAccessRuleAll($acl)
            Set-Acl -Path "$($Path)\$f" $acls
        }
        else { write-host "$($acl.IdentityReference.value) will stay put" -ForegroundColor Gray }
           
        IF (($acl.accesscontroltype -eq "Allow") -and
            ($acl.identityreference.value -like "*$($f.name)*") -and
            ($acl.isinherited -eq $False) -and
            ($acl.inheritanceflags -eq "containerinherit,ObjectInherit") -and 
            ($acl.PropagationFlags -eq "InheritOnly")) {
            Write-host "$($f.name) already has correct permissions to the folder" -ForegroundColor Magenta
        }
        else {
            Write-host "$($f.name) does not have correct permssions to the folder $($path)\$f" -ForegroundColor Cyan
            $user = (Get-ADUser $f.name).UserPrincipalName
            Set-Acl -Path "$($path)\$f" -AclObject $acls
            $acls = Get-ACL -Path "$($path)\$f" 
            $permission = "$user", "FullControl", 'ObjectInherit,ContainerInherit', "None", "allow"
            $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $permission
            $acls.SetAccessRule($accessRule)
            $acls | Set-Acl -Path "$($Path)\$f"

        }
       
    }
}