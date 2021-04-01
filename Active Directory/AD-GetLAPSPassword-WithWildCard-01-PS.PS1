$computer = Read-Host "Enter computer name (wildcards are accepted)"
$cred = Get-Credential
Get-ADComputer -Credential $cred -Filter { name -like $computer } -Properties name, ms-mcs-admpwd, ms-mcs-admpwdexpirationtime, canonicalname | Select-Object name, ms-mcs-admpwd, ms-mcs-admpwdexpirationtime, canonicalname
