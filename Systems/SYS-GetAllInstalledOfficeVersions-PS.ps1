$computerList = @()
$officeVersions = @()

$Root = [ADSI]"LDAP://RootDSE"
$domainRoot = $Root.rootDomainNamingContext

$Filter = "(&(objectCategory=computer)(!userAccountControl:1.2.840.113556.1.4.803:=2))"
$searchOU = "ou=computers,ou=weirs gmc"
$searchbase = "$($searchOU),$($domainRoot)"

$Searcher = New-Object DirectoryServices.DirectorySearcher
$Searcher.SearchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($searchbase)")
$Searcher.Filter = $Filter
$enabledComputers = ($Searcher.FindAll()).properties

foreach ($i in $enabledComputers) {

    $object = New-Object -TypeName psobject

    $object | Add-Member -MemberType NoteProperty -Name name -Value $i.name
    $object | Add-Member -MemberType NoteProperty -Name dnshostname -Value $i.dnshostname
    $object | Add-Member -MemberType NoteProperty -Name distinguishedname -Value $i.distinguishedname
        
    $computerList += $object

}


$computerList

$computerListTotal = $computerList.count
$count = 1

Foreach ($c in $computerlist) {
    Write-Progress -Activity "Gathering Software List" -Status "Imported $count of $($computerListTotal)" -PercentComplete (($count / $computerListTotal) * 100)
    Write-Host "$($count) of $($computerListTotal) computers $($c.name) is now processing" -ForegroundColor Magenta
    $count ++
    IF ((Test-Connection -ComputerName $c.dnshostname -count 1 -quiet) -eq "True") {
        Write-host "$($c.name) was reachable" -ForegroundColor Green
        $officeSearch = Get-WmiObject win32_product -ComputerName $c.name | Where-Object { ($_.Name -like "*click-to-run licensing*") -or 
            (($_.Name -like "*office*") -and
                ($_.Name -like "*starter*") -or
                ($_.Name -like "*online*") -or
                ($_.Name -like "*basic*") -or
                ($_.Name -like "*personal*") -or
                ($_.Name -like "*university*") -or
                ($_.Name -like "*proplus*") -or 
                ($_.Name -like "*rt*") -or
                ($_.Name -like "*student*") -or
                ($_.Name -like "*business*") -or
                ($_.Name -like "*ultimate*") -or
                ($_.Name -like "*standard*") -or
                ($_.Name -like "*enterprise*")) } | Select-Object Name, Version
        Write-Host "Found $($officeSearch.count) office entries on $($c.name)" -ForegroundColor Cyan

        Foreach ($o in $officesearch) {

            $object = New-Object -TypeName psobject
      
            $object | Add-Member -MemberType NoteProperty -Name Computer -Value $c.name
            $object | Add-Member -MemberType NoteProperty -Name ProductName -Value $o.name
            $object | Add-Member -MemberType NoteProperty -Name Version -Value $o.version
            $object | Add-Member -MemberType NoteProperty -Name Caption -Value $o.caption
        
            $officeVersions += $object

        }
    }
    Else { Write-host "$($c.name) was not reachable" -ForegroundColor Yellow }

}$officeVersions

