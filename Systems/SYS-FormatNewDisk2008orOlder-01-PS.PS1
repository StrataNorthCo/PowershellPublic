<#
.SYNOPSIS
  Format a newly added disk on any Windows Server OS 2008 or older where Get-Disk module is not supported
.DESCRIPTION
  This script utilizes DISKPART and regEx to parse string data to gather the newly added raw disk and format it.
  
  Disclaimer: This script is offered "as-is" with no warranty. 
  While the script is tested and working in my environment, it is recommended that you test the script
  in a test environment before using in your production environment.
 
.NOTES
  Version:        1.0
  Author:         Julian Thibeault and Dempsey Dunkin
  Creation Date:  2021/04/14
  Purpose/Change: Initial
.LINK
  https://github.com/StrataNorthCo/PowershellPublic/blob/main/Systems/SSYS-FormatNewDisk2008orOlder-01-PS.ps1
#>


<#User defined variables#>
$computer = 'computer'

<#Start invoke session#>>
Invoke-Command -ComputerName $Computer -ScriptBlock {
    $i = 0

    <#Specify the drive letter you wish to assign#>
    $Drive = 'L'
    <#Specify the drive label you wish to assign#>
    $Label = "LOGS"
    <#Specify the partition type#>
    $partition = "GPT"
    <#Specify the filesystem type to format the volume with#>
    $fs = "ntfs"

    <#Get the disks with 0 partitions#>
    $disks = Get-WmiObject win32_diskdrive | Select-Object name, partitions | Where-Object { $_.partitions -eq 0 }
    $disks | ForEach-Object { $i++ }
    IF ($i -ne 1) {
        Write-host " Found $i disks which is greater than the expected number of raw disks which should be 1." -ForegroundColor Yellow
        break
    }
    <#Identifies the disk number of the disk with 0 partitions#>
    $diskNumber = $disks.name -replace '^\\*.\\*[A-Za-z]*'

    <#Set the SAN policy to online. The format may error if this is not set#>
    "SAN Policy=OnlineAll" | diskpart

    <#Bring the disk online
    ***VERIFY THAT THIS IS THE DISK YOU WANT TO FORMAT***#>
    "Select Disk $diskNumber",
    "online disk",
    "convert $partition",
    "attributes disk clear readonly",
    "CREATE PARTITION PRIMARY" | 
    diskpart
    <#List the volumes on the system#>
    $volumeOutput = "SELECT disk $diskNumber", "LIST Volume" | diskpart.exe
    
    <#Loop through the string data returned with regex and choose the line containing the RAW disk#>
    Foreach ($line in $volumeOutput) {
        $volume = $null
        $vNumber = $null
        $vFS = $null
        $vType = $null
        $vSize = $null
        $vStatus = $null
        $vRegexPattern = "(?<volume>^\s+[A-Za-z]+\s)(?<number>\d\s+)(?<FS>[A-Za-z]+\s+)(?<type>[A-Za-z]+\s+)(?<Size>\d+\s[A-Za-z]+\s+)(?<Status>[A-Za-z]+\s+)"
        $vRegexMatch = [regex]::Matches($line, $vRegexPattern)
        if ($vRegexMatch[0].Groups.count -eq 7) {
            $volume = ($vRegexMatch[0].Groups['volume'].Value).trim()
            $vNumber = ($vRegexMatch[0].Groups['number'].Value).trim()
            $vFS = ($vRegexMatch[0].Groups['FS'].Value).trim()
            $vType = ($vRegexMatch[0].Groups['type'].Value).trim()
            $vSize = ($vRegexMatch[0].Groups['Size'].Value).trim()
            $vStatus = ($vRegexMatch[0].Groups['Status'].Value).trim()
            if ($vFS -eq "RAW") {
                Write-host "$Volume $vNumber is $vFS, is a $vType, is $vSize, and is $vStatus." -foregroundcolor magenta
                break
            }
        }
    }
    "select volume $vNumber", "format fs=$fs Label=$label QUICK", "Assign Letter = $drive" | diskpart.exe
    
  
}
    
    

