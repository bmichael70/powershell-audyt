Clear-Host

$komputer = Get-ComputerInfo

$prog = Get-WmiObject -Class Win32_Product | Format-Table -Property Name -HideTableHeaders

#$siec = Get-NetIPConfiguration
$IP_address = (Get-NetIPConfiguration).IPv4Address.IPAddress

$keyWin = (Get-WmiObject -query "Select * from SoftwareLicensingService").OA3xOrginalProductKey

#Write-Output ("Nazwa uzytkownika     : {0}" -f $env:USERNAME)
Write-Output (" ")
Write-Output ("--------------------------")
Write-Output (" ")
Write-Output ("{0} | {1}" -f ([string]$komputer.CSModel).Trim(), [string]$komputer.CSManufacturer)
Write-Output ("{0} | {1}" -f ([string]$komputer.CSName).Trim(), [string]$IP_address)
#Write-Output ("Nazwa domeny          : {0}" -f ([string]$komputer.CSDomain).Trim())
Write-Output ("{0}" -f ([string]$komputer.OSName).Trim())  # WindowsProductName
Write-Output ("key:   {0}" -f $keyWin)
Write-Output ("--------------------------")
Write-Output ("Zainstalowane programy")
Write-Output $prog

Read-Host -Prompt "Press Enter Key"

#$prog.GetType()
#$prog.Length
#$prog.Count
#$prog | ConvertTo-Json

#$pc = New-Object PSObject
#$pc | Add-Member "PCName" $komputer.CSName