$komputery = Get-ComputerInfo

$siec = Get-NetIpConfiguration

$prog = Get-WmiObject -class win32_product | Select-Object -Property name



write-output ("Nazwa Uzytkownika        :{0}" -f $komputery.csusername)
write-output ("Nazwa Komputera          :{0}" -f $komputery.CSName)
write-output ("Nazwa Domeny             :{0}" -f $komputery.csdomain)
write-output ("Zainstalowany system os  :{0}" -f $komputery.Windowsproductname)

write-output $siec

write-output $prog | ft

#(Get-WmiObject -query "Select * from SoftwareLicensingService").OA3xOrginalProductKey

read-host -prompt "Press enter"