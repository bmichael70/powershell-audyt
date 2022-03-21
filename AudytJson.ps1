Clear-Host
$today = (Get-Date).ToString("yyyy-MM-dd HH:mm")
Write-Output ("DIS WSS nr 3 Rybnik Audyt {0}" -f $today)
Write-Output "==========================================="
Write-Output " "
$nrEwid = Read-Host -Prompt "Podaj proszę nr ewidencyjny/IL lub wciśnij enter"
Write-Output " "
$lokalizacja = Read-Host -Prompt "Podaj lokalizację komputera "
Write-Output " "
Write-Output "Proszę czekać pracuję ... :)"


$bios = Get-CimInstance -ClassName Win32_BIOS
$computerInfo = Get-ComputerInfo
$prog = Get-CimInstance -ClassName Win32_Product #| Format-Table -Property Name -HideTableHeaders
$net = Get-NetIPConfiguration

$pcName = ($computerInfo).CsName

$file = ("{0}.json" -f $pcName)

$object = New-Object -TypeName PSObject
$object | Add-Member -MemberType NoteProperty -Name PCName -Value $pcName
$object | Add-Member -MemberType NoteProperty -Name IPaddr -Value $net.IPv4Address.IPAddress
$object | Add-Member -MemberType NoteProperty -Name Model -Value ("{0} | {1}" -f ([string]$computerInfo.CSModel).Trim(), [string]$computerInfo.CSManufacturer)
$object | Add-Member -MemberType NoteProperty -Name Serial -Value ("{0}" -f ([string]$bios.SerialNumber).Trim())
$object | Add-Member -MemberType NoteProperty -Name BiosVersion -Value ("{0}" -f ([string]$bios.Version).Trim())
$object | Add-Member -MemberType NoteProperty -Name NrEwid -Value $nrEwid
$object | Add-Member -MemberType NoteProperty -Name Lokalizacja -Value $lokalizacja
$object | Add-Member -MemberType NoteProperty -Name AudytDate -Value $today
$object | Add-Member -MemberType NoteProperty -Name Progs -Value $prog.Name

$object | ConvertTo-Json | Out-File ("\\ad1\wrzutowy\00json\{0}" -f $file)

Write-Output $object | ConvertTo-Json
Write-Output ("===============================")
Write-Warning "GOTOWE!!"
Write-Output ("Sprawdz: \\ad1\wrzutowy\00json\{0}" -f $file)

Read-Host -Prompt "Nacisnij klawisz Enter!"
