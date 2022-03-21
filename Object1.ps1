$prog = Get-WmiObject -query "select * from Win32_Product" #| Format-Table -Property Name -HideTableHeaders

$Name = (Get-ComputerInfo).CsName
$file = ("{0}.json" -f $Name)

$object = New-Object -TypeName PSObject
$object | Add-Member -MemberType NoteProperty -Name OSBuild -Value 'OSBuild'
$object | Add-Member -MemberType NoteProperty -Name OSVersion -Value 'Version'
$object | Add-Member -MemberType NoteProperty -Name PCName -Value $Name
$object | Add-Member -MemberType NoteProperty -Name Progs -Value $prog.Name

$object | ConvertTo-Json | Out-File $file

Write-Host $object | ConvertTo-Json