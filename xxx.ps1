$computerInfo = Get-ComputerInfo

$file = ($computerInfo).CsName

Write-Output $file