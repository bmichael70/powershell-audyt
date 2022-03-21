# Ustawienie daty
$today = (Get-Date).ToString("dd.MM.yyyy")

# Zainstalowane programy
$programy = Get-Package -ProviderName Programs,msi | Select-Object Name,Version,Providername

# Adresy IPv4
$adresy = Get-NetIPAddress | 
    Where-Object { $_.AddressFamily -eq "IPv4" -and 
        (
            $_.InterfaceAlias -icontains "Ethernet" -or 
            $_.InterfaceAlias -icontains "Wi-fi"
        )} | Select-Object IPAddress,InterfaceAlias

# Tworzenie obiektu JSON

$zmienna = [PSObject]@{
    pcDate = $today
    pcName = $ENV:COMPUTERNAME
    pcUserName = $ENV:USERNAME
    pcUserDomain = $ENV:USERDOMAIN
    pcPrograms = $programy
    pcAdresy = $adresy
} | ConvertTo-Json

Write-Output $zmienna
#Write-Output $adresy