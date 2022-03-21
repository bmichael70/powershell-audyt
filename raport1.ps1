param(
    $computername = 'localhost',
    $filepath = '.\report.html'
)

# Get OS Info
$os = Get-wmiObject -class Win32_OperatingSystem -ComputerName $computername |
    Select-Object BuildNumber,Caption,ServicePackMajorVersion,ServicePackMinorVersion |
    ConvertTo-Html -Fragment -As List -PreContent "Generated $(Get-Date)<br><br><h2>Operating System</h2>" |
    Out-String

# Get Hardware info
$comp = Get-wmiObject -class Win32_ComputerSystem -ComputerName $computername |
    Select-Object DNSHostname, Domain,DomainRole,Manufacturer,Model,Name,NumberOfLogicalProcessors,TotalPhysicalMemory |
    ConvertTo-Html -Fragment -As List -PreContent "<h2>Hardware</h2>" |
    Out-String

# Get service list
#$services = Get-WmiObject -Class Win32_Service -ComputerName $computername |
#    Select-Object Name,State,StartMode |
#    ConvertTo-Html -Fragment -As Table -PreContent "<h2>Services</h2>" |
#    Out-String

# Get Product list
$products = Get-WmiObject -Class Win32_Product -ComputerName $computername |
    Select-Object Name,Version |
    ConvertTo-Html -Fragment -As Table -PreContent "<h2>Programs</h2>" |
    Out-String


# Combine HTML
$final = ConvertTo-Html -Title "System information for $computername" -PreContent $os,$comp,$products -Body "<h1>Information for $computername</h1>"

$final | Out-File $filepath
