# Ustawienie daty
$today = (Get-Date).ToString("dd.MM.yyyy")

# Win32_OperatingSystem dane o OS
$w32_OS = Get-CimInstance -Class Win32_OperatingSystem

# Win32_Processor
$w32_CPU = Get-CimInstance -Class Win32_Processor

# Win32_ComputerSystem
$w32_Comp = Get-CimInstance -Class Win32_ComputerSystem

# Win32_BIOS
$w32_BIOS = Get-CimInstance -Class Win32_BIOS

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

$zmPC = [PSObject]@{
    pcAudytDate = $today
    pcName = $ENV:COMPUTERNAME
    pcUserName = $ENV:USERNAME
    pcUserDomain = $ENV:USERDOMAIN
    pcBiosSerial = ($w32_BIOS.SerialNumber).Trim()
    pcMan = ($w32_Comp.Manufacturer).Trim()
    pcModel = ($w32_Comp.Model).Trim() 
    pcCPU = $w32_CPU.Name
    pcMem = [math]::Ceiling(($w32_Comp.TotalPhysicalMemory)/1024/1024/1024)
    pcOS = $w32_OS.Caption
    pcSerialOS = $w32_OS.SerialNumber
    pcInstallDate = $w32_OS.InstallDate.ToString("dd.MM.yyyy HH:mm")
    pcPrograms = $programy
    pcIP = $adresy
} # | ConvertTo-Json


$v1 = $zmPC.pcModel | Out-String
$v1_1 = $zmPC.pcMan | Out-String
$v2 = $zmPC.pcBiosSerial | Out-String
$v3 = $zmPC.pcName | Out-String
$v4 = $zmPC.pcIP | forEach { $_.IPAddress + ' '} | Out-String
$v5 = $zmPC.pcCPU | Out-String
$v6 = $zmPC.pcMem | Out-String
$v7 = $zmPC.pcOS | Out-String
$v8 = $zmPC.pcSerialOS | Out-String


$p1 = $zmPC.pcPrograms | Where-Object {$_.Name -notlike "Microsoft Visual C*"} | Select-Object -Property Name,Version | ConvertTo-Html -Fragment

#Write-Output $zmPC 
#Write-Output $adresy

$css = @"
<style>
* {
    margin: 0;
    padding: 0;
    font-size: 10pt;
    line-height: 2;
}

h1 {
    text-align: center;
    font-size: 12pt;
}

h2 {
    text-align: center;
    font-size: 10pt;
}

h4 {
    font-size: 10pt;
}

p {
    font-size: 8pt;
}

table {
    width: 100%;
    border: 1px solid #000;
    border-collapse: collapse;
}

table td {
    border: 1px solid #000;
    padding-left: 3pt;
    font-size: 8pt;
}

.small {
    text-align: center;
    font-size: 6pt;
}

input {
    border: none;
    width: 100%;

}

@media print {
    @page {
        size: A4;
        margin: 8mm;
    }

    @page :left {
        margin-right:18mm;
    }

    @page :right {
        margin-left:18mm;
    }

    body {
        width:100%;
        font-size: 9pt;
        line-height: 2;
    }

    h1 {
        text-align: center;
        font-size: 12pt;
    }
    
    h2 {
        text-align: center;
        font-size: 10pt;
    }
    
    h4 {
        font-size: 9pt;
    }
    
    p {
        font-size: 8pt;
    }
    
    table {
        width: 100%;
        border: 1px solid #000;
        border-collapse: collapse;
    }
    
    table td {
        border: 1px solid #000;
        padding-left: 3pt;
        font-size: 8pt;
    }
    
    .small {
        text-align: center;
        font-size: 6pt;
    }
}
</style>
"@

$page = @"
<section>
    <p class="small">SP ZOZ Wojewówdzki Szpittal Specjalistyczny nr 3 w Rybniku</p>    
    <h1>Karta stanowiska pracy</h1>
    <h2>Dane komputera / stacji komputerowej / zestawu komputerowego</h2>
    <table>
        <thead>
            <tr>
                <th widt="40%">Rodzaj wyposażenia</th>
                <th>Cecha (model, nr seryjny lub inwentarzowy)</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>Komputer</td>
                <td><input type="text" name="pcModel" id="pcModel" value="$v1 | $v1_1"></td>
                <td><input type="text" name="pcBiosSerial" id="pcBiosSerial" value="$v2"></td>
            </tr>
            <tr>
                <td>Nazwa | adres IP</td>
                <td><input type="text" name="pcName" id="pcName" value="$v3"></td>
                <td><input type="text" name="pcIP" id="pcIP" value="$v4"></td>
            </tr>
            <tr>
                <td>Processor | Pamięć</td>
                <td><input type="text" name="pcCPU" id="pcCPU" value="$v5"></td>
                <td><input type="text" name="pcMem" id="pcMem" value="$v6 GB"></td>
            </tr>
            <tr>
                <td>Zasilacz awaryjny</td>
                <td colspan="2"><input type="text" name="Ups" id="Ups" value=""></td>
            </tr>
            <tr>
                <td>Drukarka</td>
                <td colspan="2"><input type="text" name="Druk" id="Druk" value=""></td>
            </tr>
            <tr>
                    <td>Skaner</td>
                    <td colspan="2"><input type="text" name="Skaner" id="Skaner" value=""></td>
                </tr>
                <tr>
                    <td>Urządzenie wielofunkcyjne</td>
                    <td colspan="2"><input type="text" name="Uw" id="Uw" value=""></td>
                </tr>
                <tr>
                    <td>Nr ewidencyjny</td>
                    <td colspan="2"><input type="text" name="Druk" id="Druk" value="$nrEwid"></td>
                </tr>
            </tbody>
        </table>
        <p>Oprogramowanie</p>
        <table>
            <thead>
                <tr>
                    <th width="30%">System Operacyjny</th>
                    <th>Licencja</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td><input type="text" name="pcOS" id="pcOS" value="$v7"></td>
                    <td><input type="text" name="pcSerialOS" id="pcSerialOS" value="$v8"></td>
                </tr>
            </tbody>
        </table>
        <br>
"@

$page3 = @"
<br>
<table>
<thead>
    <tr>
        <th width="100%" style="text-align:left;">Inne:</th>
    </tr>
</thead>
<tbody>
    <tr>
        <td>$lokalizacja</td>
    </tr>
    <tr>
        <td>&nbsp;</td>
    </tr>
    <tr>
        <td>&nbsp;</td>
    </tr>
    <tr>
        <td>&nbsp;</td>
    </tr>
</tbody>
</table>
<p>kartę przygotował:</p>
<p class="small">....................................................</p>
<p class="small">Data i podpis</p>
</section>
<br>
<br>
<section>
<h1>OŚWIADCZENIA</h1>
<p>1. Przyjmuję na stan sprzęt komputerowy wg. powyższej specyfikacji.</p>
<p>2. Nie jestem upoważniona(y) do dokonywania ingerencji w przekazany mi sprzęt komputerowy.</p>
<p>3. Przyjmuję do wiadomości, że samodzielne instalowanie oprogramowania w powierzonym komputerze jest
niedozwolone.</p>
<table>
<thead>
    <tr>
        <th>lp.</th>
        <th>Data</th>
        <th>Imię, Nazwisko (użytkownika</th>
        <th>Podpis użytkownika</th>
    </tr>
</thead>
<tbody>
    <tr>
        <td>1.</td>
        <td><input type="text" name="audytdate" id="audytdate" value="$today"></td>
        <td><input type="text" name="osoba" id="osoba" value="$odpowiedzialny"></td>
        <td></td>
    </tr>
    <tr>
        <td>2.</td>
        <td></td>
        <td></td>
        <td></td>
    </tr>
    <tr>
        <td>3.</td>
        <td></td>
        <td></td>
        <td></td>
    </tr>
    <tr>
        <td>4.</td>
        <td></td>
        <td></td>
        <td></td>
    </tr>
    <tr>
        <td>5.</td>
        <td></td>
        <td></td>
        <td></td>
    </tr>
    <tr>
        <td>6.</td>
        <td></td>
        <td></td>
        <td></td>
    </tr>
</tbody>
</table>
</section>
"@


ConvertTo-Html -Title "Strona Html" -Body "$page$p1$page3" -Head "$css" | Out-File .\PCTest.html
Invoke-Item .\PCTest.html