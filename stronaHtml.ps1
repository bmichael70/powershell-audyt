Clear-Host
$today = (Get-Date).ToString("dd-MM-yyyy")
Write-Output ("DIS WSS nr 3 Rybnik Audyt {0}" -f $today)
Write-Output "==========================================="
Write-Output " "
$nrEwid = Read-Host -Prompt "Podaj proszę nr ewidencyjny/IL lub wciśnij enter"
Write-Output " "
$lokalizacja = Read-Host -Prompt "Podaj lokalizację komputera "
Write-Output " "
$odpowiedzialny = Read-Host -Prompt "Podaj imię nazwisko użytkownika "
Write-Output " "
Write-Output "Proszę czekać pracuję ... :)"


$bios = Get-CimInstance -ClassName Win32_BIOS
$computerInfo = Get-ComputerInfo
$prog = Get-CimInstance -ClassName Win32_Product #| Format-Table -Property Name -HideTableHeaders
$net = Get-NetIPConfiguration


$pc = ([string]$computerInfo.CSModel).Trim()
$pcSerial = ([string]$bios.SerialNumber).Trim()
$pcIP = ([string]$net.IPv4Address.IPAddress)
$pcName = ($computerInfo).CsName

#$file = ("{0}.json" -f $pcName)

#$object = New-Object -TypeName PSObject
#$object | Add-Member -MemberType NoteProperty -Name PCName -Value $pcName
#$object | Add-Member -MemberType NoteProperty -Name IPaddr -Value $net.IPv4Address.IPAddress
#$object | Add-Member -MemberType NoteProperty -Name Model -Value ("{0} | {1}" -f ([string]$computerInfo.CSModel).Trim(), [string]$computerInfo.CSManufacturer)
#$object | Add-Member -MemberType NoteProperty -Name Serial -Value ("{0}" -f ([string]$bios.SerialNumber).Trim())
#$object | Add-Member -MemberType NoteProperty -Name BiosVersion -Value ("{0}" -f ([string]$bios.Version).Trim())
#$object | Add-Member -MemberType NoteProperty -Name NrEwid -Value $nrEwid
#$object | Add-Member -MemberType NoteProperty -Name Lokalizacja -Value $lokalizacja
#$object | Add-Member -MemberType NoteProperty -Name AudytDate -Value $today
#$object | Add-Member -MemberType NoteProperty -Name Progs -Value $prog.Name


$css = @"
<style>
* {
    margin: 0;
    padding: 0;
    font-size: 10pt;
    line-height: 2;
}

section {
    break-after: page;
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
    font-size: 10pt;
}

table {
    width: 100%;
    border: 1px solid #000;
    border-collapse: collapse;
}

table td {
    border: 1px solid #000;
    padding-left: 5pt;
}

.small {
    text-align: center;
    font-size: 6pt;
}

@media print {
* {
    margin: 0;
    padding: 0;
    font-size: 10pt;
    line-height: 2;
}

@page {
    margin: 2cm;
}

section {
    break-after: page;
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
    font-size: 10pt;
}

table {
    width: 100%;
    border: 1px solid #000;
    border-collapse: collapse;
}

table td {
    border: 1px solid #000;
    padding-left: 5pt;
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
                    <td>$pc | $pcSerial</td>
                </tr>
                <tr>
                    <td>Nazwa komputera / adres IP</td>
                    <td>$pcName | $pcIP</td>
                </tr>
                <tr>
                    <td>Monitor</td>
                    <td></td>
                </tr>
                <tr>
                    <td>Zasilacz awaryjny</td>
                    <td></td>
                </tr>
                <tr>
                    <td>Drukarka</td>
                    <td></td>
                </tr>
                <tr>
                    <td>Skaner</td>
                    <td></td>
                </tr>
                <tr>
                    <td>Urządzenie wielofunkcyjne</td>
                    <td></td>
                </tr>
                <tr>
                    <td>Nr ewidencyjny</td>
                    <td>$nrEwid</td>
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
                    <td>Windows 10 Pro</td>
                    <td></td>
                </tr>
            </tbody>
        </table>
        <br>
        <table>
            <thead>
                <tr>
                    <th width="30%">Pakiet biurowy</th>
                    <th>Licencja</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>LibreOffice</td>
                    <td></td>
                </tr>
            </tbody>
        </table>
        <br>
        <table>
            <thead>
                <tr>
                    <th width="30%">Oprogramowanie</th>
                    <th>Licencja</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td>TightVNC</td>
                    <td></td>
                </tr>
                <tr>
                    <td>ESET</td>
                    <td></td>
                </tr>
                <tr>
                    <td>Adobe Reader</td>
                    <td></td>
                </tr>
            </tbody>
        </table>
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
                    <td>$today</td>
                    <td>$odpowiedzialny</td>
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

ConvertTo-Html -Title "Strona Html" -Body "$page" -Head "$css" | Out-File .\StronaPage.html