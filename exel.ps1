
$InputFile = "./adresy.xlsx"
$OutputFile = "./wyniki.xlsx"

$SheetName = "IP-Addresses"

if (-Not (Test-Path $InputFile)) {
    Write-Host "plik wejsciowy $InputFile nie istnieje."
    exit 1
}


try {
    $IPAddresses = Import-Excel -Path $InputFile -WorksheetName $SheetName |
                   Select-Object -First 5 -ExpandProperty 'Adresy'
} catch {
    Write-Host "nie mozna odczytac pliku Excel lub arkusza $SheetName. Upewnij sie, ze plik istnieje i zawiera odpowiednie dane."
    exit 1
}

$Results = @()

foreach ($IP in $IPAddresses) {
    if ([System.Net.IPAddress]::TryParse($IP, [ref]$null)) {
        try {
            $PingResult = Test-Connection -ComputerName $IP -Count 1 -ErrorAction Stop
            $ResultText = "reply from $($PingResult.Address): Time=$($PingResult.ResponseTime)ms"
        } catch {
            $ResultText = "blad: Nie mozna polaczyc sis z $IP."
        }
    } else {
        $ResultText = "nieprawidlowy adres IP: $IP"
    }

    $Results += [PSCustomObject]@{
        Adres = $IP
        Wynik = $ResultText
    }
}

try {
    $Results | Export-Excel -Path $OutputFile -WorksheetName "wyniki" -AutoSize
    Write-Host "wyniki zapisano do pliku: $OutputFile"
} catch {
    Write-Host "nie mozna zapisac wynikow do pliku Excel: $_"
    exit 1
}
