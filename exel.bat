
$InputFile = "./adresy.xlsx"
$OutputFile = "./wyniki.xlsx"

$SheetName = "IP-Addresses"

if (-Not (Test-Path $InputFile)) {
    Write-Host "plik wejściowy $InputFile nie istnieje."
    exit 1
}


try {
    $IPAddresses = Import-Excel -Path $InputFile -WorksheetName $SheetName |
                   Select-Object -First 5 -ExpandProperty 'Column1'
} catch {
    Write-Host "nie można odczytać pliku Excel lub arkusza $SheetName. Upewnij się, że plik istnieje i zawiera odpowiednie dane."
    exit 1
}

$Results = @()

foreach ($IP in $IPAddresses) {
    if ([System.Net.IPAddress]::TryParse($IP, [ref]$null)) {
        try {
            $PingResult = Test-Connection -ComputerName $IP -Count 1 -ErrorAction Stop
            $ResultText = "Reply from $($PingResult.Address): Time=$($PingResult.ResponseTime)ms"
        } catch {
            $ResultText = "Błąd: Nie można połączyć się z $IP."
        }
    } else {
        $ResultText = "Nieprawidłowy adres IP: $IP"
    }

    $Results += [PSCustomObject]@{
        Adres = $IP
        Wynik = $ResultText
    }
}

try {
    $Results | Export-Excel -Path $OutputFile -WorksheetName "Wyniki" -AutoSize
    Write-Host "Wyniki zapisano do pliku: $OutputFile"
} catch {
    Write-Host "Nie można zapisać wyników do pliku Excel: $_"
    exit 1
}
