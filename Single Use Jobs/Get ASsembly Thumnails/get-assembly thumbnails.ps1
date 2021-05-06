import-module powervault
open-vaultconnection -server "wendt-vault" -vault "wendt" -user "coolorange" -password "nhg544FK"

$csv = import-csv -Path "C:\Vault Data\Script Input\Get-Thumbnails\get-thumbnails.csv" -Header File
$FileName = "C:\Vault Data\9-10 Thumbnails.csv"

foreach($item in $csv.File){

$file = Get-VaultFiles -FileName $item -properties @{'Latest Version' = "$true"}
        
$output = @()
$h1 = 'Name' 
$h2 = 'Thumbnail'
$row = "" |select-Object $h1, $h2


$image = $file.Thumbnail.Image

$fname = $file._PartNumber

$image2 = [convert]::ToBase64String($image)
$row.$h1 = $file.Name
$row.$h2 = $image2

$output = $row 


$Output | Export-Csv -Append -Encoding ascii  -NoTypeInformation -Path $FileName -force
#Set-Content -Path "P:\Hossack\images\$fname.png" -Value $image -Encoding Byte


}


