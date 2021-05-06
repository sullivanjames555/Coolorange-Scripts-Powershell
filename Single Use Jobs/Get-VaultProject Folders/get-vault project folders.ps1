 
######import modules ######
import-module powervault 
Add-Type -AssemblyName Microsoft.VisualBasic
add-type -AssemblyName System.Windows.Forms
import-module "C:\Users\Public\Documents\New-VaultConnection.psm1"
import-module "C:\Users\Public\Documents\New-cooltable.psm1"

$server = "Wendt-vault"
$vaultserver =  "Wendt"
$username = "coolorange"
$vaultpw = "nhg544FK"

new-vaultconnection -server $server -vault $vaultserver -user $username  -password $vaultpw

New-CoolTable -TableName "Vaultprojects" -ColumnNames ("Foldername","Full Path")
$vaultname = $vaultconnection.Server

$date = get-date -Format "MM-dd-yy"
$vaultcsv = "C:\Vault Data\Vault Reports\Other Reports\$vaultname Folders - $date.csv"

$mainfolder = $vault.DocumentService.getfolderbypath("$/Wendt")
$folders = $vault.DocumentService.GetFoldersByParentId($mainfolder.id,$false)

foreach($folder in $folders){


if($folder){
$itemid = $folder.id
$vfolder = $vault.DocumentService.GetFolderById($itemid)
$itemname = $vfolder.Name
$itempath = $vfolder.Fullname
$itemid = $vfolder.id


$Vaultprojects.Rows.add($itemname,$itempath)


}}
$Vaultprojects |export-csv -NoTypeInformation -Path  $vaultcsv -Append -Force