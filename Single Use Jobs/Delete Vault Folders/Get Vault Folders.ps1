 
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

New-CoolTable -TableName "DeleteReport" -ColumnNames ("Foldername","Full Path","ID","Deleted")


$csv = "C:\Vault Data\Script Input\Delete Folders\Wendt-Vault To Be Deleted.csv"
$deletereportcsv = "C:\Vault Data\Vault Reports\Other Reports\Wendt-Vault Folders Deleted - $date.csv"
$folders = @("M945-001 - Nobel","M945-002 - Southern Waste Systems","M945-003 - Rochester Iron and Metal","M945-004 - BN","M945-005 - Greer","M945-006 - Triple M","M945-007 - IPI","M945-008 - John Ross & Sons - 32496","M945-009 - Weitsman - 31926","M945-010 - Chisick Metals - 31657")

foreach($folder in $folders){
if($folder){
$folderpath = "$/Wendt/$folder"
try{
$itemfolder = $vault.DocumentService.GetFolderByPath($folderpath)}
catch {"Error getting folder"}

if($itemfolder){
$itemid = $itemfolder.id
$vfolder = $vault.DocumentService.GetFolderById($itemid)
$itemname = $vfolder.Name
$itempath = $vfolder.Fullname
$itemid = $vfolder.id
$result = $true

$deletereport.Rows.add($itemname,$itempath,$itemid,$result)


}}}
$deletereport |export-csv -NoTypeInformation -Path  $csv -Append