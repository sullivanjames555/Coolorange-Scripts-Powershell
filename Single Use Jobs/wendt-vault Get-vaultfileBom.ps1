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
New-CoolTable -TableName FileBomtable -ColumnNames "Bom_RowOrder","Bom_PositionNumber","Bom_Number","Bom_Structure","Bom_Quantity","Bom_ItemQuantity","Bom_UnitQuantity","Bom_Unit","Bom_Material","File Name"


$file = get-vaultfiles -FileName "907-117-4-9014.iam"
$filename = $file.name

$fileBom = Get-VaultFileBOM -File $file._FullPath -GetChildrenBy LatestReleasedVersion


#$filebomtable = $fileBom | Format-Table Bom_RowOrder,Bom_PositionNumber,Bom_Number,Bom_Structure,Bom_Quantity,Bom_ItemQuantity,Bom_UnitQuantity,Bom_Unit
foreach($itemv in $filebom){
$filebomtable.rows.add($itemv.Bom_RowOrder,$itemv.Bom_PositionNumber,$itemv.Bom_Number,$itemv.Bom_Structure,$itemv.Bom_Quantity,$itemv.Bom_ItemQuantity,$itemv.Bom_UnitQuantity,$itemv.Bom_Unit,$itemv.Bom_Material,$itemv._Name)

$filebomtable | export-csv -Path "C:\Vault Data\Vault Reports\Other Reports\$filename bom.csv" -append -NoTypeInformation
}
