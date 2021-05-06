
######import modules ######
import-module powervault 
Add-Type -AssemblyName Microsoft.VisualBasic
######### current user for vault session #################################################################

open-vaultconnection -server "wal-vaultarc" -vault "Wendt" -user "coolorange" -password "nhg544FK"
$Foldereport = "C:\Vault Data\Vault Reports\Other Reports\archiveFolders.csv"
$clear = clear-content $Foldereport


$ht1 = "Foldername"
$ht2 = "Full Path"
$ht3 = "ID"
$ht4 = 'Value'
$folderrow = "" |select-Object $ht1, $ht2,$ht3,$ht4
$folderout = @()



$toplevelpathID = $vault.DocumentService.GetFolderByPath("$/Wendt")
$subfolderid = $vault.DocumentService.GetFolderIdsByParentIds($toplevelpathID.Id,$false)

foreach($item in $subfolderid){
$foldername = $vault.DocumentService.GetFolderById($item)
$folderN = $foldername.Name
$folderFP = $foldername.FullName
$folderID = $foldername.Id
$folderrow.$ht1 = $folderN
$folderrow.$ht2 = $folderFP
$folderrow.$ht3 = $folderID
$folderrow.$ht4 = '$True'
$folderrow | Export-Csv -Append  -NoTypeInformation -Path $Foldereport -force
}

