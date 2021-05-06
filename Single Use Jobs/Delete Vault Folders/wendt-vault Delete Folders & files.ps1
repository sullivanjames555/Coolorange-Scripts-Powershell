######import modules ######
import-module powervault 
Add-Type -AssemblyName Microsoft.VisualBasic
add-type -AssemblyName System.Windows.Forms
import-module "C:\Users\Public\Documents\New-VaultConnection.psm1"
import-module "C:\Users\Public\Documents\New-cooltable.psm1"

$date = get-date -Format "MM-dd-y"
#when this is false it means it will not delete the folder  and only output the result #
$deleted = $true
if($deleted -eq $true){write-host "are you sure you want to delete files from active vault?"
$deleteanswer = read-host "Write yes or no"}
if($deleteanswer -ilike "no"){$delete = $false}

$server = "Wendt-vault"
$vaultserver =  "Wendt"
$username = "coolorange"
$vaultpw = "nhg544FK"


function DeleteFile
{
    Param($folderID)
    if($deleted)
    {
        try
        {
            $vault.DocumentService.DeleteFolderHierarchyUnconditional($folderID)
            $result = "Succeeded"
        }
        catch
        {
            $result = "Failed $_"
        }
    }
    else
    {
        $result = "Simulated"
    }
    
    return $result
}

######### current user for vault session #################################################################
new-vaultconnection -server $server -vault $vaultserver -user $username  -password $vaultpw

if($vaultconnection.Server -notlike "Wendt-vault"){"not connected to Wendt-Vault"}
else{
$csv = import-csv "C:\Vault Data\Script Input\Delete Folders\Wendt-Vault To Be Deleted.csv" -Header "Foldername", "full Path", "ID","Delete" | Select-Object -Skip 1
$deletereportcsv = "C:\Vault Data\Vault Reports\Other Reports\Wendt-Vault Folders Deleted - $date.csv"

New-CoolTable -TableName "DeleteReport" -ColumnNames ("Foldername","Full Path","ID","Deleted")


foreach($item in $csv){

if ($item.Delete -eq '$true'){
$foldername1 = $item.Foldername
if($foldername1){
$folderpath = "$/Wendt/$foldername1"
try{
$itemfolder = $vault.DocumentService.GetFolderByPath($folderpath)}
catch{
$itemfolder = $false
$itemerror = $_}
}
if($itemfolder){
$itemid = $itemfolder.id
$folder = $vault.DocumentService.GetFolderById($itemid)
$itemname = $folder.Name
$itempath = $folder.Fullname
$itemid = $folder.id

$result = DeleteFile -folderID $itemid -verbose

}
else {
$itemname = $item.Folder
$itempath = $folderpath
$itemid = "0"
$result = "there was an error"}

$deletereport.Rows.add($itemname,$itempath,$itemid,$result)


}

}
$deletereport |export-csv -NoTypeInformation -Path  $deletereportcsv -Append
}
