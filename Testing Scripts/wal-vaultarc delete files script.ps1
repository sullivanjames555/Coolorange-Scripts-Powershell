######import modules ######
import-module powervault 
Add-Type -AssemblyName Microsoft.VisualBasic
add-type -AssemblyName System.Windows.Forms
import-module "C:\Users\Public\Documents\New-VaultConnection.psm1"
import-module "C:\Users\Public\Documents\New-cooltable.psm1"

$simulated = $true

$server = "wal-vaultarc"
$vault =  "Wendt"
$username = "coolorange"
$vaultpw = "nhg544FK"


function DeleteFile
{
    Param($folderID)
    if($deleteFiles)
    {
        try
        {
            $vault.DocumentService.DeleteFileFromFolder($folderID)
            $result = "Succeeded"
        }
        catch
        {
            $result = "Failed"
        }
    }
    else
    {
        $result = "Simulated"
    }
    
    return $result
}

######### current user for vault session #################################################################
new-vaultconnection -server $server -vault $vault  -user $username  -password $vaultpw

if($vaultconnection.Server -notmatch "wal-vaultarc"){"not connected to vaultarc"}
else{
$csv = import-csv "\\wal-vaultcron\Vault Data\Vault Reports\Other Reports\ActiveVault_Folders Test.csv" -Header "Full Path", "Delete" | Select-Object -Skip 1
$deletereport = "C:\Vault Data\Vault Reports\Other Reports\Folders Deleted.csv"

New-CoolTable -TableName "DeleteReport" -ColumnNames ("Foldername","Full Path","ID","Deleted","Error")


foreach($item in $csv){

if ($item.Delete -eq $true){
$itemfolder = $vault.DocumentService.GetFolderByPath($item.'Full Path')
$itemid = $itemfolder.id
$folder = $vault.DocumentService.GetFolderById($itemid)
$itemname = split-path $csv.'Full Path' -leaf
$itempath = $folder.Fullname
$itemid = $folder.id

try{
if($simulated -eq $false){
$deleted = $vault.DocumentService.DeleteFolderHierarchyUnconditional($itemid)}
else{}
catch{$booleanerror = $true ; $deleteerror += $_}
if($booleanerror -eq  $false){
$deletereport.Rows.add($itemname,$itempath,$itemid,$boolean,"0")


}
}
else{
$deletereport.Rows.add($itemname,$itempath,$itemid,$boolean,$deleteerror)



}

$deletereport |export-csv -NoTypeInformation -Path  $deletereport -Append
}
}
