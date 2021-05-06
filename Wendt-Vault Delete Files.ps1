Import-Module powervault
Import-Module "C:\ProgramData\coolOrange\powerJobs\Modules\VaultSearch.psm1"
open-vaultconnection -server "wendt-vault" -vault "wendt" -user "coolorange" -password "nhg544FK"
$date = Get-Date 
$deletedSearch =  New-SearchCondition -PropertyName Name -searchOperator:Contains -searchRuleType:Must -searchText "Delete*"
$notobsolete = New-SearchCondition -PropertyName State -searchOperator:NotContains -searchText "Released"
$csvfile = "C:\Vault Data\Vault Reports\Deleted List\Deleted Files.csv"
$result = Find-VaultFiles -RootFolderPath "$/Wendt" -SearchConditions @($deletedsearch , $notobsolete) -LatestFilesOnly -RecurseFolders
$deletedfolder =  $vault.DocumentService.GetFolderByPath("$/Wendt/ToBeDeleted")
$ErrorActionPreference = "SilentlyContinue"

############## Creating a report Table #########################
$h1 = "File Name"
$h2 = "File Path"
$h3 = "Time Deleted"

$row = "" |select-Object $h1, $h2,$h3
$output = @()

####################################################

foreach($item in $result){
$folder = $vault.DocumentService.GetFoldersByFileMasterId($item.MasterId)
$folderpath = $folder.FullName
$name = $item.name
$row.$h1 = $name
$row.$h2 = $folderpath
$row.$h3 = $date
$output = $row 
$Output | Export-Csv -Append  -NoTypeInformation -Path $csvfile -force

$item.Name
$movefile = $vault.DocumentService.MoveFile($item.MasterId,$item.FolderId,$deletedfolder.Id)

}


$result = $null
$item = $null
#$deletefolder = New-SearchCondition -PropertyName Path -searchOperator:Contains -searchRuleType:Must -searchText "$/Wendt/ToBeDeleted"
$result1 = Find-VaultFiles -RootFolderPath "$/Wendt/ToBeDeleted"

foreach ($dfile in $result1){
try{
$getfile =$null
$namefile = $null
$folderentity = $null
$namefile = $dfile.Name
write-host "Deleting the file: $namefile"
$getfile = Get-VaultFiles -Properties @{"Name"="$namefile"}
if($getfile -ne $null){
$folderentity = $vault.DocumentService.GetFolderByPath($getfile._EntityPath)
$vault.DocumentService.DeleteFileFromFolderUnconditional($getfile.MasterId, $folderentity.Id )}}
catch{"There was an error with $namefile check the log"}
}
