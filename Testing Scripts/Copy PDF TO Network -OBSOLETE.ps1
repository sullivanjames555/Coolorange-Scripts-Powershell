########## The intention of this script is to batch grab pdf files and move them to the network ############## 




import-module powervault 
import-module powerjobs
Import-Module "C:\ProgramData\coolOrange\powerJobs\Modules\VaultSearch.psm1"
######### ***** current user for vault session ******* #############################################
open-vaultconnection -server "Wendt-vault" -vault "Wendt" -user "coolorange" -password "nhg544FK"

$networkfilepath = "\\wal-file.wendtcorp.local\eng\PDF Files"


$pdf = New-SearchCondition -PropertyName "File Extension" -searchOperator:Contains -searchRuleType:Must -searchText "pdf"
$released = New-SearchCondition -PropertyName "State" -searchOperator:Contains -searchRuleType:Must -searchText "Released"
$state = New-SearchCondition -PropertyName State -searchOperator:NotEmpty
$srchCond_managernotempty = New-SearchCondition -PropertyName Manager -searchOperator:NotEmpty
$srchCond_titlenotempty = New-SearchCondition -PropertyName Title -searchOperator:NotEmpty

$results = Find-VaultFiles -RecurseFolders -SearchConditions @($pdf,$released,$state,$srchCond_managernotempty,$srchCond_titlenotempty) -LatestFilesOnly


foreach($file in $results){
################ testing the path to ensure the folder exists ############## 
$name = $file.name
$files  = Get-VaultFiles -FileName $name
$manager = $files."Manager"

 if(!(test-path -Path "$networkfilepath\$manager")){
     $networklocation = New-Item -Path $networkfilepath -name $manager -itemtype "directory" 
         write-host "Creating Folder on the network: $networklocation" 
     }
     else{
     $networklocation = "$networkfilepath\$manager\"}

   #$child =  Get-ChildItem -Path $networklocation
 

   $savefile = Save-VaultFile -File $files._FullPath -DownloadDirectory $networklocation


     }### end of foreach $file