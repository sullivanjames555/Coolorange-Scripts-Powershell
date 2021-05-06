########## Importing Cool Orange Modules #######
import-module powervault 
import-module powerjobs
######### ***** current user for vault session ******* #############################################
open-vaultconnection -server "Wendt-vault" -vault "Wendt" -user "coolorange" -password "nhg544FK"
#**************************************************************************************************#
$csvfile = 'C:\Vault Data\Vault Reports\Rubber Part.CSV'

$h1 = "Filename"
$h2 = "Description"
$h3 = "Material"
$h4 = "Date Modified"
$h5 = "State"
$h6 = "Created By"

$row = "" |select-Object $h1, $h2,$h3,$h4,$h5,$h6
$output = @()

$files = Get-VaultFiles -properties @{'Description' = '*Rubber*'; 'File Extension' = 'ipt'}


foreach($file in $files){
$name = $file.Name
$Description = $file.Description
$Material =$file._Material  
$Moddate = $file._ModDate 
$state = $file._State
$createdby = $file._CreateUserName

$row.$h1 = $name 
$row.$h2 = $description
$row.$h3 = $material
$row.$h4 = $Moddate
$row.$h5 = $state
$row.$h6 = $createdby


$output = $row 
$Output | Export-Csv -Append  -NoTypeInformation -Path $CSVFile -force
}