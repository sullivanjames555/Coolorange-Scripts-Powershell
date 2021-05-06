############***##############*** needs powervault installed on machine ***##############***###############

####################### Gets installed Modules ########################

import-module powervault 
import-module powerjobs 

######### current user for vault session #################################################################
open-vaultconnection -server "wendt-vault" -vault "wendt" -user "coolorange" -password "nhg544FK"

$vfile = get-vaultfiles -properties @{"Name"="*.ipt"} 

$results = @()

foreach ($file in $vfile){

$fileass = Get-VaultFileAssociations -File $file."Full Path" -dependencies
if ($fileass.count -lt 1) {
$results += $file
$vfile.remove($file)

}}
$results1 = $results | Select-object "Full Path", "Name", "Latest Version" 

$results1 | export-csv "C:\vault data\no dependents.csv"