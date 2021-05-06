#region functions
######### Function #########
function Add-VaultFolderandfiles {
[CmdletBinding()]
	param(
$totals, #this is how many items need to be made for the priority
$name, #this is the higher level folder
[int]$id, #id of folder this folder goes into 
[string]$file, #the beginning path of the file 
$prioritynum, # the priority number 
$standardnum, #D85xx - Guards
$beginningnum, #the beginiing of the number 
$name2,
[int]$id1, #id of folder this goes into
[string]$file1,
$standardnum1,
$beginnningnum2,
$categoryid #category of the folder to add Obsolete is 30

)

for ($i = 1; $i -le $totals; $i++) { 
$p1c = "{0:D2}" -f $i
$var = $name + $p1c + "-"+$prioritynum
$var2 = $name2 + $p1c + "-"+$prioritynum
$mainfolder = $vault.DocumentService.AddFolder($var,$id,$false)
$vault.DocumentServiceExtensions.AddFolderWithCategory("Obsolete",$mainfolder.ID,$false,$categoryid)

$secondaryfolder = $vault.DocumentService.AddFolder($var,$id1,$false)
$vault.DocumentServiceExtensions.AddFolderWithCategory("Obsolete",$secondaryfolder.ID,$false,$categoryid)
$filepath1 = $file + "/"+ $var + "/"+$standardnum +"-4-"+$beginnningnum+$p1c+"-"+$prioritynum+ "-dummy-A0.ipt"
$filepath2 = $file1 + "/"+ $var + "/"+$standardnum1 +"-4-"+$beginnningnum2+$p1c+"-"+$prioritynum+ "-dummy-A0.ipt"
$conconfile = add-vaultfile -from $partpath -to $filepath1
$updateconfile = Update-vaultfile -file $conconfile._fullpath -properties $hashtable2
$plantconfile = add-vaultfile -from $partpath -to $filepath2
$updateconfile2 = Update-vaultfile -file $plantconfile._fullpath -properties $hashtable2
}}

#endregion





Foreach($prioriti in $priorities) {
set-variable priority -value $prioriti
$hashtable2 = @{"_Manager" = "$salesnum"; "_Title" = "$prioriti"}
$pconveyorinput = read-host -prompt "how many $priority $itemtype need to be made?:"
if (!$pconveyorinput) { Write-Host "variable is null" }
else{
Add-VaultFolderandfiles -totals $pconveyorinput -name "D09" -prioritynum $priority -id $conveyorpath1.Id -file $conveyorpath1.FullName -standardnum $conveyors -beginningnum "51" -name2 "D60" -id1 $conveyorpath2.Id -file1 $conveyorpath2.FullName -standardnum1 $jobnumber -beginnningnum2 "60" -categoryid 30 


for ($i = 1; $i -le $pconveyorinput; $i++) { 
$p1c = "{0:D2}" -f $i
$var = "D09$p1c-$priority"
$conroot = $vault.DocumentService.AddFolder($var, $conveyorpath1.Id, $false)
$conroot2 = $vault.DocumentService.AddFolder($var, $plantpath.Id, $false)
$obsoletefolder1 = $vault.DocumentServiceExtensions.AddFolderWithCategory("Obsolete",$conroot.ID,$false,30)
$obsoletefolder2 = $vault.DocumentServiceExtensions.AddFolderWithCategory("Obsolete",$conroot2.ID,$false,30)

$conchute = add-vaultfile -from $partpath -to "$/$conveyorname/D09xx - Chutes/D09$p1c-$priority/$conveyors-4-09$p1c-$priority-dummy-A0.ipt"
$updatechute = Update-vaultfile -file $conchute._fullpath -properties $hashtable2
$plantchute = add-vaultfile -from $partpath -to "$/$plant/D09xx - Chutes/D09$p1c-$priority/$jobnumber-4-09$p1c-$priority-dummy-A0.ipt"
$updatechute2 = Update-vaultfile -file $plantchute._fullpath -properties $hashtable2
}}}