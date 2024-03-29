##########################################################################################################
#        This script is designed to build a project within vault and adds files into the subfolders      #
#                              *** Currently only for sandbox vault ***                                  #
#                                                                                                        #
#                                                                                                        #
#                                                                                                        #
#                     *****    This needs to be ran by a vault administrator    *****                    #
#       #error logs for the script can be found in %userprofile%\appdata\local\coolorange\powervault     #
#                   Log for things done found in %USERPROFILE%\Documents\output$date.txt"                #
#                                  Created by James Sullivan 7-30-19 rev 4                               #
##########################################################################################################
##################################needs powervault installed on machine ##################################
import-module powervault 

######### current user for vault session #################################################################
##********* when single sign on is enabled for vault, will need to switch to using local credentials
open-vaultconnection -server "Wendt-vault" -vault "Wendt" -user "coolorange" -password "nhg544FK"

$date = get-date -format "M-dd-yyyy"
Start-Transcript -path "$env:USERPROFILE\Documents\Vault Project output-$date.txt" -force

################################## sales and order variables ############################################################
$salesorder = read-host -prompt 'What is the customer name and salesorder number? (ex.. turnkey - 3335)'
$jobnumber = read-host -prompt 'what is the Plant folder plus job number? (ex... 930-12345)'
$plant = "M$jobnumber - $salesorder"
$salesnum = $salesorder.split('-')[1].split(',')[0] -replace " "


################################# the folder & priority arrays ##########################################################
#****** subfolders for plant layout main folder
$subfolders = "D0100 - Layout", "D0101 - Footprint", "D01xx - Equipment","D85xx - Guards", "D60xx - Conveyors", "D90xx - Platforms", "D09xx - Chutes" , "D77xx - Railings"
#****** all the priorities to be used 
$priorities = "P1","P2","P3","P4","P5","P6","P7","P8","P9"
$priority = "P1"
#################################### conveyor variables #################################################################
$conveyors = read-host -prompt "what is the Conveyor Number? :"
$conveyorname = "M$conveyors - $salesorder"
$conveyorfold = "D09xx - Chutes", "D51xx - Conveyors","D77xx - Supports", "D85xx - Guards"

#################################### vault variables ####################################################################
$vaultfolder = $vault.DocumentService.GetFolderByPath
$Vroot = $vault.DocumentService.GetFolderByPath("$")

# adding the top level Conveyor Folder
$conveyorroot = $vault.DocumentService.AddFolder($Conveyorname, $Vroot.Id, $false) 
# adding top level Plant Folder
$plantroot = $vault.DocumentService.AddFolder($plant, $Vroot.Id, $false)


######################## building subfolders for conveyor folder and main plant folders ##################################
#Creating subfolders for each item in Conveyorfold
foreach($conveyor in $conveyorfold){$vault.DocumentService.AddFolder($conveyor, $conveyorroot.Id, $false) }
#Creating subfolders for each item in Subfolders
foreach($subfolder in $subfolders){$vault.DocumentService.AddFolder($subfolder, $plantroot.id, $false) }
#grabbing a part from the network to transfer later
$partpath = "C:\Vault Data\Dummy Part\Dummy_Part.ipt" 

########### These are vault Properties that are updated on each part when put into vault ############################
$hashtable1 = @{"_Manager" = "$salesnum";"_Title" = "P1"}


############################# putting individual files in empty folders ##################################################
###*** the first command puts a file in a folder, the second command updates that file to have the appropriate properties
## *** these did not work with pipelines to each other *** ####
$file1 = add-vaultfile -from $partpath -to "$/$plant/D85xx - Guards/$conveyors-4-D85xx - Guards-dummy-A0.ipt" 
Update-vaultfile -file $file1._fullpath -properties $hashtable1
$file2 = add-vaultfile -from $partpath -to "$/$conveyorname/D85xx - Guards/$conveyors-4-D85xx - Guards-dummy-A0.ipt" 
Update-vaultfile -file $file2._FullPath -properties $hashtable1
$file3 = add-vaultfile -from $partpath -to "$/$plant/D0100 - Layout/$jobnumber-4-D0100 - Layout-dummy-A0.ipt"
Update-vaultfile -file $file3._fullpath -properties $hashtable1
$file4 = add-vaultfile -from $partpath -to "$/$plant/D0101 - Footprint/$jobnumber-4-D0101 - Footprint-dummy-A0.ipt"
Update-vaultfile -file $file4._fullpath -properties $hashtable1 
$file5 = add-vaultfile -from $partpath -to "$/$plant/D01xx - Equipment/$jobnumber-4-D01xx - Equipment-dummy-A0.ipt" 
Update-vaultfile -file $file5._fullpath -properties $hashtable1


############################ Building foreach statement and folders for conveyors  #######################################
$itemtype = "conveyors"
$conveyorpath1 = $vault.DocumentService.GetFolderByPath("$/$conveyorname/D51xx - Conveyors")
$conveyorpath2 = $vault.DocumentService.GetFolderByPath("$/$plant/D60xx - Conveyors")
Foreach($prioriti in $priorities) {
set-variable priority -value $prioriti
$hashtable2 = @{"_Manager" = "$salesnum"; "_Title" = "$prioriti"}
$pconveyorinput = read-host -prompt "how many $priority $itemtype need to be made?:"
if (!$pconveyorinput) { Write-Host "variable is null" }
else{
for ($i = 1; $i -le $pconveyorinput; $i++) { 
$p1c = "{0:D2}" -f $i
$var = "D51$p1c-$priority"
$conroot = $vault.DocumentService.AddFolder($var, $conveyorpath1.Id, $false)
$conroot2 = $vault.DocumentService.AddFolder("D60$p1c-$priority", $conveyorpath2.Id, $false)

$conconfile = add-vaultfile -from $partpath -to "$/$conveyorname/D51xx - Conveyors/$var/$conveyors-4-51$p1c-$priority-dummy-A0.ipt" 
$updateconfile = Update-vaultfile -file $conconfile._fullpath -properties $hashtable2
$plantconfile = add-vaultfile -from $partpath -to "$/$plant/D60xx - Conveyors/D60$p1c-$priority/$jobnumber-4-60$p1c-$priority-dummy-A0.ipt"
$updateconfile2 = Update-vaultfile -file $plantconfile._fullpath -properties $hashtable2

}}}


############################ Building foreach statement and folders for Chutes  #########################################
$itemtype = "Chutes"
$conveyorpath1 = $vault.DocumentService.GetFolderByPath("$/$conveyorname/D09xx - Chutes")
$plantpath = $vault.DocumentService.GetFolderByPath("$/$plant/D09xx - Chutes")
Foreach($prioriti in $priorities) {
set-variable priority -value $prioriti
$hashtable2 = @{"_Manager" = "$salesnum"; "_Title" = "$prioriti"}
$pconveyorinput = read-host -prompt "how many $priority $itemtype need to be made?:"
if (!$pconveyorinput) { Write-Host "variable is null" }
else{
for ($i = 1; $i -le $pconveyorinput; $i++) { 
$p1c = "{0:D2}" -f $i
$var = "D09$p1c-$priority"
$conroot = $vault.DocumentService.AddFolder($var, $conveyorpath1.Id, $false)
$conroot2 = $vault.DocumentService.AddFolder($var, $plantpath.Id, $false)

$conchute = add-vaultfile -from $partpath -to "$/$conveyorname/D09xx - Chutes/D09$p1c-$priority/$conveyors-4-09$p1c-$priority-dummy-A0.ipt"
$updatechute = Update-vaultfile -file $conchute._fullpath -properties $hashtable2
$plantchute = add-vaultfile -from $partpath -to "$/$plant/D09xx - Chutes/D09$p1c-$priority/$jobnumber-4-09$p1c-$priority-dummy-A0.ipt"
$updatechute2 = Update-vaultfile -file $plantchute._fullpath -properties $hashtable2
}}}



############################ Building foreach statement and folders for Supports  ######################################

$itemtype = "Supports"
$conveyorpath1 = $vault.DocumentService.GetFolderByPath("$/$conveyorname/D77xx - Supports")

Foreach($prioriti in $priorities) {
set-variable priority -value $prioriti
$hashtable2 = @{"_Manager" = "$salesnum"; "_Title" = "$prioriti"}
$pconveyorinput = read-host -prompt "how many $priority $itemtype need to be made?:"
if (!$pconveyorinput) { Write-Host "variable is null" }
else{
for ($i = 1; $i -le $pconveyorinput; $i++) { 
$p1c = "{0:D2}" -f $i
$var = "D77$p1c-$priority"
$conroot = $vault.DocumentService.AddFolder($var, $conveyorpath1.Id, $false)


$consupport = add-vaultfile -from $partpath -to "$/$conveyorname/D77xx - Supports/$var/$conveyors-4-77$p1c-$priority-dummy-A0.ipt"
$update1 = Update-vaultfile -file $consupport._fullpath -properties $hashtable2
}}}


############################ Building foreach statement and folders for Platforms  ######################################
$itemtype = "Platforms & Railings"
$conveyorpath1 = $vault.DocumentService.GetFolderByPath("$/$plant/D90xx - Platforms")
$plantpath = $vault.DocumentService.GetFolderByPath("$/$plant/D77xx - Railings")
Foreach($prioriti in $priorities) {
set-variable priority -value $prioriti
$hashtable2 = @{"_Manager" = "$salesnum"; "_Title" = "$prioriti"}
$pconveyorinput = read-host -prompt "how many $priority $itemtype need to be made?:"
if (!$pconveyorinput) { Write-Host "variable is null" }
else{
for ($i = 1; $i -le $pconveyorinput; $i++) { 
$p1c = "{0:D2}" -f $i
$var = "D90$p1c-$priority"
$var2 = "D77$p1c-$priority"
$conroot = $vault.DocumentService.AddFolder($var, $conveyorpath1.Id, $false)
$plantroot = $vault.DocumentService.AddFolder($var2, $plantpath.Id, $false)

$plantplat = add-vaultfile -from $partpath -to "$/$plant/D90xx - Platforms/$var/$jobnumber-4-90$p1c-$priority-dummy-A0.ipt"
Update-vaultfile -file $plantplat._fullpath -properties $hashtable2
$plantrail = add-vaultfile -from $partpath -to "$/$plant/D77xx - Railings/$var2/$jobnumber-4-77$p1c-$priority-dummy-A0.ipt"
Update-vaultfile -file $plantrail._fullpath -properties $hashtable2
}}}


############################# End of script Below will open the transcript on what happened ############################
stop-transcript

