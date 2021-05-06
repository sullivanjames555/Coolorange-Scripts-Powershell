######## Powershell Coolorange Script To create Files and folders for Structural folders inside of a subfolder #########
#### Created By James Sullivan ######
##using coolorange powershell and vault API Calls
#created 10-21-2020 Rev 1

### folder structure to Create ######

#Top Level Folder (project)
#     Sub folder 
#         obsolete(dir)
#         File1.zip
#         File2.zip

$ErrorActionPreference = "SilentlyContinue"
######import modules ######
import-module powervault 
Add-Type -AssemblyName Microsoft.VisualBasic
######### current user for vault session #################################################################

open-vaultconnection -server "Wendt-vault" -vault "Wendt" -user "coolorange" -password "nhg544FK"


#region Input *********************************************************************************************************************************************
    ########Inputs##########
$folder =           [Microsoft.VisualBasic.Interaction]::InputBox('Please Enter The Structural Name', 'Structural Name', "(EX..5500-906-4-9002)")
$description =      [Microsoft.VisualBasic.Interaction]::InputBox('Please Enter The Description', 'Structural Description', "(EX..Railings and Rails")
$foldersplit = $folder.SubString($folder.Length-4)
$topfolder = $folder.SubString(0,8)
$foldername = $foldersplit +" " + "-" +" " +  $description

#endregion input ************************************************************************************************************************************************


################# setting up naming variables ##########

$files =            get-vaultfiles -properties @{"Folder Path" = "$/Wendt/M$topfolder*";"Name" = "*i*"} 

$firstfile =        $files |select -first 1 |where-object{ $_.FullPath -notlike "*Obsolete"}

$fpath =            $firstfile._FolderPath

$toplevelp =     split-path $fpath -Parent

$toplevelpath = $toplevelp.Replace("\","/")

$toplevelpathID = $vault.DocumentService.GetFolderByPath($toplevelpath)
$zipfiles = @()


#region get folder ********************************************************************************************************

## This section will try to create the folder and pass true or false #

try{
$subfolderpath = $toplevelpathID.FullName + "/" + $foldername
$subf = $vault.DocumentService.GetFolderByPath($subfolderpath)
"This Folder Exists We will not Create This Folder "
$vaultfolder = $true
}
Catch{
"Folder Doesnt exist we will create it"
$vaultfolder = $false
}
#endregion get folder *******************************************************************************************************


    if($vaultfolder -eq $false){
    #region Start Create folder ********************************************************************************************
#!! quick Note:
#!! this section will create the Appropriate folder 
#!! if it doesnt exist based on the try and catch folder region
#

####### ------ you can now use this as the folder path to put the file in ------ #######
    $subf = $vault.DocumentServiceExtensions.AddFolderWithCategory("$foldername",$toplevelpathID.Id,$false,22)}

######### creating the OBSOLETE folder into the now created folder ###########

    $obsolete = $vault.DocumentServiceExtensions.AddFolderWithCategory("Obsolete",$subf.ID,$false,30)

    ## ****  for not ading folders with structural category **** ###
    #$subfolder = $vault.DocumentService.AddFolder($foldername, $toplevelpathID.Id, $false)



#endregion Create Folder ***********************************************************************************************




###########################################


### getting files from vault ###
$templatezip = get-vaultfiles -folder "$/Wendt/Engineering/Start Parts/5500 - Structural"

foreach($zipfile in $templatezip){
$zippath = $zipfile._FullPath
##saving files locally #####
$savedir = save-vaultfile -File $zippath -DownloadDirectory "C:\TEMP\Structural"


########### getting the files to put into vault #############
$zipfiles += get-childitem -path $savedir.LocalPath
}

##start region Zipfiles
foreach($zip in $zipfiles){
 # creating the variables for the foreach loop and name 
$localtoplevelpath = "C:\TEMP\Structural\"
$localfilename = $folder + " - " + $zip.Name
$localpath = $localtoplevelpath + $zip.Name
$vaultpath = $subf.FullName + "/" + $localfilename
 # end 
 $checkfile =""
if($vaultfolder -eq $true){


$checkfile = get-vaultfiles -FileName $localfilename -folder $subf.FullName
if(!$checkfile){ 
write-host "putting zip file $zip.name in folder"

$add = Add-VaultFile -From $localpath -To $vaultpath
}
else{ write-host "the zip folder exists, not adding $zip.Name"}
#clean-up -folder "C:\TEMP\Structural"
$add = Add-VaultFile -From $localpath -To $vaultpath

#### end region zipfiles
}
}


########


#### scripters notes 
### get the type and id of an existing folder the easiest way ###
## $testing = $vault.DocumentService.GetFolderByPath("$/Wendt/M5500-096 - General Iron NFe - 206955")
# $testing.Cat will give you ID and ID Name 
