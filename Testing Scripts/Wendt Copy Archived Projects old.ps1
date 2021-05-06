
#region Sources
import-module powervault 
Add-Type -AssemblyName Microsoft.VisualBasic
import-module "C:\Users\Public\Documents\New-cooltable.psm1"
#endregion

#region $$$$$$$$ Function $$$$$$$$$ ###

function New-VaultConnection {
[CmdletBinding()]
param(
[Parameter(Mandatory=$true)]
[String]$server,
[Parameter(Mandatory=$true)]
$vault,
[Parameter(Mandatory=$true)]
$user,
[Parameter(Mandatory=$true)]
$password

)
$output =@()
if(!($vaultconnection)){
try{
open-vaultconnection -server $server -vault $vault -user $user -password $password
}catch{
$output = "Unable To Connect to vault;$error"}
if($vaultconnection){
$vaultname = $vaultconnection.Server
$output = "Now connected to $vaultname"
}
}
else{

$result = [Autodesk.DataManagement.Client.Framework.Vault.Library]::ConnectionManager.LogOut($global:VaultConnection)
Remove-Variable Vault -Scope Global
Remove-Variable VaultConnection -Scope Global
Remove-Variable vaultExplorerUtil -Scope Global


try{
open-vaultconnection -server $server -vault $vault -user $user -password $password
}catch{$errorconnect = "there was an error connecting to $server"}
if($vaultconnection){
$vaultname = $vaultconnection.server
$vaultuser = $vaultconnection.UserName
$output = "$vaultuser is now connected to $vaultname"
}
else{
$output = "Unable To Connect to vault;$errorconnect"
}
 
 
}
$output | out-host
}

#endregion 
#**************************************

#***** importing the csv file 
$csvfolders = import-csv "C:\Vault Data\Script Input\Copy Archive Project\Archive Project.csv" -header "path","type" | select -skip 1
# sorting to show child first
$foldercsvpaths = $csvfolders | sort -Property Type
foreach($pathinfo in $foldercsvpaths.path){
new-vaultconnection -server "Wendt-vault" -vault "Wendt" -user "coolorange" -password "nhg544FK"
$path = "$/Wendt/$pathinfo"
######### current user for vault session ##############################################################
#region Wendt-Vault 




#$path = read-host "what is the path of the folder from active vault?"


#region Creating Tables
#**************************************************************************************************
new-cooltable -TableName "TopFolder" -ColumnNames "Folder Name","Folder Category","Folder Path","FolderID"
new-cooltable -TableName "SubFolder" -ColumnNames "Folder Name","Folder Category","Folder Path","FolderID","PathCount"
new-cooltable -TableName "Files" -ColumnNames "File Name","File State","File Path","Local Path","FileClassification","FileCat","Lifecycle","Revision"
new-cooltable -TableName "ChildFiles" -ColumnNames "Parent File","Parent File Path","File Name","File Path"
new-cooltable -TableName "AttatchedFiles" -ColumnNames "Parent File","Parent File Path","File Name","File Path"
#**************************************************************************************************
#endregion

#region top folder variables
#**************************************************************************************************
$wendt = $vault.DocumentService.GetFolderByPath("$/Wendt")
$toppath = $vault.DocumentService.GetFolderByPath($path)
$topfoldercat = $toppath.Cat.CatId
$topfoldername = $toppath.Name
$topfolderpath = $toppath.FullName
$subfolderpaths = $vault.DocumentService.GetFolderIdsByParentIds($toppath.id,$true)
$TopFolder.rows.add($topfoldername,$topfoldercat,$topfolderpath,$toppath.Id)
$getfolderfiles = @()
#**************************************************************************************************
#endregion
write-host "getting files for $topfoldername"
#region Subfolders for top folder 
#**************************************************************************************************
foreach($folder in $subfolderpaths){

$individualfolder = $vault.DocumentService.GetFolderById($folder)
$subfolderindname = $individualfolder.Name
$subfolderindpath = $individualfolder.FullName
$Subfoldercat = $individualfolder.Cat.CatId
$folderpathcount = ($individualfolder.FullName).split("/")
$subfolderrows = $subfolder.rows.add($subfolderindname,$subfoldercat,$subfolderindpath,$folder,$folderpathcount.count)
$getfolderfiles += $vault.DocumentService.GetLatestFilesByFolderId($individualfolder.id,$true)
}
#**************************************************************************************************
#endregion subfolders 

#region Getting Files and creating a table 
#***************************************************************************************************
foreach($file in $getfolderfiles){
if($file.Name){
$filename = $file.Name
$fileid = $file.id 
$vaultfile = get-vaultfiles -FileName $filename
$filestate = $vaultfile.State
$filefullpath = $vaultfile.'Full Path'

##### File Children ########
$vaultassociatationschild = Get-VaultFileAssociations -File $vaultfile.'Full Path' -Dependencies
if($vaultassociatationschild -ne $null){
foreach($child in $vaultassociatationschild){
$childrows = $childfiles.Rows.Add($filename,$filefullpath,$child.Name,$child.'Full Path')}}

#### File Attatchments ####
$vaultassociatationsattatched = Get-VaultFileAssociations -File $vaultfile.'Full Path' -Attachments
if($vaultassociatationsattatched -ne $null){
foreach($attatchment in $vaultassociatationsattatched){
$attatchedrows = $attatchedFiles.Rows.Add($filename,$filefullpath,$attatchment.Name,$attatchment.FullPath)}}

### Saving files, Creating Table ####
$savefile = save-vaultfile -file $filefullpath -ExcludeChildren
$filelocalpath = $savefile.LocalPath
$filesrow = $files.rows.add($filename,$filestate,$filefullpath,$filelocalpath,$file.FileClass,$vaultfile._CategoryName, $vaultfile._LifeCycleDefinition, $vaultfile.Revision)
}
}
#***************************************************************************************************
#endregion files
#endregion wendt-vault

$vaultconnection = $vaultconnection.Server
write-host "now leaving $vaultconnection "

###### now exiting the current vault session to goto the archive vault #######
#region ********* archive server *************

$result = [Autodesk.DataManagement.Client.Framework.Vault.Library]::ConnectionManager.LogOut($global:VaultConnection)
$newvault = New-VaultConnection -server "wal-vaultarc" -vault "Wendt" -user "coolorange" -password "nhg544FK"
$newvault
Open-VaultConnection -server "wal-vaultarc" -vault "Wendt" -user "coolorange" -password "nhg544FK"
$vaultserverarchive = $vaultconnection.Server
start-sleep -Seconds 10
if($vaultserverarchive -like "wal-vaultarc"){
write-host "now connected to wal-vaultarc"
$topfolderadd = $vault.DocumentServiceExtensions.AddFolderWithCategory($topfoldername,$wendt.Id,$false,$topfoldercat)
$mainfolderid = $topfolderadd.Id
$topfolderpath = $topfolderadd.FullName
$subfoldersorted = $subfolder | sort "PathCount"
#sorting the ipt files to add them first #
$subfiles = $files | Where-Object {$_.'File Name' -notlike "*.iam"}
$iamfiles = $files | where-object {$_.'File Name' -like "*.iam"}


### getting the files and adding them this is ipt and everything else excluding iam files ###
foreach($subfile in $subfiles){
$checkresult = Get-VaultFile -File $subfile.Name
if(!($checkresult)){
  if($subfile.'File Name' -like "*.dwf"){
  $filehidden = $true} else{$filehidden = $false}
$addfile = Add-VaultFile -From $subfile.'Local Path' -to $subfile.'File Path' -hidden $filehidden -FileClassification $subfile.FileClassification
$updatefile = Update-VaultFile -File $subfile.'File Path' -LifecycleDefinition $subfile.lifecycle -category $subfile.FileCat 
$updatefile2 = Update-VaultFile -File $subfile.'File Path' -Status $subfile.'File State' -revision $subfile.Revision

    }
}

## now adding iam files ###
foreach($iamfile in $iamfiles){
$addiam = Add-VaultFile -From $iamfile.'Local Path' -to $iamfile.'File Path' -FileClassification $iamfile.FileClassification
$updateiamfilestate = Update-VaultFile -file $iamfile.'File Path' -Status $iamfile.'File State'
}

foreach($childfile in $childfiles){
$updateall = Update-VaultFile -File $childfile.'Parent File Path' -AddChild $childfile.'File Path'

}
foreach($AttatchedFile in $AttatchedFiles){
$updateall = Update-VaultFile -File $attatchedfile.'Parent File Path' -AddChild $attatchedfile.'File Path'

}

foreach($subdir in $subfoldersorted){
$subfolderfolder =$null
$subfolderfolder = $vault.DocumentService.GetFolderByPath($subdir.'Folder Path')
$archiveupdatefold = $vault.DocumentServiceExtensions.UpdateFolderCategories($subfolderfolder.id,$subdir.'Folder Category')

}


$result = [Autodesk.DataManagement.Client.Framework.Vault.Library]::ConnectionManager.LogOut($global:VaultConnection)
}
else{write-host "Could not connect to the proper vault"
exit
}
}
#endregion 