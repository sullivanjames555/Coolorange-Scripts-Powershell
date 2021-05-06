
$ErrorActionPreference = "Continue"
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

new-vaultconnection -server "Wendt-vault" -vault "Wendt" -user "coolorange" -password "nhg544FK"

<#
#***** importing the csv file 
$csvfolders = import-csv "C:\Vault Data\Script Input\Copy Archive Project\Archive Project.csv" -header "path","type" | select -skip 1
# sorting to show child first
$foldercsvpaths = $csvfolders | sort -Property Type
foreach($pathinfo in $foldercsvpaths.path){
#>
######### current user for vault session ##############################################################
#region Wendt-Vault 

$path = "$/Libraries"


#$path = read-host "what is the path of the folder from active vault?"


#region Creating Tables
#**************************************************************************************************
new-cooltable -TableName "TopFolder" -ColumnNames "Folder Name","Folder Category","Folder Path","FolderID"
new-cooltable -TableName "SubFolder" -ColumnNames "Folder Name","Folder Category","Folder Path","FolderID","PathCount"
new-cooltable -TableName "Files" -ColumnNames "File Name","File State","File Path","Local Path","FileClassification","FileCat","Lifecycle","Revision"
new-cooltable -TableName "ChildFiles" -ColumnNames "Parent File","Parent File Path","File Name","File Path"
new-cooltable -Tablename "CopyReport" -columnnames "File Name", "error","transfertype"
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
foreach($file1 in $getfolderfiles){
$file = $file1 |select -first 1
if($file.Name){
$filename = $file.Name
$fileid = $file.id 
$vaultfile = get-vaultfile -FileId $fileid 
$filestate = $vaultfile.State
$filefullpath = $vaultfile.'Full Path'


##### File Children ########
$vaultassociatationschild = Get-VaultFileAssociations -File $vaultfile.'Full Path' -Dependencies
if($vaultassociatationschild -ne $null){
foreach($child in $vaultassociatationschild){
$childrows = $childfiles.Rows.Add($filename,$filefullpath,$child.Name,$child.'Full Path')}}


### Saving files, Creating Table ####
$savefile = save-vaultfile -file $filefullpath -ExcludeChildren
$filelocalpath = $savefile.LocalPath
$filesrow = $files.rows.add($filename,$filestate,$filefullpath,$filelocalpath,$file.FileClass,$vaultfile._CategoryName, $vaultfile._LifeCycleDefinition, $vaultfile.Revision)
}
}
#***************************************************************************************************
#endregion files
#endregion wendt-vault



###### now exiting the current vault session to goto the archive vault #######
#region ********* archive server *************

$result = [Autodesk.DataManagement.Client.Framework.Vault.Library]::ConnectionManager.LogOut($global:VaultConnection)
New-VaultConnection -server "wal-vaultarc" -vault "Wendt" -user "coolorange" -password "nhg544FK"

$vaultserverarchive = $vaultconnection.Server

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
$suberror = "none"
if(!($checkresult)){
  if($subfile.'File Name' -like "*.dwf"){
  $filehidden = $true} else{$filehidden = $false}
  try{
$addfile = Add-VaultFile -From $subfile.'Local Path' -to $subfile.'File Path' -hidden $filehidden -FileClassification $subfile.FileClassification
}catch{$suberror = "error"}
$updatefile = Update-VaultFile -File $subfile.'File Path' -LifecycleDefinition $subfile.lifecycle -category $subfile.FileCat 
$updatefile2 = Update-VaultFile -File $subfile.'File Path' -Status $subfile.'File State' -revision $subfile.Revision
$copyreport.rows.add($subfile.'File Name', $suberror,"add Standalone Files")
clear-variable suberror 

    }
}

## now adding iam files ###
foreach($iamfile in $iamfiles){
$iamerror = "none"
try{
$addiam = Add-VaultFile -From $iamfile.'Local Path' -to $iamfile.'File Path' -FileClassification $iamfile.FileClassification
}catch{$iamerror = "error adding file"}
try{
$updateiamfilestate = Update-VaultFile -file $iamfile.'File Path' -Status $iamfile.'File State'
} catch { $iamerror = $_ }

$copyreport.rows.add($iamfile.'File Name',$iamerror,"Add Iam File")

}

foreach($childfile in $childfiles){
if($childfile.'File Name'){
try{
$updateerror = "none"
$updateall = Update-VaultFile -File $childfile.'Parent File Path' -AddChild $childfile.'File Path'
}catch{$updateerror = "error attatching $childfile.'File Name' to $childfile.'Parent File' "}
$copyreport.rows.add($childfile.'File name',$updateerror,"Add child File")

}
}


foreach($subdir in $subfoldersorted){
$subfolderfolder =$null
$subfolderfolder = $vault.DocumentService.GetFolderByPath($subdir.'Folder Path')
$archiveupdatefold = $vault.DocumentServiceExtensions.UpdateFolderCategories($subfolderfolder.id,$subdir.'Folder Category')

}
$date = get-date -Format M-d-y
$copyreport | export-csv -Path "C:\vault Data\Vault Reports\Other Reports\Copy library Archive Report $date.csv" -NoTypeInformation
$result = [Autodesk.DataManagement.Client.Framework.Vault.Library]::ConnectionManager.LogOut($global:VaultConnection)
clear-variable vaultserverarchive
}
else{write-host "Could not connect to the proper vault"
exit
}

