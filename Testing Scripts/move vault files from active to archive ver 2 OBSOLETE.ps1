########################################################################################################




########################################################################################################
# ~~~~~~~~ Importing of the modules necessary to run the script ~~~~~~~~~ # 

import-module powervault 
import-module "C:\Users\Public\Documents\New-cooltable.psm1"


######### current user for vault session #################################################################

open-vaultconnection -server "Wendt-vault" -vault "Wendt" -user "coolorange" -password "nhg544FK"

########################################################################################################

#~~~ This script will get values once complete from a list of files inside a goup of folders based on user input ~~~~~~
$TempDir = "C:\temp"
$file = Get-VaultFiles -FileName "970-038-4-8001-A0.iam"
$filepath = $file._FullPath
# coolorange saves a group of files locally that is necessary to load the current assembly or component. THis is the array used to get the children & Attatchments for each.
$save = Save-VaultFile -File $filepath -DownloadDirectory $TempDir
#this is going to be the main file.
$savefile = $save |select -first 1 
$savepath = $savefile._FullPath

# These are the tables that are created to store the child relations 

New-CoolTable -TableName Attatched -ColumnNames "Parent","File ID","State","ParentFullPath","Attatched","Atatched ID","Attatched Path","attatchedState","ParentLocalPath","AttatchedLocalPath"
New-CoolTable -TableName Children -ColumnNames "parent","File ID","State","ParentFullPath","Child","Child ID","Child Path","ChildState","ParentLocalPath","ChildlocalPath"

######################################################################

#region the For each loop that grabs information and stores them into the table 

foreach($sub in $save){
$localparentpath = $sub.'Full Path'
$fileid = $sub.Id
$parentname = $sub.Name
$savevaultattachments = Get-VaultFileAssociations -File $sub.'Full Path' -Attachments 

if($savevaultattachments){
foreach($attatched in $savevaultattachments){
$saveattatch = Save-VaultFile -File $attatched.'FullPath' -DownloadDirectory $TempDir
$localattatchpath = $saveattatch.LocalPath
$attatched.rows.add($parentname,$sub.id,$sub.State,$sub.'Full Path',$attatched.name,$attatched.id,$attatched.'Full Path',$attatched.State,$localparentpath,$localattatchpath)

}}
$savevaultchildren =  Get-VaultFileAssociations -File $sub.'Full Path' -Dependencies

if($savevaultchildren){

foreach($child in $savevaultchildren){
$savechild = Save-VaultFile -file $child.'Full Path' -DownloadDirectory $TempDir
$localchildpath = $savechild.LocalPath 
$children.rows.add($parentname,$sub.id,$sub.State,$sub.'Full Path',$child.name,$child.id,$child.'Full Path',$child.State,$localparentpath,$localchildpath)

}}

$fileBom = Get-VaultFileBOM -File $sub.'Full Path' -GetChildrenBy LatestVersion
$subname = $sub.name
$attatchname = $savevaultattachments.name
$childname = $savevaultchildren.Name
$bomname = $filebom.name

}
#endregion 


#region Archive Vault
#000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000
########### ending the vault connection to wendt-vault active vault #########################

$result = [Autodesk.DataManagement.Client.Framework.Vault.Library]::ConnectionManager.LogOut($global:VaultConnection)

Remove-Variable Vault -Scope Global
Remove-Variable VaultConnection -Scope Global
Remove-Variable vaultExplorerUtil -Scope Global

######000000000000000000000############ starting vault connection to wal-vaultarc ###########00000000000000000################

open-vaultconnection -server "wal-vaultarc" -vault "Wendt" -user "coolorange" -password "nhg544FK"
write-host "now putting files into archive vault"

$mainfile = add-vaultfile -From $savefile.localpath -to $savefile.'Full Path'  

# adding parents 
foreach($row in $children){
$checkCFile = Get-VaultFiles -FileName $row.Child
if(!($checkCFile)){
$addchild = Add-VaultFile -From $row.ParentLocalPath -to $row.'Child Path'
}
}

foreach($row in $children){
#creating parent 
$checkPFile = get-vaultfiles -FileName $row.Parent
if(!($checkPFile)){
$addF = add-vaultfile -From $row.ParentLocalPath -To $row.ParentFullPath
 
}
Update-VaultFile -File $addF.'FullPath' -Childs $row.'Child Path'
}








<#
foreach($item in $save){
$item.Name
$savelocalpath = $item.LocalPath
$savevaultpath =  $item.'Full Path'
$savevaultstate = $item.State
$exisitingfile = Get-VaultFiles -FileName $item.FullPath
#adding files and updating states
if(!($exisitingfile)){
$savearc = Add-VaultFile -From $savelocalpath -To $savevaultpath 
$updatefilewithchild = Update-VaultFile -File $savevaultpath -Status $savevaultstate
}
####

foreach($child in $savevaultchildren){
if($Child){
$child
$childpath = $child.Fullpath
$updatefilewithchild = Update-VaultFile -File $item.'Fullpath' -AddChilds $childpath  
}
}

foreach($attach in $attacthmentpath){
if($attatch){

$attacthmentpath = $savevaultattachments.FullPath
$updatefilewithattach = Update-VaultFile -File $item.'Fullpath' -AddAttachments $attacthmentpath 
}}
}

$result = [Autodesk.DataManagement.Client.Framework.Vault.Library]::ConnectionManager.LogOut($global:VaultConnection)

Remove-Variable Vault -Scope Global
Remove-Variable VaultConnection -Scope Global
Remove-Variable vaultExplorerUtil -Scope Global

#endregion
#>