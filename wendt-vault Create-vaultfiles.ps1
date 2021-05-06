import-module powervault 
Add-Type -AssemblyName Microsoft.VisualBasic
######### current user for vault session #################################################################

open-vaultconnection -server "Wendt-vault" -vault "Wendt" -user "coolorange" -password "nhg544FK"
<#
$in = 5

$filepath = "M907-126 - Upstate - 209771"
$title = "P1"
$manager = "209771"
#>

######### Function #########
function create-vaultfiles{
[CmdletBinding()]
param(

[string]$count, 
[string]$filepath, 
[string]$title,
[string]$manager

)
$filename0 = $filepath.SubString(1,8)
$filename1 = $filename0+"-4-XXXX-dummy"
$filename2 = $filename1 -replace (' ')
$pathname = "$/Wendt/$filepath"
$hashtable = @{"Title" = $title ; "Manager" = $manager}
$vaultfolder = $vault.DocumentService.GetFolderByPath($pathname)

$partpath = "C:\Vault Data\Dummy Part\Dummy_Part.ipt" 
$path = $vaultfolder.Fullname

for ($i = 1; $i -le $count; $i++) { 
$p1c = "{0:D2}" -f $i

 
$filename = $path + "/"+ $filename2 + $p1c + "-A0.ipt"
write-host $filename 

$file1 = add-vaultfile -from $partpath -to "$filename" 
$file2 = Update-vaultfile -file $file1._FullPath -properties $hashtable
$filenameout = $file.Name
write-host "$filenameout has been added"
}
}

$fileinputpath = [Microsoft.VisualBasic.Interaction]::InputBox('Please Enter The Full Folder Name', 'Folderpath', "(EX..M907-126 - Upstate - 209771)")
$filecount = [Microsoft.VisualBasic.Interaction]::InputBox('Please Enter The total number of parts you would like to create.', 'Part Count', "(EX..50)")
$managerinput = [Microsoft.VisualBasic.Interaction]::InputBox('Please Enter The File Manager', 'Manager', "(EX..209771)")
$titleinput = [Microsoft.VisualBasic.Interaction]::InputBox('Please Enter The File Title', 'Title', "(EX..P1)")

create-vaultfiles -count $filecount -filepath $fileinputpath -manager $managerinput -title $titleinput 