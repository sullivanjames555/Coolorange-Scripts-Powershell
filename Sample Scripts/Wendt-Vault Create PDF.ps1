#<####################################################################################################################################################################

<#
   ▄███████▄ ████████▄     ▄████████       ▄████████  ▄██████▄  ███▄▄▄▄    ▄█    █▄     ▄████████    ▄████████     ███        ▄████████    ▄████████ 
  ███    ███ ███   ▀███   ███    ███      ███    ███ ███    ███ ███▀▀▀██▄ ███    ███   ███    ███   ███    ███ ▀█████████▄   ███    ███   ███    ███ 
  ███    ███ ███    ███   ███    █▀       ███    █▀  ███    ███ ███   ███ ███    ███   ███    █▀    ███    ███    ▀███▀▀██   ███    █▀    ███    ███ 
  ███    ███ ███    ███  ▄███▄▄▄          ███        ███    ███ ███   ███ ███    ███  ▄███▄▄▄      ▄███▄▄▄▄██▀     ███   ▀  ▄███▄▄▄      ▄███▄▄▄▄██▀ 
▀█████████▀  ███    ███ ▀▀███▀▀▀          ███        ███    ███ ███   ███ ███    ███ ▀▀███▀▀▀     ▀▀███▀▀▀▀▀       ███     ▀▀███▀▀▀     ▀▀███▀▀▀▀▀   
  ███        ███    ███   ███             ███    █▄  ███    ███ ███   ███ ███    ███   ███    █▄  ▀███████████     ███       ███    █▄  ▀███████████ 
  ███        ███   ▄███   ███             ███    ███ ███    ███ ███   ███ ███    ███   ███    ███   ███    ███     ███       ███    ███   ███    ███ 
 ▄████▀      ████████▀    ███             ████████▀   ▀██████▀   ▀█   █▀   ▀██████▀    ██████████   ███    ███    ▄████▀     ██████████   ███    ███ 
                                                                                                    ███    ███                            ███    ███ 
#>

###################################################################################################################################################################>
                             #################### Made By James Sullivan IT Analyst ##########################
                             #          Script to convert IDW Files Released today to PDF Files              # 
                             #                                   Ver. 3.0                                    # 
                             #                              Created 7/14/2020                                #
                             #           Logs Can be located at C:\Vault Data\Vault Reports\PDF Report       #
                             #################################################################################


########## Importing Cool Orange Modules #######
import-module powervault 
import-module powerjobs
######### ***** current user for vault session ******* #############################################
open-vaultconnection -server "Wendt-vault" -vault "Wendt" -user "coolorange" -password "nhg544FK"
#**************************************************************************************************#


$date = Get-Date -Format "dddd, MMMM d, yyyy"

############ CSV Report output ##################

$CSVFile = "C:\Vault Data\Vault Reports\PDF Report\PDF $date.csv"

##########Creating a table for the report #############

$h1 = "Filename"
$h2 = "PDF File"
$h3 = "Moved File"
$h4 = "PDF Location"

$row = "" |select-Object $h1, $h2,$h3,$h4
$output = @()
########################################

$date = Get-Date -Format "dddd, MMMM d, yyyy"
$datetime = [datetime] $date

############\\ Use this switch if you would like to get files from yesterday // ###############
#$datetime = [DateTime]::Today.AddDays(-1)
#******************************************#


################____ Getting vault files _____ ##################

$idwfilemain = get-vaultfiles -properties @{'Name'='*.idw';'State'='Released*'}

####### Selecting vault files checked in today without Property Edit ##############

$idwfiles = $idwfilemain | where-object {$_."Checked In (Date Only)" -like $datetime -and $_.Comment -notcontains "Property Edit"}

######## IF $idwfiles is empty this will do nothing and output in the coolorange log that there were no idw files checked in today  ###########

if($idwfiles){

$hidePDF = $false
$workingDirectory = "C:\Temp\Wendt"


####### Starting to parse each file ###############

foreach($idw in $idwfiles){

############ Nulling the variables in the event any are empty it will not copy any dta from previous files #################

$obsoletefolder = $null
$folderitems = $null
$currentfileselect = $null
$exsistingobspath = $null
$folderfiles = $null
$idwrev = $idw.Revision
$idwfilename = [io.path]::GetFileNameWithoutExtension($idw._ClientFileName)

#########adding name to the table##############****
$row.$h1 = $idw.Name
###############################################****

$pdffilename = $idwfilename +"-"+ $idwrev + ".pdf"
$localPDFfileLocation = "$workingDirectory\$pdffilename"
$vaultPDFfileLocation = $idw._EntityPath +"/"+ (Split-Path -Leaf $localPDFfileLocation)
$fastOpen = $idw._Extension -eq "idw" -or $idw._Extension -eq "dwg" -and $idw._ReleasedRevision
$manager = $idw._Manager
$networkfilepath = "\\wal-file.wendtcorp.local\eng\PDF Files\"

#######getting existing obsolete folder ###########
 
$vaultfolder = $vault.DocumentService.GetFolderByPath
$idwfolderpath = $idw._FolderPath
$exsistingobspath = $idw._FolderPath +"/Obsolete"
$IdwpathID = $vault.DocumentService.GetFolderByPath($idwfolderpath)
$filepartnumber = $idw._PartNumber

##### Creating or getting obsolete folder from the correct path (If folder doesnt exist it creates the folder) #####################
try{
$obsoletefolder = $vault.DocumentServiceExtensions.AddFolderWithCategory("Obsolete",$IdwpathID.ID,$false,30)
Write-Host "Creating Folder $obsoletefolder._FullPath"
}
catch{ "Obsolete folder exists, No Folder Created"
$obsoletefolder = $vault.DocumentService.GetFolderByPath($exsistingobspath)
$folderitems = $vault.DocumentService.GetLatestFilesByFolderId($obsoletefolder.id,$false)
}



############## getting the correct folders and files #################


$folderfiles = $vault.DocumentService.GetLatestFilesByFolderId($idwpathID.id,$false)



### moving the current pdf file #########
$currentfileselect  = $folderfiles |where-object {$_.Name -like "$filepartnumber-??.pdf"}


################ if there are already exisitng files move them to obsolete ###############


###### Saving the pdf  #############
$idwfilename = $idw.Name
Write-Host "Starting job Create PDF for file $idwfilename .."


$downloadedFiles = Save-VaultFile -File $idw._FullPath -DownloadDirectory $workingDirectory -ExcludeChildren:$fastOpen -ExcludeLibraryContents:$fastOpen
$filedown = $downloadedFiles | select -First 1
$openResult = Open-Document -LocalFile $filedown.LocalPath -Options @{ FastOpen = $fastOpen } 

#### creating the network path or getting the network path if it doesnt exist ##############

    
     if(!(test-path -Path "$networkfilepath\$manager")){
     $networklocation = New-Item -Path $networkfilepath -name $manager -itemtype "directory" 
         write-host "Creating Folder on the network: $networklocation" 
     }
     else{
     $networklocation = "$networkfilepath\$manager\"}

############ exporting the pdf ################

if($openResult) {
    if($openResult.Application.Name -like 'Inventor*') {
        $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF_2D.ini"
    } else {
        $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF.dwg" 
    }                   
    $exportResult = Export-Document -Format 'PDF' -To $localPDFfileLocation -Options $configFile

############# End of export #################

    if($exportResult) {   
    ################## Putting the file where it needs to go ################    
        $PDFadd = Add-VaultFile -From $localPDFfileLocation -To $vaultPDFfileLocation -FileClassification None -Hidden $hidePDF 
        $filechange = Update-VaultFile -File $PDFadd._FullPath -Category $idw._CategoryName 
        $filechange2 = Update-VaultFile -File $PDFadd._FullPath -properties @{"Description"= $idw."Description";"Project01" =$idw."Project01";"Author" = $idw._Author;"Stock Number" = $idw._StockNumber ; "Manager"= $idw._Manager; "Title"= $idw._Title; "Work Order"= $idw.'Work Order'; "Revision"="$idw.Revision"} 
        $filechange3 = Update-VaultFile -File $PDFadd._FullPath -Status $idw._State
##### copy to Network #######        
        $filecopy = Copy-Item -Path $localPDFfileLocation -Destination $networklocation 
        ########### Adding more results to the Table ########
        $row.$h2 = $pdffilename
        $row.$h4 = $vaultPDFfileLocation
        ##################### Moving the file if there was an exisitng file .. only if the file gets created #########################
        $nameofile = $idw.Name
        if($currentfileselect.Name -match $filechange.Name){
        write-host "Did not move any files related to ($nameofile)"
        
        ##### new line need to test to ensure it still doesnt try to move fiels that dont exist ############
        if(!($currentfileselect.Name)){ write-host "No exisiting Files"
         }$row.$h3 = "No Pre Exisiting PDF"
      
         }

        #####################  ###################################
        else{
            $currentfilepath = $idw.Path + "/" + $currentfileselect.Name
            $current = get-vaultfiles -properties @{'Name'= $currentfileselect.Name}
            $pdfchangestate = Update-VaultFile -File $current._FullPath -status "Obsolete"
            $pdfsource = $vault.DocumentService.GetFolderByPath($current.Path)
            $movefile = $vault.DocumentService.MoveFile($current.MasterID,$pdfsource.Id, $obsoletefolder.Id)
                write-host "Moved $nameofile into the Obsolete Folder"
                $row.$h3 = $currentfileselect.Name
                }
                $output = $row 
                $Output | Export-Csv -Append  -NoTypeInformation -Path $CSVFile -force

                
                }

                ########################## END OF LOOP ##################################

else{ write-host "No old pdf File Found to move"}
   }
     }
$closeResult = Close-Document

Write-Host "Completed job 'Create PDF '"
}
else { write-host "There are no IDW Files Checked in today"}

