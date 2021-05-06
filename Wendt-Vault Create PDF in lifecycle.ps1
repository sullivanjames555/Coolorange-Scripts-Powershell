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
                             #                                   Ver. 3.5                                    # 
                             #                              Created 7/17/2020                                #
                             #           Logs Can be located at C:\Vault Data\Vault Reports\PDF Report       #
                             #################################################################################
############## building the report #############



####################################################
write-host "Starting Job Create PDF On Lifecycle"
########## Importing Cool Orange Modules #######
import-module powervault 
import-module powerjobs
######### ***** current user for vault session ******* #############################################
open-vaultconnection -server "Wendt-vault" -vault "Wendt" -user "coolorange" -password "nhg544FK"
#**************************************************************************************************#
$ErrorActionPreference = "SilentlyContinue"


# for testing purposes you can remove this # symbol ######

#$file = Get-VaultFile -Properties @{'Name'="930-060-4-6009.idw"} #Search for a Drawing file"910-055-4-7002.idw"
$thename = $file.name

if($file.name -NotLike "*.idw"){ 

write-host "Not an idw file"
$ErrorActionPreference = "SilentlyContinue"
}
else{
try{
write-host "starting conversion for $thename"
$timerstart = get-date 
$timereportcsv = "C:\Vault Data\Vault Reports\Other Reports\CoolorangeJobs.csv"
$pdfcount = @()

$ht1 = "Script Name"
$ht2 = "Start time"
$ht3 = "End Time"
$ht4 = 'Count'
$ht5 = 'Total Time'
$timerow = "" |select-Object $ht1, $ht2,$ht3,$ht4,$ht5
$timeout = @()

$timerow.$ht1 = "Wendt- Vault Create PDF"
$timerow.$ht2 = $timerstart



###### 
$times = Measure-command {
  
$date = Get-Date -Format "dddd, MMMM d, yyyy"

############ CSV Report output ##################

$CSVFile = "C:\Vault Data\Vault Reports\PDF Report\PDF $date.csv"

##########Creating a table for the report #############

$hp1 = "Filename"
$hp2 = "PDF File"
$hp3 = "Moved File"
$hp4 = "PDF Location"

$prow = "" |select-Object $hp1, $hp2,$hp3,$hp4
$output = @()
########################################

$date1 = Get-Date 
$date = $date1.AddHours(-1)
$datetime = [datetime] $date


## variables ##
$hidePDF = $false
$workingDirectory = "C:\Temp\$($file._Name)"
$idwfilename = [io.path]::GetFileNameWithoutExtension($file._Name)
$idwrev = $file._Revision
$pdffilename = $idwfilename +"-"+ $idwrev + ".pdf" 
$exisitingfile = Get-VaultFiles -FileName $pdffilename
if(!($exisitingfile)){
$localPDFfileLocation = "$workingDirectory\$pdffilename"
$vaultPDFfileLocation = $file._EntityPath +"/"+ (Split-Path -Leaf $localPDFfileLocation)
$fastOpen = $file._Extension -eq "idw" -or $file._Extension -eq "dwg" -and $file._ReleasedRevision
$manager = $file._Manager
$networkfilepath = "\\wal-file.wendtcorp.local\eng\PDF Files"


#######getting existing obsolete folder ###########
 
$vaultfolder = $vault.DocumentService.GetFolderByPath
$idwfolderpath = $file._FolderPath
$exsistingobspath = $file._FolderPath +"/Obsolete"
$IdwpathID = $vault.DocumentService.GetFolderByPath($idwfolderpath)
$filepartnumber = $file._PartNumber

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


###### Saving the pdf  #############
$idwfilename = $file.Name
Write-Host "Starting job Create PDF for file $idwfilename .."
$downloadedFiles = Save-VaultFile -File $file._FullPath -DownloadDirectory $workingDirectory -ExcludeChildren:$fastOpen -ExcludeLibraryContents:$fastOpen
$filedown = $downloadedFiles | select -First 1
$openResult = Open-Document -LocalFile $filedown.LocalPath -Options @{ FastOpen = $fastOpen } -application "InventorServer"
$openResult





############ exporting the pdf ################

if($openResult  -ne $false) {
    if($openResult.Application.Name -like 'Inventor*') {
        $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF_2D.ini"
    } else {
        $configFile = "$($env:POWERJOBS_MODULESDIR)Export\PDF.dwg" 
    }                   
    $exportResult = Export-Document -Format 'PDF' -To $localPDFfileLocation -Options $configFile
    
   

############# End of export #################

    if($exportResult) {   
    ################## Putting the file where it needs to go ################   
    try{ 
        $PDFadd = Add-VaultFile -From $localPDFfileLocation -To $vaultPDFfileLocation -FileClassification None -Hidden $hidePDF 

       
        $filechange = Update-VaultFile -File $PDFadd._FullPath -Category $file._CategoryName 
        $filechange2 = Update-VaultFile -File $PDFadd._FullPath -properties @{"Comments" = $file._Comments ;"Rev Number" = $file._RevNumber; "Stock Number" = $file._StockNumber ; "Material" = $file.Material; "Project01" = $file."Project01"; "Description"= $file."Description"; "Manager"= $file."Manager"; "Title"= $file."Title"; "Work Order"= $file."Work Order"} 
        $filechange3 = Update-VaultFile -File $PDFadd._FullPath -Status $file._State 
            }catch{
       $PDFadd = $false
       "could not add File"}
        $nameofile = $currentfileselect.Name
        ########### Adding more results to the Table ########
        if($PDFadd){
        $prow.$hp1 = $($file._Name)
        $prow.$hp2 = $pdffilename
        $prow.$hp4 = $vaultPDFfileLocation}
        ##################### Moving the file if there was an exisitng file .. only if the file gets created #########################
       
    
       
        
        
        ##### new line need to test to ensure it still doesnt try to move fiels that dont exist ############
        if((!($currentfileselect.Name)) -or ($currentfileselect.Name -match $filechange.Name)) { write-host "No exisiting Files"
        $prow.$hp3 = "No Pre Exisiting PDF"
         }
      
         

        #####################  ###################################
        else{
            $currentfilepath = $file.Path + "/" + $currentfileselect.Name
            $current = get-vaultfiles -properties @{'Name'= $nameofile}
            $pdfchangestate = Update-VaultFile -File $current._FullPath -status "Obsolete"
            $pdfsource = $vault.DocumentService.GetFolderByPath($current.Path)
            $movefile = $vault.DocumentService.MoveFile($current.MasterID,$pdfsource.Id, $obsoletefolder.Id)
                write-host "Moved $nameofile into the Obsolete Folder"
                $pdfcount += 1
                $prow.$hp3 = $nameofile
                }
                $output = $prow 
                $Output | Export-Csv -Append  -NoTypeInformation -Path $CSVFile -force

                
                }

                ########################## END OF LOOP ##################################


   }
  
     }
     
$closeResult = Close-Document

Write-Host "Completed job 'Create PDF '"


#Clean-Up -folder $workingDirectory

$ErrorActionPreference = "SilentlyContinue"
}

$minuteoftime = $times.minutes
$timerow.$ht4 = "1"
$timerow.$ht5 = $minuteoftime
$timerend = Get-Date
$timerow.$ht3 = $timerend
$timeout = $timerow
$timeout | export-csv -Append -NoTypeInformation -path $timereportcsv -force
}catch{"There was an error please check the job"}
stop-process -name "Inventor*"
}