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
####################################################
write-host "Starting Job Create PDF"
########## Importing Cool Orange Modules #######
import-module powervault 
import-module powerjobs
Import-Module "C:\ProgramData\coolOrange\powerJobs\Modules\VaultSearch.psm1"
######### ***** current user for vault session ******* #############################################
open-vaultconnection -server "Wendt-vault" -vault "Wendt" -user "coolorange" -password "nhg544FK"
#**************************************************************************************************#
$times = Measure-command {
#$lasthour =  
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

############\\ Use this switch if you would like to get files from yesterday // ###############
#$datetime = [DateTime]::Today.AddDays(-1)

#******************************************#
$filestoget = New-SearchCondition -PropertyName "Date Version Created" -searchOperator:GE -searchRuleType:Must -searchText $datetime
$released = New-SearchCondition -PropertyName "File Extension" -searchOperator:Contains -searchRuleType:Must -searchText "idw"
$filestogetidw = New-SearchCondition -PropertyName "State" -searchOperator:Contains -searchRuleType:Must -searchText "Released"
$nopropertyedit = New-SearchCondition -PropertyName "Comment" -searchOperator:NotContains -searchText "Property Edit"
$results = Find-VaultFiles -RecurseFolders -SearchConditions @($filestoget,$released,$filestogetidw,$nopropertyedit) -LatestFilesOnly

################____ Getting vault files _____ ##################
foreach($result in $results){
$resultname = $result.Name
$idwfiles += get-vaultfiles -properties @{'Name'=$resultname}}


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


###############################################****

$pdffilename = $idwfilename +"-"+ $idwrev + ".pdf"
$localPDFfileLocation = "$workingDirectory\$pdffilename"
$vaultPDFfileLocation = $idw._EntityPath +"/"+ (Split-Path -Leaf $localPDFfileLocation)
$fastOpen = $idw._Extension -eq "idw" -or $idw._Extension -eq "dwg" -and $idw._ReleasedRevision
$manager = $idw._Manager
$networkfilepath = "\\wal-file.wendtcorp.local\eng\PDF Files"

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
$openResult = Open-Document -LocalFile $filedown.LocalPath -Options @{ FastOpen = $fastOpen } -application "InventorServer"
$openResult



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
        try{ 
        $PDFadd = Add-VaultFile -From $localPDFfileLocation -To $vaultPDFfileLocation -FileClassification None -Hidden $hidePDF 
      
      
        $filechange = Update-VaultFile -File $PDFadd._FullPath -Category $idw._CategoryName 
        $filechange2 = Update-VaultFile -File $PDFadd._FullPath -properties @{"Comments" = $idw._Comments ;"Rev Number" = $idw._RevNumber; "Stock Number" = $idw._StockNumber ; "Material" = $idw.Material; "Project01" = $idw."Project01"; "Description"= $idw."Description"; "Manager"= $idw."Manager"; "Title"= $idw."Title"; "Work Order"= $idw."Work Order"} 
        $filechange3 = Update-VaultFile -File $PDFadd._FullPath -Status $idw._State }
        catch{
       $pdfadd = $false
       "could not add File"}

        $nameofile = $currentfileselect.Name
    
        ########### Adding more results to the Table ########
       if($pdfadd){
        $prow.$hp1 = $idw.Name
        $prow.$hp2 = $pdffilename
        $prow.$hp4 = $vaultPDFfileLocation
        ##################### Moving the file if there was an exisitng file .. only if the file gets created #########################
       
    
       
      
        
        ##### new line need to test to ensure it still doesnt try to move fiels that dont exist ############
        if((!($currentfileselect.Name)) -or ($currentfileselect.Name -match $filechange.Name)) { write-host "No exisiting Files"
      
        $prow.$hp3 = "No Pre Exisiting PDF"}
         
      
         

        #####################  ###################################
        else{
            $currentfilepath = $idw.Path + "/" + $currentfileselect.Name
            $current = get-vaultfiles -properties @{'Name'= $nameofile}
            $pdfchangestate = Update-VaultFile -File $current._FullPath -status "Obsolete"
            $pdfsource = $vault.DocumentService.GetFolderByPath($current.Path)
            $movefile = $vault.DocumentService.MoveFile($current.MasterID,$pdfsource.Id, $obsoletefolder.Id)
                write-host "Moved $nameofile into the Obsolete Folder"
                $pdfcount += 1
                $prow.$hp3 = $nameofile
                }
                $output = $prow 
                $Output | Export-Csv -Append  -NoTypeInformation -Path "C:\Vault Data\Vault Reports\DWF_Lists\dwflist-$date1.csv" -force

                
                }
                }
                Close-Document
                Clean-Up -folder $workingDirectory
                }

                ########################## END OF LOOP ##################################


   }
     }


Write-Host "Completed job 'Create PDF '"


$timerow.$ht4 = $pdfcount
$timerow.$ht5 = $times.minutes
$timerend = Get-Date
$timerow.$ht3 = $timerend
$timeout = $timerow
$timeout | export-csv -Append -NoTypeInformation -path $timereportcsv -force

Stop-Process -name inventor*


