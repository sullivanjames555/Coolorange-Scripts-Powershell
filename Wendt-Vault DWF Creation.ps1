##########################################################################################################
#        This script is designed to Create Dwf Files based on what items were checked in today           #
#                              *** Functions for Real Vault  ***                                         #
#                                                                                                        #
#                        ______        _______    ____                _   _                              #
#                       |  _ \ \      / /  ___|  / ___|_ __ ___  __ _| |_(_) ___  _ __                   #
#                       | | | \ \ /\ / /| |_    | |   | '__/ _ \/ _` | __| |/ _ \| '_ \                  #
#                       | |_| |\ V  V / |  _|   | |___| | |  __/ (_| | |_| | (_) | | | |                 #
#                       |____/  \_/\_/  |_|      \____|_|  \___|\__,_|\__|_|\___/|_| |_|                 #                                                                                                                #
#                                  Above Created with ANSII Art Generator                                #
#                                                                                                        #
#            *****    This needs to be ran by a vault administrator to overwrite files    *****          #
#       #error logs for the script can be found in %userprofile%\appdata\local\coolorange\powerjobs      #
#                                  Created by James Sullivan 9-3-20 rev 5                                #
##########################################################################################################
############***##############*** needs powervault installed on machine ***##############***###############

####################### Gets installed Modules ########################
$ErrorActionPreference = "SILENTLYCONTINUE"
import-module powervault 
import-module powerjobs 

######### current user for vault session #################################################################
open-vaultconnection -server "wendt-vault" -vault "wendt" -user "coolorange" -password "nhg544FK"
set-location -path "C:\Temp\"
######### getdate and turn into date-time format ##############
$returnobj = @()
$date = Get-Date -Format "dddd, MMMM d, yyyy"
$datetime = [datetime] $date
$datetime = [DateTime]::Today.AddDays(-1)

#$datetime = [datetime] $date
########### gets all iam files in vault #########
$vfile = get-vaultfiles -properties @{'Name'='*.iam';'State' = '*?';'Manager'='*?'}
$vfile1 = $vfile |where-object {$_."Date Modified (Date Only)" -like $datetime -and ($_.Name -inotlike "delete*")}
$workingDirectory = "C:\Temp\"

#region creating tables for reports 
### creating report for the daily count 

$timerstart = get-date 
$timereportcsv = "C:\Vault Data\Vault Reports\Other Reports\CoolorangeJobs.csv"
$dwfcount = @()

$ht1 = "Script Name"
$ht2 = "Start time"
$ht3 = "End Time"
$ht4 = 'Count'
$ht5 = 'Total Time'
$timerow = "" |select-Object $ht1, $ht2,$ht3,$ht4,$ht5

$timerow.$ht1 = "Wendt - Vault Create dwf's"
$timerow.$ht2 = $timerstart
$timerow.$ht4 = $vfile1.count
####

##### Creating table for dwf report ####


$h1 = "Dwf Name"
$h2 = "Full Path"
$h3 = "CreatedBy"
$h4 = 'date'

$dwfrow = "" |select-Object $h1, $h2,$h3,$h4
$csvpath = "C:\Vault Data\Vault Reports\DWF_Lists\dwflist-$date.csv"

#endregion 

#region Start Dwf PRocess
$times = Measure-command {

############ loop statements to save the files to your local ###########

#if $vfile1 is null it will not run this script (if nothing is checked in that day)
if ($vfile1 -ine $null){
write-host "Found Dwf Files, Now saving and converting"
$dwf = $files._EntityPath

$entitypath = $save._FullPath

############# selects and parses each item that fits the criteria for $iamfile needed to convert the correct files and not sub assemblies ###########
foreach($item in $vfile1){
$newname = $item._Name
if (($item._Comment -eq "Property Edit") -and ($item.Comments -eq $null)){
write-host "$newname does not need to be converted"}
else{
write-host "now converting $newname"
$save = save-vaultfile -file $item._FullPath -DownloadDirectory "C:\Temp\"
$iamfile = $save | where-object {$_.Name -eq $item.Name}
$manager = $iamfile._Manager
$title = $item._Title
$hashtable = @{ "_Manager" = $manager;"_Title" = $title}

############ creates a name for each item in the loop ##########################################################

$Dwfname = $iamfile.LocalPath + ".dwf"
$localdwffileLocation = $iamfile.LocalPath + ".dwf"
$vaultdwffileLocation = $item.Path +"/"+ $iamfile.Name + ".dwf"
$dwf_name = split-path $vaultdwffileLocation -leaf

$dwfrow.$h1. $dwf_name
$dwfrow.$h2 = $vaultdwffileLocation
$dwfrow.$h3 = $item._CreateUserName
$dwfrow.$h4 = $item._ModDate
$dwfrow | Export-Csv -Append  -NoTypeInformation -Path $csvpath -force

########### opens the vault item with inventor server in the background, this uses resources ###################
$doc = open-document -localfile $iamfile.LocalPath
$configFile = "$($env:POWERJOBS_MODULESDIR)Export\DWF_3D.ini"
######################### Beginning the Export to dwf using the inventor server large files may utilize up to 40 gb of ram to export ################
$export = Export-Document -Format "DWF" -To "$Dwfname" -Options $configFile -OnExport {
            param($export)
                $document = $export.SourceDocument
                switch([Inventor.DocumentTypeEnum]$document.Instance.DocumentType)
                {
                    ([Inventor.DocumentTypeEnum]::kPartDocumentObject) { 
                        if($document.Instance.ComponentDefinition.IsiPartFactory) {
                            $export.Options["iPart_All_Members"] = $export.Options["iPart_3D_Models"] = $false
                            if(-not $export.Options["iPart_All_Members"]) {
                                $defaultRow = $document.Instance.ComponentDefinition.iPartFactory.DefaultRow
                                $export.Options['iAssemblies'] = $document.Application.Instance.TransientObjects.CreateNameValueMap()
                                $iassemblyOptions = $document.Application.Instance.TransientObjects.CreateNameValueMap()
                                $iassemblyOptions.Value("Name", $defaultRow.MemberName)
                                $iassemblyOptions.Value("3DModel", $true)
                                $export.Options['iParts'].Value('Name', $iassemblyOptions)
                            }
                        }
                    } 
                    ([Inventor.DocumentTypeEnum]::kAssemblyDocumentObject) {
                        $export.Options['Design_Views'] = $document.Application.Instance.TransientObjects.CreateNameValueMap()
                        $designViewRepresentations = $document.Instance.ComponentDefinition.RepresentationsManager.DesignViewRepresentations
                        for($i=1; $i -le $designViewRepresentations.Count; $i++) {
                            $designViewRepresentation = $designViewRepresentations.Item($i)
                            if ($designViewRepresentation.DesignViewType -eq [Inventor.DesignViewTypeEnum]::kMasterDesignViewType -or $designViewRepresentation.DesignViewType -eq [Inventor.DesignViewTypeEnum]::kPublicDesignViewType ) {
                                $designViewRepresentationOptions = $document.Application.Instance.TransientObjects.CreateNameValueMap()
                                $designViewRepresentationOptions.Add("Name",  $designViewRepresentation.Name)
                                $export.Options['Design_Views'].Value("Design_View" + $i.ToString("D")) = $designViewRepresentationOptions
                            }
                        }

                        $export.Options['Positional_Representations'] = $document.Application.Instance.TransientObjects.CreateNameValueMap()
                        $positionalRepresentations = $document.Instance.ComponentDefinition.RepresentationsManager.PositionalRepresentations
                        for($i=1; $i -le $positionalRepresentations.Count; $i++) {
                            $representation = $positionalRepresentations.Item($i)
                            $representationOptions = $document.Application.Instance.TransientObjects.CreateNameValueMap()
                            $representationOptions.Add("Name",  $representation.Name)
                            $export.Options['Positional_Representations'].Value("Positional_Representation" + $i.ToString("D")) = $representationOptions
                        }

                        if($document.Instance.ComponentDefinition.IsiAssemblyFactory) {
                            $export.Options["iAssembly_All_Members"] = $export.Options["iAssembly_3D_Models"] = $false
                            if(-not $export.Options["iAssembly_All_Members"]) {
                                $defaultRow = $document.Instance.ComponentDefinition.iAssemblyFactory.DefaultRow
                                $export.Options['iAssemblies'] = $document.Application.Instance.TransientObjects.CreateNameValueMap()
                                $iassemblyOptions = $document.Application.Instance.TransientObjects.CreateNameValueMap()
                                $iassemblyOptions.Value("Name", $defaultRow.MemberName)
                                $iassemblyOptions.Value("3DModel", $true)
                                $export.Options['iAssemblies'].Value('Name', $iassemblyOptions)
                            }
                        }
                    }
                    ([Inventor.DocumentTypeEnum]::kPresentationDocumentObject) {
                        $presentationExplodedViews = $document.Application.Instance.PresentationExplodedViews
                        if($presentationExplodedViews.Count -ne 0) {
                            $export.Options['Presentations'] = $document.Application.Instance.TransientObjects.CreateNameValueMap()
                            for($i=1; $i -le $presentationExplodedViews.Count; $i++) {
                                $presentationExplodedView = $presentationExplodedViews.Item($i)
                                $presentationExplodedViewOptions = $document.Application.Instance.TransientObjects.CreateNameValueMap()
                                $presentationExplodedViewOptions.Add("Name",  $presentationExplodedView.Name)
                                $export.Options['Presentations'].Value("Presentation" + $i.ToString()) = $presentationExplodedViewOptions
                            }
                        }
                    }
                }}
       

       #export is done and saved locally
############# # adds the vault file abck in vault. may not be successful based on vault restrictions and pre exisiting dwf files
	
         
         $addfile = Add-VaultFile -From $dwfname -To $vaultdwffileLocation  -FileClassification "DesignVisualization" -Hidden $true -force $true
 	 $updatedwf = Update-VaultFile -file $vaultdwffilelocation -properties $hashtable
         #$updateiam = Update-VaultFile -file $item._FullPath -AddAttachments @($vaultdwffileLocation) -comment "CoolOrange Created"





########################################################################################           
       
       

       ### closes out the current document in the loop ######
        $closeResult = Close-Document
        }
        ####### Cleans out the temp folder directory. May still be files left in vault workspace *potential bug

}}
# if there is no files in $vfile this script clsoes
    else{
write-host "Error: no files were present"
     } 
#$returnobj | Export-CSV "C:\Vault Data\Vault Reports\DWF_Lists\dwflist-$date.csv" -NoTypeInformation


Clean-Up -folder $workingDirectory
}
#endregion dwf

$minuteoftime = $times.minutes
$timerow.$ht5 = $minuteoftime
$timerend = Get-Date
$timerow.$ht3 = $timerend

$timerow | export-csv -Append -NoTypeInformation -path $timereportcsv -force

stop-process -name "Inventor*"