##########################################################################################################
#        This script is designed to Create Dwf Files based on what information is inputted               #
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
#                                  Created by James Sullivan 11-12-19 rev 3                              #
##########################################################################################################
############***##############*** needs powervault installed on machine ***##############***###############


import-module powervault 
import-module powerjobs 
Add-Type -AssemblyName Microsoft.VisualBasic

######### current user for vault session #################################################################
open-vaultconnection -server "wendt-vault" -vault "wendt" -user "coolorange" -password "nhg544FK"
$date = get-date -format "MM/dd/yyyy"

$vfile = [Microsoft.VisualBasic.Interaction]::InputBox('Please Enter The File Name. Include File Extensions', 'Assembly File', "(EX..800-123-1234.iam)")
$vfile1 = get-vaultfiles -properties @{'Name'="$vfile"}
write-host $vfile1
$inifile = "C:\ProgramData\coolOrange\powerJobs\Modules\Export\DWF_3D.ini"
$workingDirectory = "C:\Temp\Wendt"
$fullpath = $vfile1._FullPath
$entitypath = $vfile1.Path

write-host "saving files"
$save = save-vaultfile -file $fullpath -DownloadDirectory "C:\Temp\" 
$iamfile = $save |where-object {$_.Name -eq $vfile1.Name}
$Dwfname = $iamfile.LocalPath + ".dwf" 
$movefile = $iamfile._Name + ".dwf"
$localdwffileLocation = $iamfile.LocalPath + ".dwf"
$vaultdwffileLocation =$entitypath + "/" + $iamfile.Name + ".dwf"

$doc = open-document -localfile $iamfile.LocalPath
$export = Export-Document -Format "DWF" -To $Dwfname -Options "$inifile" -OnExport {
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
write-host "Progress"
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
write-host "progress"
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
                }
       
       }
       
      
       
         
   
      Add-VaultFile -From $localdwffileLocation -To $vaultdwffileLocation  -FileClassification "DesignVisualization" -Hidden $true -force $true
      Update-VaultFile -file $iamfile._FullPath -AddAttachments @($vaultdwffileLocation) -comment "CoolOrange Created"     
               
       
       $closeResult = Close-Document
        
#Clean-Up -folder $workingDirectory
	 #read-host "Click enter to stop"
       
        