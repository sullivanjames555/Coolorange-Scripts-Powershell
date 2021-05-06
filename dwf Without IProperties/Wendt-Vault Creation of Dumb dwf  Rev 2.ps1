﻿import-module powervault 
import-module powerjobs 
$inventor = Get-Application 'Inventor'
Add-Type -AssemblyName Microsoft.VisualBasic

######### current user for vault session #################################################################
open-vaultconnection -server "wendt-vault" -vault "wendt" -user "coolorange" -password "nhg544FK"
$date = get-date -format "MM/dd/yyyy"
$name = [Microsoft.VisualBasic.Interaction]::InputBox('Please Enter The File Name', 'Assembly File', "(EX..800-123-1234.iam)")
$vaultfile = get-vaultfiles -properties @{'Name'="$name"}
write-host $vaultfile
$workingDirectory = "C:\temp"
$dwfini = "C:\ProgramData\coolOrange\powerJobs\Jobs\dwfx Without IProperties\DWF_Dumb.ini"
$fullpath = $vaultfile._FullPath
$entitypath = $vaultfile.Path

write-host "Now Saving the files for $name"
$save = save-vaultfile -file $fullpath -DownloadDirectory "C:\temp" 
$iamfile = $save |where-object {$_.Name -eq $vaultfile.Name}

$Dwfname = $iamfile.LocalPath +  "-CC"+ ".dwf" 
$movefile = $iamfile._Name + "-CC" + ".dwf"
$localdwffileLocation = $iamfile.LocalPath + "-CC" + ".dwf"
#$vaultdwffileLocation =$entitypath + "/" + $iamfile.Name + ".dwf"
<#
######Dumb dwf local##############
write-Host "lets remove the Iproperties from the document"
[System.Reflection.Assembly]::LoadWithPartialName("Autodesk.Inventor.Interop")

$invApp = New-Object Inventor.ApprenticeServerComponentClass
$oDoc = $invApp.Open($iamfile.LocalPath) # open the assembly
If($odoc.NeedsMigrating -contains "$true") {
[System.Windows.MessageBox]::Show('The FIle Needs to be migrated. Please migrate the file and start again!')
write-host "the file was not migrated, the script will attempt to convert the document this may be unsuccessful"}

$designprops = $odoc.propertysets.Item("Design Tracking Properties")
foreach($propname in $designprops.Name){$designProps.Item(“Part Number”).Value = “Wendt”}

$invApp.FileSaveAs.AddFileToSave($oDoc, $oDoc.FullFileName)
write-host "Now Saving  the Document locally"
$invApp.FileSaveAs.ExecuteSave() # save the assembly
$invApp.Close()

#>

#########end of dumb dwf creation ############


$doc = open-document -localfile $iamfile.LocalPath -Options @{'FastOpen'=$True } -Application "InventorServer"
write-host "converting Document...0%"

$export = Export-Document -Format "DWF" -To $Dwfname -Options ".\DWF_Dumb.ini" -OnExport {
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
				write-host "Still converting Document...10%"
			
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
				write-host "Still converting Document...25%"
                            }
                        }

                        $export.Options['Positional_Representations'] = $document.Application.Instance.TransientObjects.CreateNameValueMap()
                        $positionalRepresentations = $document.Instance.ComponentDefinition.RepresentationsManager.PositionalRepresentations
                        for($i=1; $i -le $positionalRepresentations.Count; $i++) {
                            $representation = $positionalRepresentations.Item($i)
                            $representationOptions = $document.Application.Instance.TransientObjects.CreateNameValueMap()
                            $representationOptions.Add("Name",  $representation.Name)
                            $export.Options['Positional_Representations'].Value("Positional_Representation" + $i.ToString("D")) = $representationOptions
				write-host "Still converting Document...50%"
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
				write-host "Still converting Document...70%"
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
				write-host "Still converting Document... almost done"
                            }
                        }
                    }
                }
       
       }
       write-host "Done converting the document ...100%"
      
       
         $savelocation = "\\wal-file.wendtcorp.local\Public\No Properties DWF"

   	copy-item -path $localdwffilelocation -destination "$savelocation\$movefile"
 	write-host " This File has been saved in: \\wal-file.wendtcorp.local\Public\No Properties DWF"

       
       $closeResult = Close-Document
      Stop-Process -Name "*Inventor"
        
 #Clean-Up -folder $workingDirectory
       
        read-host "Click enter to stop"
