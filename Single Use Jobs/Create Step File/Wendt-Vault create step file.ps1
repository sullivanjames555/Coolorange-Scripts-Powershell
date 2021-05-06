<# #######################################################################

#  _________ __                  ___________.__.__          
# /   _____//  |_  ____ ______   \_   _____/|__|  |   ____  
# \_____  \\   __\/ __ \\____ \   |    __)  |  |  | _/ __ \ 
# /        \|  | \  ___/|  |_> >  |     \   |  |  |_\  ___/ 
#/_______  /|__|  \___  >   __/   \___  /   |__|____/\___  >
#        \/           \/|__|          \/                 \/ 
        Step file generation Script. this file converts an iam
        to a step file using coolorange powerjobs and powervault 
       ###################### V1.0 ###########################
        Created 4-1-21
        Creator: James Sullivan


#> #########################################################################

#import-modules and classes 
import-module powervault 
import-module powerjobs 
Add-Type -AssemblyName Microsoft.VisualBasic

######### current user for vault session ##########################################################
#region Account
$vserver = "wendt-vault"
$vvault = "wendt" 
$vuser = "coolorange" 
$vpw = "nhg544FK"
#endregion

open-vaultconnection -server $vserver  -vault $vvault -user $vuser -password $vpw
$date = get-date -format "MM/dd/yyyy"

##Building the variables 

$vfile = [Microsoft.VisualBasic.Interaction]::InputBox('Please Enter The File Name. Include File Extensions', 'Assembly File', "(EX..800-123-1234.iam)")
$vfile1 = get-vaultfiles -properties @{'Name'="$vfile"}
$inifile = "C:\ProgramData\coolOrange\powerJobs\Modules\Export\STEP.ini"
$workingDirectory = "C:\Temp\Wendt"
$fullpath = $vfile1._FullPath
$entitypath = $vfile1.Path

write-host "saving files"
#saving the file
$save = save-vaultfile -file $fullpath -DownloadDirectory "C:\Temp\" 
$iamfile = $save |where-object {$_.Name -eq $vfile1.Name}
$Stpname = $iamfile.LocalPath + ".stp" 
$movefile = $iamfile._Name + ".stp"
$localstpfileLocation = $iamfile.LocalPath + ".stp"
$vaultstpfileLocation =$entitypath + "/" + $iamfile.Name + ".stp"
$doc = open-document -localfile $iamfile.LocalPath
write-host "now exporting to step"

#one line export using the configuration file 
$export = Export-Document -Format "STEP" -To $stpname -Options "$inifile"

#adding the vault file 
Add-VaultFile -From $localstpfileLocation -To $vaultstpfileLocation  -FileClassification "NONE" -force $true

$closeResult = Close-Document
clean-up $workingDirectory