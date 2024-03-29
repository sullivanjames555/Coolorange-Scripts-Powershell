#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a STEP file and add it to Autodesk Vault as Design Vizualization    #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#

$hideSTEP = $false
$workingDirectory = "C:\Temp\$($file._Name)"
$localSTEPfileLocation = "$workingDirectory\$($file._Name).stp"
$vaultSTEPfileLocation = $file._EntityPath +"/"+ (split-path -Leaf $localSTEPfileLocation)

Write-Host "Starting job 'Create STEP as attachment' for file '$($file._Name)' ..."

if( @("iam","ipt") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}

$file = Get-VaultFile -File $file._FullPath -DownloadPath $workingDirectory
$openResult = Open-Document -LocalFile $file.LocalPath

if($openResult) {    
    $exportResult = Export-Document -Format 'STEP' -To $localSTEPfileLocation -Options "$($env:POWERJOBS_MODULESDIR)Export\STEP.ini"
    if($exportResult) {
        $STEPfile = Add-VaultFile -From $localSTEPfileLocation -To $vaultSTEPfileLocation -FileClassification DesignVisualization -Hidden $hideSTEP
        $file = Update-VaultFile -File $file._FullPath -AddAttachments @($STEPfile._FullPath)
    }
    $closeResult = Close-Document
}
Clean-Up -folder $workingDirectory

if(-not $openResult) {
    throw("Failed to open document $($file.LocalPath)! Reason: $($openResult.Error.Message)")
}
if(-not $exportResult) {
    throw("Failed to export document $($file.LocalPath) to $localSTEPfileLocation! Reason: $($exportResult.Error.Message)")
}
if(-not $closeResult) {
    throw("Failed to close document $($file.LocalPath)! Reason: $($closeResult.Error.Message))")
}
Write-Host "Completed job 'Create STEP as attachment'"