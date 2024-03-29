#=============================================================================#
# PowerShell script sample for coolOrange powerJobs                           #
# Creates a IGES file and add it to Autodesk Vault as Design Vizualization    #
#                                                                             #
# Copyright (c) coolOrange s.r.l. - All rights reserved.                      #
#                                                                             #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER   #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.  #
#=============================================================================#

$hideIGES = $false
$workingDirectory = "C:\Temp\$($file._Name)"
$localIGESfileLocation = "$workingDirectory\$($file._Name).igs"
$vaultIGESfileLocation = $file._EntityPath +"/"+ (split-path -Leaf $localIGESfileLocation)

Write-Host "Starting job 'Create IGES as attachment' for file '$($file._Name)' ..."

if( @("iam","ipt") -notcontains $file._Extension ) {
    Write-Host "Files with extension: '$($file._Extension)' are not supported"
    return
}

$file = Get-VaultFile -File $file._FullPath -DownloadPath $workingDirectory
$openResult = Open-Document -LocalFile $file.LocalPath

if($openResult) {
    $exportResult = Export-Document -Format 'IGES' -To $localIGESfileLocation -Options  "$($env:POWERJOBS_MODULESDIR)Export\IGES.ini"
    if($exportResult) {
        $IGESfile = Add-VaultFile -From $localIGESfileLocation -To $vaultIGESfileLocation -FileClassification DesignVisualization -Hidden $hideIGES
        $file = Update-VaultFile -File $file._FullPath -AddAttachments @($IGESfile._FullPath)
    }        
    $closeResult = Close-Document
}
Clean-Up -folder $workingDirectory

if(-not $openResult) {
    throw("Failed to open document $($file.LocalPath)! Reason: $($openResult.Error.Message)")
}
if(-not $exportResult) {
    throw("Failed to export document $($file.LocalPath) to $localIGESfileLocation! Reason: $($exportResult.Error.Message)")
}
if(-not $closeResult) {
    throw("Failed to close document $($file.LocalPath)! Reason: $($closeResult.Error.Message))")
}
Write-Host "Completed job 'Create IGES as attachment'"