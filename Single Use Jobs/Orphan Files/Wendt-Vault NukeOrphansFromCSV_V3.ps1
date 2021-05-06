######################
#  Vault Login Information
######################

$vaultName = "Wendt"
$vaultServerName = "wendt-vault"
$vaultUserName = "coolorange"
$vaultPassword = "nhg544FK"


Open-VaultConnection -Vault $vaultName -Server $vaultServerName -user $vaultUserName -Password $vaultPassword

#########  SCRIPT VARIABLES
$csvFileLocation = "C:\Vault Data\Vault Reports\orphans 4-12-21.csv"
$outputFileName = "C:\Temp\Deleted_Orphans$(Get-Date -Format "yyyyMMdd-hhmmss")_Log" 

####### File prefixes to delete


$90dayDeletePrefixes = "801", 803

$90dayDeletePrefixes += "804"
$90dayDeletePrefixes += "807"
$90dayDeletePrefixes += "810"
$90dayDeletePrefixes += "840"
$90dayDeletePrefixes += "854"
$90dayDeletePrefixes += "907"
$90dayDeletePrefixes += "909"
$90dayDeletePrefixes += "910"
$90dayDeletePrefixes += "914"
$90dayDeletePrefixes += "917"
$90dayDeletePrefixes += "928"
$90dayDeletePrefixes += "930"

$120dayDeletePrefixes = “907-4”,”930-4”,“290-4”,”325-4”,”914-4”
$managerpreffix =  "203121","208458"


##The below variable is used to determine if files will actually be deleted.  The CSV log will always be created
$deleteFiles = $False


#####################


function DeleteFile
{
    Param($masterID, $folderID)
    if($deleteFiles)
    {
        try
        {
            $vault.DocumentService.DeleteFileFromFolder($masterID, $folderID)
            $result = "Succeeded"
        }
        catch
        {
            $result = "Failed"
        }
    }
    else
    {
        $result = "Simulated"
    }
    
    return $result
}



$dt = New-Object System.Data.Datatable
[void]$dt.Columns.Add("FileName")
[void]$dt.Columns.Add("VaultPath")
[void]$dt.Columns.Add("Date")
[void]$dt.Columns.Add("Result")

[System.DateTime]$currFileDate = Get-Date


$fileTable = Import-Csv -Path $csvFileLocation
$now = Get-Date

$deleteHash = @{}
foreach ($fileRow in $fileTable)
{
    #Write-Host "$($fileRow."FileName") -- $($fileRow.Date)"
        
        try
        {
            $currFileDate = $fileRow.Date
        }
        catch
        {
            #if this fails the file is probably currently checked out; continue to the next row of the table
            continue
        }

        #See if this is a file we're interested in based on prefix for 90 day delete 
        foreach($prefix in $90dayDeletePrefixes)
        {
            if($fileRow.FileName.StartsWith($prefix))
            {
                #See if it's older than 90 days
                if (($now - $currFileDate).Days -gt 90)
                {
                    # Ignore the -4- files
                    if ($fileRow.FileName.StartsWith($nukeStartString + "-4")) {continue}
                    #otherwise delete it
                    
                    $result = DeleteFile -masterID $fileRow.MasterID -folderID $fileRow.FolderID

                }
                else
                {
                    $result = "Skipped - Age"
                }
       
                $dt.Rows.Add($fileRow.FileName, $fileRow.VaultPath, $fileRow.Date, $result)
                #file is handled - continue the to the next one
                continue
            }
        }

        #See if this is a file we're interested in based on prefix for 120 day delete 
        foreach($prefix in $120dayDeletePrefixes)
        {
            if($fileRow.FileName.StartsWith($prefix))
            {
                #See if it's older than 90 days
                if (($now - $currFileDate).Days -gt 120)
                {
                    #delete it
                    $result = DeleteFile -masterID $fileRow.MasterID -folderID $fileRow.FolderID
                }
                else
                {
                    $result = "Skipped - Age"
                }
                $dt.Rows.Add($fileRow.FileName, $fileRow.VaultPath, $fileRow.Date, $result) | Out-Null
            }
        }
    
}


##Write Output
if($deleteFiles -eq $false)
{
    $outputFileName = "$($outputFileName)_SIMULATED.txt"
}
else
{
    $outputFileName = "$($outputFileName).txt"
}
$dt | Export-Csv $outputFileName