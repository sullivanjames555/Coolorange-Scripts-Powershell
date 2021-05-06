import-module powervault 

######################
#  Vault Login Information
######################

$vaultName = "wendt"
$vaultServerName = "wendt-vault"
$vaultUserName = "coolorange"
$vaultPassword = "nhg544FK"

######################
#  Additional script variables
######################

$outputFileName = "C:\temp\orphans.csv"
$logFileName = "C:\temp\OrphanSearchLog.txt"


Open-VaultConnection -Vault $vaultName -Server $vaultServerName -user $vaultUserName -Password $vaultPassword

$dt = New-Object System.Data.Datatable
[void]$dt.Columns.Add("File Name")
[void]$dt.Columns.Add("Vault Path")
[void]$dt.Columns.Add("Highest Version")
[void]$dt.Columns.Add("Category")
[void]$dt.Columns.Add("State")

$allFiles = Get-VaultFiles -properties @{"Name" = "800*.ipt"}
ForEach ($file in $allFiles)
{
    try{
        $ids = @($file.Id)
        $parents = $vault.DocumentService.GetFileAssociationLitesByIds($ids, [Autodesk.Connectivity.WebServices.FileAssocAlg]::LatestConsumable, [Autodesk.Connectivity.WebServices.FileAssociationTypeEnum]::All, $false, [Autodesk.Connectivity.WebServices.FileAssociationTypeEnum]::None, $false, $false, $false, $false)
        if ($parents.Count -eq 0)
        {
            $dt.Rows.Add($file.Name, $file._FullPath, $file.Version, $file.Category, $file.State)
        }
    }
    catch
    {
        "Error procesing file $($file.Name)" | Out-File -FilePath $logFileName -Append
    }
}

$dt | Export-Csv $outputFileName

