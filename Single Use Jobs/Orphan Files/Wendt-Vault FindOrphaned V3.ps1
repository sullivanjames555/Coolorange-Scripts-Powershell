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
$logFileName = "C:\Temp\OrphanSearchLog.txt"


Open-VaultConnection -Vault $vaultName -Server $vaultServerName -user $vaultUserName -Password $vaultPassword

$dt = New-Object System.Data.Datatable
[void]$dt.Columns.Add("FileName")
[void]$dt.Columns.Add("VaultPath")
[void]$dt.Columns.Add("HighestVersion")
[void]$dt.Columns.Add("Category")
[void]$dt.Columns.Add("State")
[void]$dt.Columns.Add("MasterID")
[void]$dt.Columns.Add("FolderID")
[void]$dt.Columns.Add("Date")



#$allFiles = Get-VaultFiles -properties @{"File Extension" = "ipt"}
$propdefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
ForEach ($currDef in $propdefs)
{
    if ($currDef.DispName -eq "File Name")
    {
        $filenameDef = $currDef
    }
}

$srchCondition = New-Object -TypeName Autodesk.Connectivity.WebServices.SrchCond
$srchCondition.PropDefId = $filenameDef.Id
$srchCondition.PropTyp = [Autodesk.Connectivity.WebServices.PropertySearchType]::SingleProperty
$srchCondition.SrchOper = 1
$srchCondition.SrchRule = [Autodesk.Connectivity.WebServices.SearchRuleType]::Must
$srchCondition.SrchTxt = "*.ipt"

$srchCondition2 = New-Object -TypeName Autodesk.Connectivity.WebServices.SrchCond
$srchCondition2.PropDefId = $filenameDef.Id
$srchCondition2.PropTyp = [Autodesk.Connectivity.WebServices.PropertySearchType]::SingleProperty
$srchCondition2.SrchOper = 2
$srchCondition2.SrchRule = [Autodesk.Connectivity.WebServices.SearchRuleType]::Must
$srchCondition2.SrchTxt = "dwf"

$bookmark = ""
$status = $null
$totalResults = $null

while ($status -eq $null -or $totalCount -lt $status.TotalHits)
{
    $results = $vault.DocumentService.FindFilesBySearchConditions(@($srchCondition,$srchCondition2), $null, $null, $false, $true, [ref]$bookmark, [ref]$status)
    write-host $status.TotalHits
    if ($results)
    {
        $totalCount += $results.Count
        Write-Host $totalCount
        ForEach ($vaultfile in $results)
        {
            try
            {
                $ids = @($vaultfile.Id)
                $parents = $vault.DocumentService.GetFileAssociationLitesByIds($ids, [Autodesk.Connectivity.WebServices.FileAssocAlg]::LatestConsumable, [Autodesk.Connectivity.WebServices.FileAssociationTypeEnum]::All, $false, [Autodesk.Connectivity.WebServices.FileAssociationTypeEnum]::None, $false, $false, $false, $false)
                if ($parents.Count -eq 0)
                {
                    $coFile = Get-VaultFile -FileId $vaultfile.Id
                    $coFolder = 
                    $dt.Rows.Add($coFile.Name, $coFile._FullPath, $coFile.Version, $file.Category, $coFile.State, $coFile.MasterId, $vaultfile.FolderId, $coFile._CheckInDate)
                }
            }
            catch
            {
                "Error procesing file $($file.Name)" | Out-File -FilePath $logFileName -Append
            }
        }
    }
    else
    {
        break
    }
}

$dt | Export-Csv $outputFileName

