import-module powervault 
Import-Module "C:\ProgramData\coolOrange\powerJobs\Modules\VaultSearch.psm1"
######### ***** current user for vault session ******* #############################################
open-vaultconnection -server "Wendt-vault" -vault "Wendt" -user "coolorange" -password "nhg544FK"

#=============================================================================#
# PowerShell script  for coolOrange powerJobs to get vault files and export   #
#                                                                             #
#                                                                             #
#                                                                             #
#                                                                             #
#=============================================================================#

<# SrchOper Table
Search Operator 			Number	operator		Valid on property types 		SearchText needed 
Contains 					1		-contains		string 							yes 
Does not contain 			2 		-notcontains	string 							yes 
Is exactly (or equals) 	3 		-eq				numeric, bool, datetime, string	yes 
Is empty 					4 		-empty			image, string 					no 
Is not empty 				5 		-				image, string 					no 
Greater than 				6 		-gt				numeric, datetime, string 		yes 
Greater than or equal to 	7 		-ge				numeric, datetime, string 		yes 
Less than 					8 		-lt				numeric, datetime, string 		yes 
Less than or equal to 		9 		-le				numeric, datetime, string 		yes 
Not equal to 				10 		-ne				numeric, bool, string 			yes 
"MM/dd/yyyy HH:mm:ss"
#>

Add-Type @"
namespace cOEnums {
public enum SearchOperators {
Contains = 0x001,
NotContains = 0x002,
EQ = 0x003,
Empty = 0x004,
NotEmpty = 0x005,
GT = 0x006,
GE = 0x007,
LT = 0x008,
LE = 0x009,
NE = 0x010,
}
}
"@

function Get-PropertyDefinitionByName {
param (
[string]$PropertyName
)
    $propdefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE")
    $propdefs | Where-Object {$_.DispName -eq $PropertyName}
}

function New-SearchCondition{
<#
.SYNOPSIS
Returns a Vault search condition
.EXAMPLE
$srchCond_commentsNotEmpty = New-SearchCondition -PropertyName Comments -searchOperator:NotEmpty
#>
param(
[Parameter(Mandatory=$True,ParameterSetName="PropName")]$PropertyName,
[Parameter(Mandatory=$True,ParameterSetName="PropdefId")]$PropDefId,
[Parameter(Mandatory=$True)][cOEnums.SearchOperators]$searchOperator,
[Autodesk.Connectivity.WebServices.SearchRuleType]$searchRuleType = [Autodesk.Connectivity.WebServices.SearchRuleType]::Must,
[string]$searchText
)
	if($PSCmdlet.ParameterSetName -eq "PropName") {
		$PropertyDefinition = Get-PropertyDefinitionByName $PropertyName
		$PropDefId = $PropertyDefinition.Id
	}
	$searchType = [Autodesk.Connectivity.Webservices.PropertySearchType]::SingleProperty
	
	$condition = New-Object Autodesk.Connectivity.WebServices.SrchCond

	$condition.PropDefId = $PropDefId
	$condition.PropTyp = $searchType
	$condition.SrchOper = $searchOperator
	$condition.SrchRule = $searchRuleType
	$condition.SrchTxt = $searchText

	return $condition
}
function New-SearchSort {
param(
[Parameter(Mandatory=$True,ParameterSetName="PropName")]$PropertyName,
[Parameter(Mandatory=$True,ParameterSetName="PropdefId")]$PropDefId,
[Parameter(Mandatory=$True)][bool]$SortAsc
)
	if($PSCmdlet.ParameterSetName -eq "PropName") {
		$PropertyDefinition = Get-PropertyDefinitionByName $PropertyName
		$PropDefId = $PropertyDefinition.Id
	}
    $srchSort = New-Object Autodesk.Connectivity.WebServices.SrchSort
    $srchSort.PropDefId = $PropDefId
    $srchSort.SortAsc = $SortAsc

    return $srchSort
}
function Find-VaultFiles {
<#
.SYNOPSIS
Returns an array of Vault file objects or powerVault file objects based on the passed in search conditions
.EXAMPLE
$srchCond_commentsNotEmpty = New-SearchCondition -PropertyName Comments -searchOperator:NotEmpty
$results = Find-VaultFiles -RecurseFolders -SearchConditions @($srchCond_commentsNotEmpty)
.EXAMPLE
$srchCond_DateApril4thMin = New-SearchCondition -PropertyName "Date Version Created" -searchOperator:GE -searchRuleType:Must -searchText "04/03/2017 22:00:00"
$srchCond_DateApril4thMax = New-SearchCondition -PropertyName "Date Version Created" -searchOperator:LT -searchRuleType:Must -searchText "04/04/2017 22:00:00"
$results = Find-VaultFiles -RecurseFolders -SearchConditions @($srchCond_DateApril4thMin, $srchCond_DateApril4thMax) -LatestFilesOnly -RootFolderPath "$"
#>
param(
[string]$RootFolderPath = "$",
[Autodesk.Connectivity.WebServices.SrchCond[]]$SearchConditions,
[Autodesk.Connectivity.WebServices.SrchSort]$SortConditions = $null,
[switch]$RecurseFolders,
[switch]$LatestFilesOnly
)
	$rootFolder = $vault.DocumentService.GetFolderByPath($RootFolderPath)
	$folderIds = @($rootFolder.Id)
	if([string]::IsNullOrEmpty($folderIds)) {
		throw("FolderId not set")
	}

	$bookmark = $null
	$searchstatus = $null

	$hits = @()
	do {
		[array]$hits += $vault.DocumentService.FindFilesBySearchConditions($SearchConditions,$SortConditions,$folderIds,($RecurseFolders.ToBool()),($LatestFilesOnly.ToBool()),[ref]$bookmark,[ref]$searchstatus)
		#Write-Host -Object ("{0}/{1}" -f $hits.Count, $searchstatus.TotalHits) -ErrorAction:SilentlyContinue
	} while($hits.Count -lt $searchstatus.TotalHits)
	
	return $hits
}





$tempelectricalfile = "C:\Users\Public\Documents\Wendt-Vault Electric Projects.csv"

$finalelectricalfile = "C:\Vault Data\Vault Reports\Wendt-Vault Electric Projects.csv"
$clear = clear-Content $tempelectricalfile 
######################### Creating A Table ####################################
$h1 = "File Extension"
$h2 = "Name"
$h3 = "Work Order"
$h4 = "Part Number"
$h5 = "State"
$h6 = "Revision"
$h7 = "Version"
$h8 = "Created By"
$h9 = "Date Modified"
$h10 = "Description"
$h11 = "File Name"
$h12 = "Manager"
$h13 = "Title"
$h14 = "Stock Number"
$h15 = "Material"
$h16 = "Project"
$h17 = "Mass"
$h18 = "Area"
$h19 = "Volume"
$H20 = "Thickness"
$h21 = "Has Drawing"
$h22 = "Checked Out By"

$row = "" |select-Object $h1, $h2,$h3,$h4,$h5,$h6,$h7,$h8,$h9,$h10,$h11,$h12,$h13,$h14,$h15,$h16,$h17,$h18,$h19,$h20,$h21,$h22
$output = @()


########################################******************* Electrical script ************************##########################


$electricalresult = Find-VaultFiles -RootFolderPath "$/Wendt/Electrical DWGs-Docs" -RecurseFolders -LatestFilesOnly
foreach($efile in $electricalresult){
$itemname = $efile.Name
$vfile1 = get-vaultfiles -Properties @{'Name'="$itemname"}
$fname = $vfile1.Name
if($fname -ne $null){

$row.$h1 = $vfile1._Extension
$row.$h2 = $vfile1.Name
$row.$h3 = $vfile1."Work Order"
$row.$h4 = $vfile1._PartNumber
$row.$h5 =$vfile1.State
$row.$h6 =$vfile1.Revision
$row.$h7 =$vfile1.Version
$row.$h8 =$vfile1._CreateUserName
$row.$h9 =$vfile1._ModDate
$row.$h10 =$vfile1.Description
$row.$h11 =$vfile1._ClientFileName
$row.$h12 =$vfile1.Manager
$row.$h13 =$vfile1.Title
$row.$h14 =$vfile1._StockNumber
$row.$h15 =$vfile1.Material
$row.$h16 =$vfile1."Project"
$row.$h17 =$vfile1.Mass
$row.$h18 =$vfile1.Area
$row.$h19 =$vfile1.Volume
$row.$h20 =$vfile1.Thickness
$row.$h21 = $vfile1."Has Drawing"
$row.$h22 =$vfile1._CheckoutUserName

$output = $row 
$Output | Export-Csv -Append  -NoTypeInformation -Path $tempelectricalfile -force


}


}
copy-item -Path $tempelectricalfile -Destination $finalelectricalfile -Force