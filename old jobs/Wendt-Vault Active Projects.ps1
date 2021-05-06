$timerstart = get-date 
$timereportcsv = "C:\Vault Data\Vault Reports\Other Reports\CoolorangeJobs.csv"

$ht1 = "Script Name"
$ht2 = "Start time"
$ht3 = "End Time"
$ht4 = "Count"
$ht5 = "Total Time"
$timerow = "" |select-Object $ht1, $ht2,$ht3,$ht4,$ht5
$timeout = @()

$timerow.$ht1 = "Wendt- Vault Active Projects"
$timerow.$ht2 = $timerstart

$timing = measure-command{
####################### Gets installed Modules ########################
$ErrorActionPreference = "SILENTLYCONTINUE"
import-module powervault 
import-module powerjobs 
$csv = import-csv -Path "C:\Vault Data\Script Input\CSV Input\Active_Projects-CSV.csv" -Header "Project"
######### current user for vault session #################################################################

open-vaultconnection -server "wendt-vault" -vault "wendt" -user "coolorange" -password "nhg544FK"
set-location -path "C:\Temp\"
Import-Module "C:\ProgramData\coolOrange\powerJobs\Modules\VaultSearch.psm1"
$result = @()

foreach($manager in $csv){
$projectname = $manager.Project
$srchCond_statenotempty = New-SearchCondition -PropertyName State -searchOperator:NotEmpty
$srchCond_statenotobsolete = New-SearchCondition -PropertyName State -searchOperator:NotContains -searchRuleType:Must -searchText "Obsolete"
$srchCond_statenotarchived = New-SearchCondition -PropertyName State -searchOperator:NotContains -searchRuleType:Must -searchText "Archived"
$srchCond_managernotempty = New-SearchCondition -PropertyName Manager -searchOperator:NotEmpty
$srchCond_titlenotempty = New-SearchCondition -PropertyName Title -searchOperator:NotEmpty
$srchcond_project = New-SearchCondition -PropertyName Manager -searchOperator:Contains -searchRuleType:Must -searchText $projectname

###################### Results statement ####################################################

$result += Find-VaultFiles -RecurseFolders -SearchConditions @($srchCond_statenotempty, $srchCond_statenotobsolete, $srchCond_statenotarchived, $srchCond_managernotempty, $srchCond_titlenotempty,$srchcond_project)
$result.count
}

$filename = "C:\Users\Public\Documents\Wendt-vault Active Projects.csv"
$finalfile =  "C:\Vault Data\Vault Reports\Wendt-vault Active Projects.csv"

$clear = clear-Content $filename 


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
$h23 = "Comments"
$h24 = "Cost"
$h25 = "Analysis"
$h26 = "Detailer"

$row = "" |select-Object $h1, $h2,$h3,$h4,$h5,$h6,$h7,$h8,$h9,$h10,$h11,$h12,$h13,$h14,$h15,$h16,$h17,$h18,$h19,$h20,$h21,$h22,$h23,$h24,$h25,$h26
$output = @()
################################### Table has Been Created #####################################


write-host "Retrieved $total files, Now parsing the data"
$i = 1
foreach($item in $result){

#######showing Progress ########

#$out = ""
#$i = $i + 1
#Write-Progress -Activity "Parsing Data" -Status "Progress:" -PercentComplete ($i/$items.count*100)
#######################################################################################################

$itemname = $item.Name
$vfile1 = get-vaultfiles -filename $itemname
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
$row.$h23 = $vfile1._Comments
$row.$h24 = $vfile1."Cost"
$row.$h25 = $vfile1."Analysis"
$row.$h26 = $vfile1."Detailer"

$output = $row 
$Output | Export-Csv -Append  -NoTypeInformation -Path $FileName -force
}


#$vault.PropertyService.GetProperties("File", @($item.id) , $item.PropDefId)
}
#newfile = $itemname.TOString}

}
copy-item -Path $filename -Destination $finalfile -Force

####ending file and exporting a report to csv file 

$timerend = Get-Date
$total = $result.count
$timerow.$ht4 = $total 
$timerow.$ht3 = $timerend
$timerow.$ht5 = $timing
$timeout = $timerow
$timeout | export-csv -Append -NoTypeInformation -path $timereportcsv -force
#### now will run the create pdf script 

$result = $null
$item = $null



#& "C:\ProgramData\coolOrange\powerJobs\Jobs\Wendt-Vault Electric Projects.ps1"