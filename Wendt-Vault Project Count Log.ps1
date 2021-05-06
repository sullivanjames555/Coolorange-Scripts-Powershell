measure-command{

import-module powervault

open-vaultconnection -server "wendt-vault" -vault "wendt" -user "coolorange" -password "nhg544FK"
Import-Module "C:\ProgramData\coolOrange\powerJobs\Modules\VaultSearch.psm1"
$csv = import-csv -Path "C:\Vault Data\Script Input\CSV Input\Active_Projects-CSV.csv" -Header "Project"
$result = @()


foreach($manager in $csv){
$projectname = $manager.Project
$srchCond_statenotempty = New-SearchCondition -PropertyName State -searchOperator:NotEmpty
$srchCond_statenotobsolete = New-SearchCondition -PropertyName State -searchOperator:NotContains -searchRuleType:Must -searchText "Obsolete"
$srchCond_statenotarchived = New-SearchCondition -PropertyName State -searchOperator:NotContains -searchRuleType:Must -searchText "Archived"
$srchCond_managernotempty = New-SearchCondition -PropertyName Manager -searchOperator:NotEmpty
$srchCond_titlenotempty = New-SearchCondition -PropertyName Title -searchOperator:NotEmpty
$srchCond_Name = New-SearchCondition -PropertyName Name -searchOperator:Contains -searchtext "*-A0.iam"
$srchCond_NoDWF = New-SearchCondition -PropertyName Name -searchOperator:NotContains -searchtext "*dwf"
$srchcond_project = New-SearchCondition -PropertyName Manager -searchOperator:Contains -searchRuleType:Must -searchText $projectname


$result += Find-VaultFiles -RecurseFolders -SearchConditions @($srchCond_statenotempty, $srchCond_statenotobsolete, $srchCond_statenotarchived, $srchCond_managernotempty, $srchCond_titlenotempty, $srchCond_Name, $srchCond_NoDWF, $srchcond_project) -LatestFilesOnly
}
########### function to process an array ############
function Get-State{
[cmdletBinding()]
param(

[array]$fileinput,

[string]$state
)


$fileinput | & {
process
{
    if(( $_.State -match $state)-and ($_.State -ne $null)){$_}
}}

} 
################ function to get mass ##############
function Get-filemass{
[cmdletBinding()]
param(

[array]$fileinput,

[string]$property
)

$massout = 0
foreach($item in $fileinput){
$mass = $item."$property"
$massout += $mass

}
$massout
}
################# end function #################

$h1 = "Project"
$h2 = "Begin Design"
$h3 = "Design In Progress"
$h4 = "Engineering Review"
$h5 = "Quick-Change"
$h6 = "Ready For Release"
$h7 = "Released"
$h8 = "Revision Pending"
$h9 = "Total Mass"
$h10 = "Mass of Begin Design"
$h11 = "Mass of Design In Progress"
$h12 = "Mass of Engineering Review"
$h13 = "Mass of Quick-Change"
$h14 = "Mass of Ready For Release"
$h15 = "Mass of Released"
$h16 = "Mass of Revision Pending"
$h17 = "Time Stamp"


$row = "" |select-Object $h1, $h2,$h3,$h4,$h5,$h6,$h7,$h8,$h9,$h10,$h11,$h12,$h13,$h14,$h15,$h16,$h17
$output = @()

$files = @()
foreach($file in $result){


$thename = $file.Name
$files += get-vaultfiles -Properties @{"Name"= "$thename"}}

$date = get-date
$projects = $files.Manager |select -unique

foreach($project in $projects){


$filename = "C:\Vault Data\Vault Reports\Project Count\$project count.csv"
 

$projectcount = $files | & {
process
{
    if($_.Manager -match $project){$_}
}
}
$mass = $projectcount.Mass
$mass | ForEach-Object -begin {$sum=0 }-process {$sum+=$_}


if($projectcount.count -gt 19){
$begindesign = Get-State -fileinput $projectcount -state "Begin Design" 
$begindesignmass = Get-filemass -fileinput $begindesign -property "mass"

$designinprogress = Get-State -fileinput $projectcount -state "Design in Progress" 
$designinprogressmass = Get-filemass -fileinput $designinprogress -property "mass"

$engineeringreview = Get-State -fileinput $projectcount -state "Engineering Review" 
$engineeringmass = Get-filemass -fileinput $engineeringreview -property "mass"

$quickchange = Get-State -fileinput $projectcount -state "Quick-Change" 
$quickchangemass = Get-filemass -fileinput $quickchange -property "mass"

$readyforrelease = Get-State -fileinput $projectcount -state "Ready For Release" 
$readyforreleasemass = Get-filemass -fileinput $readyforrelease -property "mass"

$released = Get-State -fileinput $projectcount -state "Released" 
$releasedmass = Get-filemass -fileinput $released -property "mass"

$revisionpending = Get-State -fileinput $projectcount -state "Revision Pending" 
$revisionpendingmass = Get-filemass -fileinput $revisionpending -property "mass"

$row.$h1 = $project
$row.$h2 = $begindesign.state.count
$row.$h3 = $designinprogress.state.count
$row.$h4 = ($engineeringreview.State).count
$row.$h5 = $quickchange.state.count
$row.$h6 = $readyforrelease.state.count
$row.$h7 = $released.state.count
$row.$h8 = $revisionpending.state.count
$row.$h9 = $sum
$row.$h10 = $begindesignmass
$row.$h11 = $designinprogressmass
$row.$h12 = $engineeringmass
$row.$h13 = $quickchangemass
$row.$h14 = $readyforreleasemass
$row.$h15 = $releasedmass
$row.$h16 = $revisionpendingmass
$row.$h17 = $date

$output = $row 

$Output | Export-Csv -Append  -NoTypeInformation -Path $FileName -force
}

}

} 


<#
$test = $files | & {
process
{
    if($_.Manager -match "206955"){$_}
}
}


$test2 = $test | & {
process
{
    if($_.State -match "Released"){$_}
}
}
$test3 = $test2 |select-object 'Id' -unique

#>