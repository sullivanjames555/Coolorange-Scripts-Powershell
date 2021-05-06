import-module Powervault
Import-module PowerJobs
$from = "wal-vault_alerts@wendtcorp.com"


$date = Get-Date -Format "dddd, MMMM d, yyyy"
$datetime = [DateTime]::Today.AddDays(-1)
$date1 = $datetime.ToString("dddd, MMMM d, yyyy")
$pdfname = "PDF $date1.csv"
$dwfname = "dwflist-$date.csv"
try{
$pdfCSVFile = import-csv -path "C:\Vault Data\Vault Reports\PDF Report\PDF $date1.csv" |  ConvertTo-Html -Fragment
}
catch {
$pdfCSVFile = "No PDF Files Created"}
try{
$dwfCSVFile = import-csv -path "C:\Vault Data\Vault Reports\DWF_Lists\$dwfname" |  ConvertTo-Html -Fragment
}
catch{
$dwfCSVFile = "No DWF Files Created, possible errors"
}



#region settings
$sender = "wal-vault_alerts@wendtcorp.com"
$receivers = @("sullivan@wendtcorp.com","nolle@wendtcorp.com","Voigt@wendtcorp.com")
$subject = "Coolorange Report"

$body = @"

<style>
body{
font-size:14px;}
</style>
CoolOrange has run the following Files Yesterday</br>
$pdfCSVFile </br>

</br>
Cool Orange Has Run These DWF's This Morning</br>
$dwfCSVFile </br>
</br>

You are receiving this from coolorange scripts at wal-vaultcron </br>
Thank you </br>
Coolorange Team </br>

"@



$smtp = "webmail.wendtcorp.com"

$attachments = @() #Fill with absolute filepaths to add attachments to the mail. E.g. @("C:\TEMP\testfile.txt", "C:\TEMP\testfile2.txt")
$useSSL = $false #Depending on the SMTP this might need to be $true
#endregion

#region send mail
foreach($address in $receivers) {
   if($attachments.Count -gt 0) {
   
        Send-MailMessage -From $sender -To $address -Attachments $attachments -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtp  -UseSsl:$useSSL        
    }
    else {
   
    Send-MailMessage -From $sender -To $address -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtp  -UseSsl:$useSSL 
    
    }
}
#endregion

