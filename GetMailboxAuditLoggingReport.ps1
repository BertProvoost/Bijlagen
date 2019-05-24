<#
.SYNOPSIS
Get-MailboxAuditLoggingReport.ps1 - Generate an Exchange Server mailbox audit logging report

.DESCRIPTION 
This PowerShell script will generate a report of the mailbox audit log entries
for all mailboxes that have auditing enabled

.OUTPUTS
Results are output to CSV/HTML that is send by email
#>

#...................................
# Variables
#...................................

$user = "username" #user that will execute the script                             
$cred = "folder\password.txt" #file with the encrypted password of the user        
$csvfile = "folder\AuditLogEntries.csv" #path to store the csv file               
$mailFrom = $username
$MailTo = "" #enter email address of recepient here                                       
$Hours = 168   

#...................................
# Script
#...................................

#make credentials from username and encrypted password
$password = get-content $cred | convertto-securestring
$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $user, $password

#import the exchange online powershell session 
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credentials -Authentication Basic -AllowRedirection
Import-PSSession $Session

#search all mailbox auditlogentries for delegate users
Write-Host "Searching mailbox auditlogentries for the last week."
$auditlogentries = @()					
$report = @()
$reportemailsubject = "Mailbox Audit Logs for the last week."
$auditlogentries = Search-MailboxAuditLog -LogonTypes Delegate -StartDate (Get-Date).AddHours(-$hours) -ShowDetails

#if there any log entries, write them to a csv file
if ($($auditlogentries.Count) -gt 0)
{
    Write-Host "Writing data to $csvfile"
    $auditlogentries | Export-CSV $csvfile -NoTypeInformation -Encoding UTF8

    #for every entry make a new object
    foreach ($entry in $auditlogentries)
    {
        $reportObj = New-Object PSObject
        $reportObj | Add-Member NoteProperty -Name "Mailbox" -Value $entry.MailboxResolvedOwnerName
        $reportObj | Add-Member NoteProperty -Name "Mailbox UPN" -Value $entry.MailboxOwnerUPN
        $reportObj | Add-Member NoteProperty -Name "Timestamp" -Value $entry.LastAccessed
        $reportObj | Add-Member NoteProperty -Name "Accessed By" -Value $entry.LogonUserDisplayName
        $reportObj | Add-Member NoteProperty -Name "Operation" -Value $entry.Operation
        $reportObj | Add-Member NoteProperty -Name "Result" -Value $entry.OperationResult
        $reportObj | Add-Member NoteProperty -Name "Folder" -Value $entry.FolderPathName
        if ($entry.ItemSubject)
        {
            $reportObj | Add-Member NoteProperty -Name "Subject Lines" -Value $entry.ItemSubject
        }
        else
        {
            $reportObj | Add-Member NoteProperty -Name "Subject Lines" -Value $entry.SourceItemSubjectsList
        }
  
        $report += $reportObj
    }

    #use the new object to make a nice HTML report for the email
	$htmlhead="<html>
				<style>
				BODY{font-family: Arial; font-size: 8pt;}
				H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
				TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
				TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
				TD{border: 1px solid #969595; padding: 5px; }
				td.pass{background: #B7EB83;}
				td.warn{background: #FFF275;}
				td.fail{background: #FF2626; color: #ffffff;}
				td.info{background: #85D4FF;}
				</style>
				<body>
                <h1>Report of mailbox audit log entries for the last week.</h1>"
	
    $htmlbody = $report | ConvertTo-Html -Fragment

	$htmltail = "</body></html>"	

	$htmlreport = "$htmlhead $htmlbody $htmltail"
    
    #send email with the HTML body and the csv file in attachment    
    Write-Host "Sending email"
    Send-MailMessage -From $mailFrom -To $MailTo -Subject $reportemailsubject `
    -BodyAsHtml $htmlreport `
    -Attachments $csvfile `
    -Credential $credentials -UseSsl -SmtpServer smtp.office365.com
    
}
else
{
    #send email without attachments
    Write-Host "Sending email"
    Send-MailMessage -From $mailFrom -To $MailTo -Subject $reportemailsubject `
    -BodyAsHtml "No audit log entries were found for the last week." `
    -Credential $credentials -UseSsl -SmtpServer smtp.office365.com
}

#close the PSSession
$session | Remove-PSSession

#...................................
# Finished
#...................................