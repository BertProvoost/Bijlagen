<#
.SYNOPSIS
ReviewMailboxForwarding.ps1 - generate a forwarding report.

.DESCRIPTION 
This PowerShell script will generate a report of the forwarding rules made by users.

.OUTPUTS
Results are output to CSV files that are send by email
#>

#...................................
# Variables
#...................................

$user = "username" #user that will execute the script
$cred = "folder\password.txt" #file with the encrypted password of the user                                                                                         
$MailTo = "" #email address to mail the report to
$UserInboxRulePath = "folder\MailboxForwardingRules.csv" #path to store the csv file
$SMTPForwardingPath = "folder\MailboxSMTPForwarding.csv" #path to store the csv file
$mailFrom = $user

#...................................
# Script
#...................................

#make credentials from username and encrypted password
$password = get-content $cred | convertto-securestring
$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $user, $password

#import the exchange online powershell session 
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credentials -Authentication Basic -AllowRedirection
Import-PSSession $Session

#select all users and make empty variables
$AllUsers = Get-Mailbox -ResultSize Unlimited | select UserPrincipalName
$UserInboxRules = @()
$SMTPForwarding = @()

#go over each user and check the forwarding rules & SMTP forwarding rules
foreach ($User in $allUsers)
{
    Write-Host "Checking inbox rules for user: " $User.UserPrincipalName;
    $UserInboxRules += Get-InboxRule -Mailbox $User.UserPrincipalname -ErrorAction SilentlyContinue | Select @{n='Sender';e={$User.UserPrincipalname}}, Name, Description, Enabled, Priority, ForwardTo, ForwardAsAttachmentTo, RedirectTo, DeleteMessage | Where-Object {($_.ForwardTo -ne $null) -or ($_.ForwardAsAttachmentTo -ne $null) -or ($_.RedirectsTo -ne $null)} 
    $SMTPForwarding += Get-Mailbox $User.UserPrincipalname -ResultSize Unlimited -ErrorAction SilentlyContinue | select DisplayName,ForwardingSMTPAddress | where {$_.ForwardingSMTPAddress -ne $null}
}

#store the data in CSV files to use as an attachment in the email
$UserInboxRules | Export-Csv MailboxForwardingRules.csv -Delimiter ";"
$SMTPForwarding | Export-Csv MailboxSMTPForwarding.csv -Delimiter ";"

#make the email
$body = "<html><h1>Report of mailbox forwarding rules log entries for the last week.</h1>"
$attachments = @()
$attachmentsnumber = 0
if($UserInboxRules){
    $attachments += $UserInboxRulePath
    $attachmentsnumber ++
    $body += "<p>The user inbox rules can be found in the attached CSV file: MailboxForwardingRules.csv.</p>"
}
else{
    $body += "<p>There are no user forwarding via forwarding rules.</p>"
}
if($SMTPForwarding){
    $attachments += $SMTPForwardingPath
    $attachmentsnumber ++
    $body += "<p>The user SMTP forwarding rules can be found in the attached CSV file: MailboxSMTPForwarding.csv.</p>"
}
else{
    $body += "<p>There are no users that forwarding emails with SMTP forwarding.</p>"
}

#send the email without attachments if there are none and with attachments if there are.
if(!$attachments){
    Send-MailMessage -From $mailFrom -To $MailTo -Subject 'Review mailbox forwarding rules' -BodyAsHtml $body -Credential $credentials -UseSsl -SmtpServer smtp.office365.com
    Write-Host "sending email to $MailTo"
}
else{
    Send-MailMessage -From $mailFrom -To $MailTo -Subject 'Review mailbox forwarding rules' -BodyAsHtml $body -Credential $credentials -Attachments $attachments  -UseSsl -SmtpServer smtp.office365.com
    Write-Host "sending email to $MailTo"
}

#close the PSSession
Get-PSSession | Remove-PSSession

#...................................
# Finished
#...................................
