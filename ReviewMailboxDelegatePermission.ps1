<#
.SYNOPSIS
ReviewMailboxDelegatePermission.ps1 - generate a delegate permissions report.

.DESCRIPTION 
This PowerShell script will generate a report with all the delegate permissions on users mailboxes.

.OUTPUTS
Results are output to a CSV file that is send by email.
#>

#...................................
# Variables
#...................................

$user = "username" #user that will execute the script
$cred = "folder\password.txt" #file with the encrypted password of the user                                                                                         
$MailTo = "" #email address to mail the report to
$UserDelegatesPath = "folder\MailboxDelegatePermissions.csv" #path to store the csv file
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

#select all users and make an empty variable
$AllUsers = Get-Mailbox -ResultSize Unlimited | select UserPrincipalName
$UserDelegates = @()

#go over each user and check the delegate permissions
foreach ($User in $allUsers)
{
    Write-Host "Checking for delegate permissions for user: " $User.UserPrincipalName;
    $UserDelegates += Get-mailbox -Identity $User.UserPrincipalName | Get-MailboxPermission -ErrorAction SilentlyContinue | select runspaceId, AccessRights, Deny, InheritanceType, User, @{N=’OnMailbox’; E='Identity'}, IsInherited, IsValid, ObjectState | Where-Object {($_.IsInherited -ne "True") -and ($_.User -notlike "*SELF*")}
}

#store the data in a CSV file to use as an attachment in the email
$UserDelegates | Export-Csv MailboxDelegatePermissions.csv -Delimiter ";"

##send an email with the report
$body = "<html><h1>Report of mailbox delegate permissions of last week.</h1>"

if($UserDelegates){
    $body += "<p>The mailbox delegations report can be found in the attached CSV file: MailboxDelegatePermissions.csv.</p></html>"
    Send-MailMessage -From $mailFrom -To $MailTo -Subject 'Review mailbox delegate permissions' -BodyAsHtml $body -Credential $credentials -Attachments $UserDelegatesPath  -UseSsl -SmtpServer smtp.office365.com
    Write-Host "sending email to $MailTo"
}
else{
    $body += "<p>There are no mailbox delegation permissions.</p></html>"
    Send-MailMessage -From $mailFrom -To $MailTo -Subject 'Review mailbox delegate permissions' -BodyAsHtml $body -Credential $credentials -UseSsl -SmtpServer smtp.office365.com
    Write-Host "sending email to $MailTo"
}

#close the PSSession
Get-PSSession | Remove-PSSession

#...................................
# Finished
#...................................
