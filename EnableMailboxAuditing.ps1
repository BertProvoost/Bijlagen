<#
.SYNOPSIS
EnableMailboxAuditingNoMFA.ps1 - Enables mailbox auditing for all users.

.DESCRIPTION 
This PowerShell script will enable mailbox auditing for all users.
#>

#...................................
# Variables
#...................................

$user = "username" #user that will execute the script                             
$cred = "folder\password.txt" #file with the encrypted password of the user 

#...................................
# Script
#...................................

#make credentials from username and encrypted password
$password = get-content $cred | convertto-securestring
$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $user, $password

#import the exchange online powershell session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credentials -Authentication Basic -AllowRedirection
Import-PSSession $Session

#enable mailbox auditing for all mailboxes
Get-Mailbox -ResultSize Unlimited | Set-Mailbox -AuditEnabled $true

#close the PSSession
Get-PSSession | Remove-PSSession

#...................................
# Finished
#...................................