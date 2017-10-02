#Requirements before running
#o365- Follow steps one, two and 3 before running: https://technet.microsoft.com/en-us/library/dn975125.aspx
#Skype - https://technet.microsoft.com/en-us/library/dn362829(v=ocs.15).aspx


Write-Host "---Clearing Active Sessions---" -foregroundcolor green -backgroundcolor gray

Get-PSSession | Remove-PSSession

Remove-Variable -name proxysettings -EA SilentlyContinue

$AlertEmail = Read-Host -Prompt 'Input Email For Error and Completion Alerts' 

Write-Host "---Retrieve User Details---" -foregroundcolor green -backgroundcolor gray

$correct = $False
$count = 0

do{
$Firstname = Read-Host 'What is the users first name?'

$Lastname = Read-Host 'What is the users last name?'

$Username = $Firstname + "." + $Lastname
$UsernoDot = $Firstname + " " + $Lastname

try{

Write-Host "Check User Information" -foregroundcolor green
$User = Get-ADUser $Username 
Write-Host "Found user $User" -foregroundcolor green
$correct = $true

}
Catch {

Write-Host "User not found from AD" -foregroundcolor green -backgroundcolor red
$count++
if ($count -eq 3){ Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
exit}

} } Until($correct -eq $true) 


Write-Host "---Exchange Onpremise Credentials---" -foregroundcolor green -backgroundcolor gray

$correct = $False
$count = 0

do{

$OnPremiseCred = Get-Credential -Message 'Please provide Exchange Onpremise credentials' 

try{

$OnpremiseSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sydco-smai-1.wtg.zone/PowerShell/ -Authentication Kerberos -Credential $OnPremiseCred -Verbose
Import-PSSession $OnpremiseSession
Write-Host "---Exchange On Premise Session started---" -foregroundcolor green -backgroundcolor gray
$correct = $true

}
Catch {

Write-Host "Account details incorrect, did you forget to include domain?" -foregroundcolor green -backgroundcolor red
$count++
if ($count -eq 3){ Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
exit} 

} } Until($correct -eq $true) 


Write-Host "---Exchange Online Credentials---" -foregroundcolor green -backgroundcolor gray
$correct = $False
$count = 0

do{

$OnlineCred = Get-Credential -Message 'Please provide MsolService credentials' 

try{
Connect-MsolService -Credential $OnlineCred
Write-Host "---MsolService Session started---" -foregroundcolor green -backgroundcolor gray
$correct = $true

}
Catch {

Write-Host "Account details incorrect, did you forget to include domain?" -foregroundcolor green -backgroundcolor red
$count++
if ($count -eq 3){ Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
exit}

} } Until($correct -eq $true) 


$SendFrom = $OnlineCred.UserName -replace ("0x5c", "")

Write-Host "---Enabling Remote Mailbox---" -foregroundcolor green -backgroundcolor gray

$correct = $False
$count = 0

do{

start-sleep -s 1

try{

Enable-RemoteMailbox $UsernoDot -RemoteRoutingAddress "$Username@WiseTechGlobal.mail.onmicrosoft.com"

$correct = $true

}
Catch{

Write-Host "Remote Mailbox creation failed, waiting and trying again" -foregroundcolor green -backgroundcolor red
$count++
if ($count -eq 3){ Write-Host "Remote Mailbox creation failed 3 times, cancelling user creation" -foregroundcolor green -backgroundcolor red 
Send-MailMessage -To $AlertEmail -From $Sendfrom -Subject "New User Creation Failed" -Body "Remote Mailbox creation creation failed for $Username"
exit}

} } Until($correct -eq $true) 


Write-Host "---Assigning Office 365 License(s)---" -foregroundcolor green -backgroundcolor gray

$correct = $False
$count = 0

do{

start-sleep -s 2

try{

Set-Msoluser -UserPrincipalName "$Username@Wisetechglobal.com" -Usagelocation AU -verbose
get-Msoluser -UserPrincipalName "$Username@Wisetechglobal.com" | fl
Set-MsoluserLicense -UserPrincipalName "$Username@Wisetechglobal.com" -Addlicenses  Wisetechglobal:ENTERPRISEPACK -verbose
$correct = $true

}
Catch{

Write-Host "License Assignment failed, waiting and trying again" -foregroundcolor green -backgroundcolor red
$count++
if ($count -eq 3){ Write-Host "License Assignment failed 3 times, cancelling user creation" -foregroundcolor green -backgroundcolor red 
Send-MailMessage -To $AlertEmail -From $Sendfrom -Subject "New User Creation Failed" -Body "License Assignment failed for $Username"
exit}

} } Until($correct -eq $true) 

Write-Host "---New User Creation Complete - Emailing Alert Address---" -foregroundcolor green -backgroundcolor gray

Send-MailMessage -To $AlertEmail -From $Sendfrom -Subject "New User Creation Complete" -Body "Creating compelte for $Username"

Exit


