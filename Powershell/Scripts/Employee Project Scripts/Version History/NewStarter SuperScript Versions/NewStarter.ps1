
Write-Host "---Exchange Onpremise Credentials---" -foregroundcolor green -backgroundcolor gray

$correct = $False
do{

$OnPremiseCred = Get-Credential -Message 'Please provide Exchange Onpremise credentials' 

try{

$OnpremiseSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sydco-smai-1.wtg.zone/PowerShell/ -Authentication Kerberos -Credential $OnPremiseCred
Import-PSSession $OnpremiseSession

Write-Host "---Exchange On Premise Session started---" -foregroundcolor green -backgroundcolor gray

$correct = $true

}
Catch {
Write-Host "Account details incorrect, did you forget to include domain?" -foregroundcolor green -backgroundcolor red
}
} Until($correct -eq $true) 


Write-Host "---Exchange Online Credentials---" -foregroundcolor green -backgroundcolor gray
$correct = $False
do{

$OnlineCred = Get-Credential -Message 'Please provide Exchange Online credentials' 

try{

$proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
$OnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $OnlineCred -Authentication Basic -AllowRedirection -SessionOption $proxysettings
Import-PSSession $OnlineSession

Write-Host "---Exchange Online Session started---" -foregroundcolor green -backgroundcolor gray

$correct = $true

}
Catch {
Write-Host "Account details incorrect, make sure it is a fully qualified account name" -foregroundcolor green -backgroundcolor red
}
} Until($correct -eq $true) 

$AlertEmail = Read-Host -Prompt 'Input Email For Error and Completion Alerts' 


$Firstname = Read-Host 'What is the new users FirstName?'
$Firstname = "$Firstname*"

$Lastname = Read-Host 'What is the new users Lastname?'
$Lastname = "$Lastname*"


Write-Host "Check User Information" -foregroundcolor green


$User = Get-ADUser -Filter {(Name -Like $Firstname) -And (Surname -Like $Lastname)}

Write-Host "Found user" $User -foregroundcolor green


#Connect to on Premise Exchange 

#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sydco-smai-1.wtg.zone/PowerShell/ -Authentication Kerberos -Credential $OnPremiseCred
 
#Import-PSSession $Session

#Create new On-Premise Mailbox 