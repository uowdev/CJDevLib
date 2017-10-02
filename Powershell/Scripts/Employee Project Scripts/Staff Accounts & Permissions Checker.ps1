# Title: Check staff permissions
# Description: THis script takes a list of users, either from User input or from .txt, and checks the user to determine what accounts they do or do not have. It checks the following:
#   1) AD CORP + Corporate
#   2) Exchange OnPremise and O365 Online
#   3) O365 License Type + Usage Location
#   4) Skype for Business
#   5) Resource Calendar Permissions
#   6) Sharedmbx1 Folder Permissions
# Requirements: Online and Onpremise sessions (Included in script- with prefix), OnPremise and Online admin account credentials, Ran as 3 letter admin (Skype Credentials), Msolservice permissions. 
# Created by : JMR on 4/9/17
# Current Version Notes: This script can take a while to run (per User) due to the Resource + Folder Permissions. This could be taken out and turned into a separate script. 

get-pssession | remove-pssession


Clear-Host
$Users = @()
$Done = $false
$UserInput = 0
$FullOutput = $false
Get-PSSession | Remove-PSSession
Write-host "Enter full names of users to check (Firstname.LastName) OR enter "l" to load the last used list from userlist.txt `nWhen finished adding users, enter "c" to continue OR When finished enter "s" to save userlist.txt (will wipe previous entries) and then continue" -foregroundcolor "White" -backgroundcolor "Black"

while ($Done -eq $false) {   
    $UserInput = read-host
    If (($UserInput -eq "C") -or ($UserInput -eq "c"))
    {   $Done = $true   }
    If (($UserInput -eq "S") -or ($UserInput -eq "s")) {
        $Users | Out-File "C:\PS\CheckUserList.txt"
        Write-host "User List Exported"
        $Done = $true
    }
    If (($UserInput -eq "L") -or ($UserInput -eq "l")) {
        Write-host "Using Saved Users"
        $Users = Get-Content "C:\PS\CheckUserList.txt"
        $Done = $true
    }
    else {
        if ($UserInput.length -le 2)
        {break}
        else 
        {   $Users += $UserInput    }
    }
}

Write-Host "Users Selected to check: "
foreach ($User in $Users)
{   Write-Host $User   }

Write-Host "Full (f) or Basic (b) Output? Full not recommended for large user Lists and will generate a text document summary for each user"
$UserInput = read-host
if ($UserInput -eq "f") {
    $FullOutput = $true
}

#OnlineSession 
$OnlineCred = Get-Credential  -Message 'OnLine + Msol (ExO)'
$proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
$OnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $OnlineCred -Authentication Basic -AllowRedirection -SessionOption $proxysettings
Import-PSSession $OnlineSession -Prefix ExO

#MsolService
Connect-MsolService -Credential $OnlineCred -Verbose

#OnPremiseSession 
$OnPremiseCred = Get-Credential -Message 'OnPremise'
$OnpremiseSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sydco-smai-1.wtg.zone/PowerShell/ -Authentication Kerberos -Credential $OnPremiseCred
Import-PSSession $OnpremiseSession 

#CORPORATE (UNsure if necessary)
#$CORPORATECred = Get-Credential -Message 'CORPORATE'

foreach ($User in $Users) {
    #Check if User is in CORP AD and return full details
    Write-Host "------------CORP AD User Information for $User :"  -foregroundcolor "White" -backgroundcolor "DarkGreen"
    (& {If ($FullOutput -eq $true) {(Get-ADUser $User | Format-List | Out-string).Trim()} Else {(Get-ADUser $User | select-Object SamAccountName, UserPrincipalName, DistinguishedName, Enabled | Format-List | Out-string).Trim()}})

    #Check if User is in CORPORATE AD and return full details
    Write-Host "------------CORPORATE AD User Information for $User :"  -foregroundcolor "White" -backgroundcolor "DarkGreen"
    (& {If ($FullOutput -eq $true) {(Get-ADUser $User -Server corporate.cargowise.com -ea SilentlyContinue | Format-List | Out-string).Trim()} Else {(Get-ADUser $User -Server corporate.cargowise.com -ea SilentlyContinue | select-Object SamAccountName, UserPrincipalName, DistinguishedName, Enabled | Format-List | Out-string).Trim()}}) 

    #Check if User has Onpremise Local Mailbox + RemoteMailbox
    Write-Host "------------OnPremise Local Mailbox Information for $User :"  -foregroundcolor "White" -backgroundcolor "DarkGreen"
    (& {If ($FullOutput -eq $true) {(Get-Mailbox $User | Format-List | Out-string).Trim()} Else {(Get-Mailbox $User | select-object ExchangeGuid, UserPrincipalName, RecipientType | Format-List | Out-string).Trim()}})
    Write-Host "------------OnPremise Remote Mailbox User Information for $User :"  -foregroundcolor "White" -backgroundcolor "DarkGreen"
    (& {If ($FullOutput -eq $true) {(Get-RemoteMailbox $Use | Format-List | Out-string).Trim()} Else {(Get-RemoteMailbox $User | select-object ExchangeGuid, UserPrincipalName, RecipientType | Format-List | Out-string).Trim()}})

    #Check if User has Online Mailbox
    Write-Host "------------Online Mailbox Information for $User :"  -foregroundcolor "White" -backgroundcolor "DarkGreen"
    (& {If ($FullOutput -eq $true) {(Get-ExOMailbox $User | Format-List | Out-string).Trim()} Else {(Get-ExOMailbox $User | select-object ExchangeGuid, UserPrincipalName, RecipientType | Format-List | Out-string).Trim()}})

    #Check is User Has Msol License
    Write-Host "------------O365 License Information for $User :"  -foregroundcolor "White" -backgroundcolor "DarkGreen"
    (& {If ($FullOutput -eq $true) {(Get-MsolUser -UserPrincipalName $User@wisetechglobal.com | Format-List | Out-string).Trim() | Format-List} Else {(Get-MsolUser -UserPrincipalName $User@wisetechglobal.com | select-object DisplayName, SigninName, UserType, UsageLocation, IsLicensed, Licenses | Format-List | Out-string).Trim()}})

    #Check if User is in Skype
    Write-Host "------------Skype For Business Information for $User :"  -foregroundcolor "White" -backgroundcolor "DarkGreen"
    (& {If ($FullOutput -eq $true) {(Get-Csuser $User | Format-List | Out-string).Trim()} Else {(Get-Csuser $User | select-object Identity, Registrarpool, Enabled, EnterpriseVoiceEnabled, LineURI | Format-List | Out-string).Trim()}})
       
    #Check User Resource Permissions
    Write-Host "------------Resource Permission Summary for $User : (This can take some time)"  -foregroundcolor "White" -backgroundcolor "DarkGreen"
    [array]$allresources = Get-ExOMailbox -RecipientTypeDetails RoomMailbox | Select-Object -expand PrimarySmtpAddress  
    foreach ($resource in $allresources) {               
        $PermCheck = Get-ExOMailboxFolderPermission -Identity "${resource}:\calendar" -User $User@wisetechglobal.com -ErrorAction SilentlyContinue
        if ($PermCheck -ne $null) {
            $PermCheck = $PermCheck | Select-Object -expand AccessRights 
            Write-Host ("$User is $PermCheck of $resource")
        }
    }
    #Check User Sharembx1 Folder Permissions
    Write-Host "------------Sharembx1 Folders Permission Summary for $User : (This can take some time)"  -foregroundcolor "White" -backgroundcolor "DarkGreen"
    [array]$folders = Get-ExOMailbox sharedmbx1 | select-object alias | foreach-object {get-ExOmailboxfolderstatistics -identity $_.alias -Folderscope inbox| select-object -expand Identity}
    foreach ($folder in $folders) {   
        $folder = $folder | out-string 
        $folder = $folder.TrimStart("sharedMBX1 ")
        $folder = $folder.TrimEnd()         
        $PermCheck = Get-ExOMailboxFolderPermission -Identity "sharedmbx1@wtg.zone:\inbox$folder" -User $User@wisetechglobal.com -ErrorAction SilentlyContinue
        if ($PermCheck -ne $null) {
            $PermCheck = $PermCheck | Select-Object -expand AccessRights 
            Write-Host ("$User is $PermCheck of $folder")
        }
    }    
}

#Edit LOG 
# 4/9/17 (JMR): Created 