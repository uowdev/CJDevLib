# Description: A list of commonly used scripts.  
# Created by : JMR on 21/7/17
# Current Version Notes: SCRIPT VERSION added to ensure up-to-date
# Edit LOG 
# 17/5/17 (JMR): Created 
# 28/7/17 (JMR): Added Script Format (Name : 28/7/17) to indicate last update. (Compare to version in Script folder)

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#OnlineSession : 18/9/17
$OnlineCred = Get-Credential  -Message 'OnLine (ExO)' 
$proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
$OnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $OnlineCred -Authentication Basic -AllowRedirection -SessionOption $proxysettings
Import-PSSession $OnlineSession -Prefix ExO

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#OnPremiseSession : 18/9/17
$OnPremiseCred = Get-Credential -Message 'OnPremise'
$OnpremiseSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sydco-smai-1.wtg.zone/PowerShell/ -Authentication Kerberos -Credential $OnPremiseCred
Import-PSSession $OnpremiseSession 

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#NewMoveRequest (Online -> Local) : 28/7/17
$remotecred = Get-Credential -UserName CORP\JMR -Message 'OnPremise'
New-MoveRequest -Identity "first.last@wisetechglobal.com" -Outbound -RemoteTargetDatabase ExchangeDB1 -RemoteHostName mail.wtg.zone -TargetDeliveryDomain wtg.zone -RemoteCredential $remotecred
 
#NewMoveRequest (Local -> Online) : 28/7/17
$remotecred = Get-Credential -UserName CORP\JMR -Message 'OnPremise'
new-moverequest -identity "beatriz.amadei@wisetechglobal.com" -Remote -RemoteHostName mail.wtg.zone -RemoteCredential $remotecred -TargetDeliveryDomain wisetechglobal.mail.onmicrosoft.com

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Add folder permissions for sharedmbx1 : 28/7/17
Add-MailboxFolderPermission -Identity sharedmbx1@wtg.zone -User Firstname.Lastname@wisetechglobal.com -AccessRights Reviewer
Add-MailboxFolderPermission -Identity sharedmbx1@wtg.zone:\Inbox -User Firstname.Lastname@wisetechglobal.com -AccessRights Reviewer
Add-MailboxFolderPermission -Identity "sharedmbx1@wtg.zone:\inbox\Company Secretary" -User Firstname.Lastname@wisetechglobal.com -AccessRights FolderOwner,CreateSubfolders,DeleteAllItems,EditAllItems,ReadItems,CreateItems

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Return list of folders in sharedmbx1
get-mailbox sharedmbx1 | select-object alias | foreach-object {get-mailboxfolderstatistics -identity $_.alias | select-object Identity, ItemsinFolder, FolderSize}

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Update Headshots - Online
#Will work with photos up to 648x648 pixels
#Make sure all photo files are below 100kb in size before running this script

$OnLineCred = Get-Credential -Message 'OnLine' -UserName James.rule@wisetechglobal.com

#To connect to exchange online with the ablilty to update headshots with large files
$proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/?proxyMethod=RPS -Credential $OnLineCred -Authentication Basic -AllowRedirection -SessionOption $proxysettings
Import-PSSession $Session 

Get-ChildItem "C:\photos" | Where-Object {$_.name -like "*.jpg"} | foreach {$_.name.substring(0,$_.name.length -4)} | out-file "C:\photos\UserList.txt"

$UserList = Get-Content "C:\Photos\UserList.txt"
Foreach ($User in $UserList){
    $UserWithDot = $User -replace ' ','.'
    Write-Host "Updating headshot for $User in Exchange Online."
    Set-UserPhoto $UserWithDot -PictureData ([System.IO.File]::ReadAllBytes("C:\photos\$user.jpg")) -Confirm:$False
}
Remove-Item "C:\photos\UserList.txt"

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Return all users within groups within groups 
$Groups = Get-ADGroup -Properties * -Filter * -SearchBase "CN=Content ALL,OU=Teams,OU=Distribution,OU=Groups,OU=root,DC=wtg,DC=zone" 
Foreach($G In $Groups)
{
    Write-Host $G.Name
    Write-Host "-------------"
    $G.Members
}

$Groups = Get-ADGroup -Properties * -Filter * -SearchBase "OU= ,OU= ,OU= ,OU= ,OU= ,DC= ,DC= " 
Foreach($G In $Groups)
{
    Write-Host $G.Name
    Write-Host "--------------------------"
    $G.Members
    Write-Host "-+-+-+-+-+-+-+-+-+-+-+-+-+-"
}

#Testing
 dsquery group "CN=Content ALL,OU=Teams,OU=Distribution,OU=Groups,OU=root,DC=wtg,DC=zone" | dsget group -members -expand | dsget user -samid | Export-Csv -Path C:\temp -Encoding ascii -NoTypeInformation

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Return a list of descriptions from computer names (Or swap 'Name' to 'Description' for inverse)

foreach ($computer in get-content -path C:\PS\computers.txt){
Get-ADComputer -Filter 'Name -like $computer' -Property Name,Description |
 Select -Property Name,Description}

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

