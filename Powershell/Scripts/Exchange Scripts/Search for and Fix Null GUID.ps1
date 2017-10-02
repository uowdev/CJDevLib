# Title: Search for and Return Null GUID
# Description: This script looks for all local mailboxes with a GUID equal to 00000000-0000-0000-0000-000000000000 and repairs their GUID. This would be used when using Enable-RemoteMailbox
# Requirements: Online and Onpremise sessions (Included in script- with prefix), OnPremise and Online admin account credentials.
# Created by: James Rule (Twitter: @jgrrule Github: ClickyJimmy

get-PSSession | remove-PSSession
Clear-Console

$OnPremiseCred = Get-Credential -Message 'OnPremise'
$OnpremiseSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri #Enteryouronpremismailboxaddresshere/PowerShell/ -Authentication Kerberos -Credential $OnPremiseCred
Import-PSSession $OnpremiseSession

$OnlineCred = Get-Credential  -Message 'OnLine (ExO)'
$proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
$OnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $OnlineCred -Authentication Basic -AllowRedirection -SessionOption $proxysettings
Import-PSSession $OnlineSession -Prefix exo

Clear-Content c:\PS\UserwithNullGuid.txt
# Find Local Mailboxes with 00000 GUID
$UserList = get-remotemailbox -filter {Exchangeguid -eq "00000000-0000-0000-0000-000000000000"}
$UserList | select-object UserPrincipalName, Exchangeguid | out-file "c:\PS\OnPremMailboxesWithEmptyGUID.txt"

foreach($User in $UserList)
{
    # Check Online Mailbox Users GUID
    $OnlineUser = get-exomailbox $User.UserPrincipalName

    if ($OnlineUser -eq $null )
    {
        Write-host "$Userout Does not have an Online Mailbox"
        Add-Content c:\PS\UserwithNullGuid.txt "Userout: $userout `t doesn't have an online mailbox"
    }

    if (($User.UserPrincipalName) -eq ($OnlineUser.UserPrincipalName))
    {
        Set-remoteMailbox $User.UserPrincipalName -Exchangeguid $OnlineUser.Exchangeguid
        Write-Host ("Mailbox " + $User.UserPrincipalName + " set to " + $OnlineUser.Exchangeguid + " from " + $OnlineUser.UserPrincipalName)
        Write-Host "----------------------"  -foregroundcolor "White" -backgroundcolor "DarkCyan"
        Add-Content c:\PS\UserwithNullGuid.txt ("Mailbox " + $User.UserPrincipalName + " set to " + $OnlineUser.Exchangeguid + " from " + $OnlineUser.UserPrincipalName)
    }
    else
    {
        Write-Host ($User.UserPrincipalName + " not changed due to variation in names")
        Write-Host "----------------------"  -foregroundcolor "White" -backgroundcolor "DarkGreen"
        Add-Content c:\PS\UserwithNullGuid.txt ("Mailbox " + $User.UserPrincipalName + " not set to " + $OnlineUser.Exchangeguid + " from " + $OnlineUser.UserPrincipalName + " Due to name variation")
    }
}
