# Description: Allows easier adding of permissions for many users to a Shared Mailbox folder. Supports saving of a list of users to easily re-run for many different folders.
# Requirements: Online Exchange Session
# Created by : James Rule (@jgrrule)
# Disclaimer : This was written with a specific organisation in mind and then generalized for suitability at other organisations. Review the code before running it.

get-pssession | remove-pssession
clear

$OnlineCred = Get-Credential  -Message 'OnLine (ExO)'
$proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
$OnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $OnlineCred -Authentication Basic -AllowRedirection -SessionOption $proxysettings
Import-PSSession $OnlineSession #-Prefix ExO

$Sharedmailbox = 'sharedMBX1' #Change this to the prefix of the Mailbox you wish to change
$Domain = 'domain.com' #Change this to  your organisations domain
$Users = @()
$Done = $false
$selection = 0
$UserInput = 0
Write-host "Enter full names of users to add permission too (Firstname.LastName) OR enter "l" to load the last used list from userlist.txt `nWhen finished adding users, enter "c" to continue OR When finished enter "s" to save userlist.txt (will wipe previous entries) and then continue"

while($Done -eq $false)
    {
        $UserInput=read-host ': '
        If(($UserInput -eq "C") -or ($UserInput -eq "c"))
        {   $Done = $true   }
        If(($UserInput -eq "S") -or ($UserInput -eq "s"))
        {
            $Users | Out-File "C:\PS\UpdateFolderPermissionsUserList.txt"
            Write-host "User List Exported"
            $Done = $true
        }
        If(($UserInput -eq "L") -or ($UserInput -eq "l"))
        {
            Write-host "Using Saved Users"
            $Users = Get-Content "C:\PS\UpdateFolderPermissionsUserList.txt"
            $Done = $true
        }
        else
        {
            if($UserInput.length -le 2)
            {break}
            else
            {   $Users += $UserInput    }
        }
    }

Write-Host "Users Selected: "
foreach($User in $Users)
{   Write-Host $User   }

write-host "`nPlease wait while Mailbox folders are retrieved`n" -foregroundcolor green

[array]$folders = get-mailbox $Sharedmailbox | select-object alias | foreach-object {get-mailboxfolderstatistics -identity $_.alias -Folderscope inbox| select-object -expand Identity}

foreach($folder in $folders)
    {
        $folder = $folder | out-string
        $folder = $folder.TrimStart("$Sharedmailbox ")
        write-host $selection. $folder `n
        $selection++
    }

$UInput=read-host 'Enter corrosponding number of folder to add permission too'
$folder = $folders[$UInput]
$folder = $folder | out-string
$folder = $folder.TrimStart("$Sharedmailbox ")
$folder = $folder.TrimEnd()
write-host("`n$folder selected`n") -foregroundcolor green

$selection = 0

$permissions = @(
            [PSCustomObject]@{ID = 0; Name = "Author"; Description = 'Author:               CreateItems, DeleteOwnedItems, EditOwnedItems, FolderVisible, ReadItems'}
            [PSCustomObject]@{ID = 1; Name = "Contributor"; Description = "Contributor:          CreateItems, FolderVisible"}
            [PSCustomObject]@{ID = 2; Name = "Editor"; Description = "Editor:               CreateItems, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderVisible, ReadItems"}
            [PSCustomObject]@{ID = 3; Name = "None"; Description = "None:                 FolderVisible"}
            [PSCustomObject]@{ID = 4; Name = "NonEditingAuthor"; Description = "NonEditingAuthor:     CreateItems, FolderVisible, ReadItems"}
            [PSCustomObject]@{ID = 5; Name = "Owner"; Description = "Owner:                CreateItems, CreateSubfolders, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderContact, FolderOwner, FolderVisible, ReadItems"}
            [PSCustomObject]@{ID = 6; Name = "PublishingEditor"; Description = "PublishingEditor:     CreateItems, CreateSubfolders, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderVisible, ReadItems"}
            [PSCustomObject]@{ID = 7; Name = "PublishingAuthor"; Description = "PublishingAuthor:     CreateItems, CreateSubfolders, DeleteOwnedItems, EditOwnedItems, FolderVisible, ReadItems"}
            [PSCustomObject]@{ID = 8; Name = "Reviewer"; Description = "Reviewer:             FolderVisible, ReadItems"}
            [PSCustomObject]@{ID = 9; Name = "Remove"; Description = "Remove:               Wipe User Access to specified folder (Will not remove \inbox permissions"}
)

foreach ($permission in $permissions)
{
    write-host $selection. $permission.Description `n
    $selection++
}

$UInput=read-host 'Enter corrosponding number of permission to use'
$permission=$permissions | Where {$_.ID -eq $UInput} | select -expand Name
Write-Host ("`n$permission selected`n")

if($permission -eq "Remove")
    {
        foreach ($User in $Users)
        {
            Write-Host("Removing Permissions for $User on $folder")
            Remove-MailboxFolderPermission  -Identity "$Sharedmailbox@$domain:\inbox$folder" -User $User@$domain
            Write-Host ("Permission summary for user on folder:") -foregroundcolor green
            Get-MailboxFolderPermission -Identity "$Sharedmailbox@$domain:\inbox$folder" -User $User@$domain
        }
    }
else
{
    foreach ($User in $Users)
    {
        Write-Host("`nUpdating Permissions for $User on $folder")
        $currentpermissions = Get-MailboxFolderPermission -Identity "$Sharedmailbox@$domain:\inbox$folder" -User $User@$domain -ErrorAction SilentlyContinue

        if($currentpermissions -ne $null)
        {
            $perms = $currentpermissions | select -expand AccessRights
            Write-Host "$User currently has" -foregroundcolor green -NoNewline
            Write-Host " $perms " -foregroundcolor Magenta -NoNewline
            Write-Host "rights for this folder, remove current permissions and overright with" -foregroundcolor green -NoNewline
            Write-Host " $permission " -foregroundcolor Magenta -NoNewline
            Write-Host "rights? Y/N : " -foregroundcolor green -NoNewline
            $UserInput = read-Host
            If(($UserInput -eq "Y") -or ($UserInput -eq "Y") -or ($UserInput -eq "Yes") -or ($UserInput -eq "yes"))
            {
                Write-Host("`nRemoving Permissions for $User on $folder")
                Remove-MailboxFolderPermission  -Identity "$Sharedmailbox@$domain:\inbox$folder" -User $User@$domain
                Write-Host("`nAdding $permission Permissions for $User on $folder")
                Add-MailboxFolderPermission -Identity $Sharedmailbox@$domain -User $User@$domain -AccessRights Reviewer #Required to view the folder
                Add-MailboxFolderPermission -Identity $Sharedmailbox@$domain:\inbox -User $User@$domain -AccessRights Reviewer #Required to view the folder
                Add-MailboxFolderPermission -Identity "$Sharedmailbox@$domain:\inbox$folder" -User $User@$domain -AccessRights $permission
                Write-Host ("Permission summary for user on folder :") -foregroundcolor green
                Get-MailboxFolderPermission -Identity "$Sharedmailbox@$domain:\inbox$folder" -User $User@$domain
            }
            else
            {
                Write-Host ("`nNothing has changed. Permission summary for user on folder :") -foregroundcolor green
                Get-MailboxFolderPermission -Identity "$Sharedmailbox@$domain:\inbox$folder" -User $User@$domain
            }

        }
        else
        {
            Write-Host("`nAdding $permission Permissions for $User on $folder")
            Add-MailboxFolderPermission -Identity $Sharedmailbox@$domain -User $User@$domain -AccessRights Reviewer #Required to view the folder
            Add-MailboxFolderPermission -Identity $Sharedmailbox@$domain:\inbox -User $User@$domain -AccessRights Reviewer #Required to view the folder
            Add-MailboxFolderPermission -Identity "$Sharedmailbox@$domain:\inbox$folder" -User $User@$domain -AccessRights $permission
            Write-Host ("`nPermission summary for user on folder :`n") -foregroundcolor green
            Get-MailboxFolderPermission -Identity "$Sharedmailbox@$domain:\inbox$folder" -User $User@$domain
        }
    }
}
