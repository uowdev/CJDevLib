# Description: Allows easier adding of many users to specific resources
# Requirements: Online Exchange Session
# Created by : James Rule (@jgrrule)
# Disclaimer : This was written with a specific organisation in mind and then generalized for suitability at other organisations. Review the code before running it.

get-pssession | remove-pssession
clear

$OnlineCred = Get-Credential  -Message 'OnLine (ExO)'
$proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
$OnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $OnlineCred -Authentication Basic -AllowRedirection -SessionOption $proxysettings
Import-PSSession $OnlineSession #-Prefix ExO

$Domain = 'domain.com' #Change this to your organisations domain
$Users = @()
$Done = $false
$selection = 0
$UserInput = 0
Write-host "Enter full names of users to add permission too (Firstname.LastName) OR enter "l" to load the last used list from userlist.txt `nWhen finished adding users, enter "c" to continue OR When finished enter "s" to save userlist.txt (will wipe previous entries) and then continue"

while($Done -eq $false)
    {
        $UserInput=read-host
        If(($UserInput -eq "C") -or ($UserInput -eq "c"))
        {   $Done = $true   }
        If(($UserInput -eq "S") -or ($UserInput -eq "s"))
        {
            $Users | Out-File "C:\PS\UpdateResourcePermissionsUserList.txt"
            Write-host "User List Exported"
            $Done = $true
        }
        If(($UserInput -eq "L") -or ($UserInput -eq "l"))
        {
            Write-host "Using Saved Users"
            $Users = Get-Content "C:\PS\UpdateResourcePermissionsUserList.txt"
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

$Done = $false
$Allflag = $false
Write-Host "`nPlease wait while Resources are retrieved`n" -foregroundcolor green
[array]$allresources = Get-Mailbox -RecipientTypeDetails RoomMailbox | select -expand PrimarySmtpAddress
$resources = @()

foreach($resource in $allresources)
    {
        write-host $selection. $resource `n
        $selection++
    }

$ResourceInput= @()
Write-host "Enter corrosponding number of resource to add permissions too, when complete enter 'c', for all resources (To clear or check), enter 'all' (Proceed with caution)"

while($Done -eq $false)
    {
        $RInput=read-host
        If(($RInput -eq "C") -or ($RInput -eq "c"))
        {   $Done = $true   }
        If(($RInput -eq "All") -or ($RInput -eq "all"))
        {
            $Allflag = $true
            $resources = $allresources
            $Done = $true
        }
        else
        {   $resources += $allresources[$RInput]    }
    }

Write-Host "Resources Selected: " -foregroundcolor green
foreach ($resource in $resources)
    {
      Write-Host $resource
    }

$selection = 0
$permissions = @(
            [PSCustomObject]@{ID = 0; Name = "Author"; Description = "Author:               CreateItems, DeleteOwnedItems, EditOwnedItems, FolderVisible, ReadItems"}
            [PSCustomObject]@{ID = 1; Name = "Contributor"; Description = "Contributor:          CreateItems, FolderVisible"}
            [PSCustomObject]@{ID = 2; Name = "Editor"; Description = "Editor:               CreateItems, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderVisible, ReadItems"}
            [PSCustomObject]@{ID = 3; Name = "None"; Description = "None:                 FolderVisible"}
            [PSCustomObject]@{ID = 4; Name = "NonEditingAuthor"; Description = "NonEditingAuthor:     CreateItems, FolderVisible, ReadItems"}
            [PSCustomObject]@{ID = 5; Name = "Owner"; Description = "Owner:                CreateItems, CreateSubfolders, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderContact, FolderOwner, FolderVisible, ReadItems"}
            [PSCustomObject]@{ID = 6; Name = "PublishingEditor"; Description = "PublishingEditor:     CreateItems, CreateSubfolders, DeleteAllItems, DeleteOwnedItems, EditAllItems, EditOwnedItems, FolderVisible, ReadItems"}
            [PSCustomObject]@{ID = 7; Name = "PublishingAuthor"; Description = "PublishingAuthor:     CreateItems, CreateSubfolders, DeleteOwnedItems, EditOwnedItems, FolderVisible, ReadItems"}
            [PSCustomObject]@{ID = 8; Name = "Reviewer"; Description = "Reviewer:             FolderVisible, ReadItems"}
            [PSCustomObject]@{ID = 9; Name = "Remove"; Description = "Remove:               Wipe User Access to specified resource"}
            [PSCustomObject]@{ID = 10; Name = "Check"; Description = "Check:               Check users Current Access to specified resource"}
)

foreach ($permission in $permissions)
{
    write-host $selection. $permission.Description `n
        $selection++
}

$UInput=read-host 'Enter corrosponding number of permission to use for these users on these resources'
$permission=$permissions | Where {$_.ID -eq $UInput} | select -expand Name
Write-Host ("`n$permission selected`n")

if($permission -eq "Remove")
    {
        foreach ($User in $Users)
        {
            foreach($resource in $resources)
            {
                Write-Host("Removing Permissions for $User on $resource")
                Remove-MailboxFolderPermission -Identity "${resource}:\calendar" -User $User@$Domain
                Write-Host ("Permission summary for user on $resource :") -foregroundcolor green
                Get-MailboxFolderPermission -Identity "${resource}:\calendar" -User $User@$Domain
            }
        }
    }

if($permission -eq "Check")
    {
        foreach ($User in $Users)
        {
            foreach($resource in $resources)
            {
                Write-Host ("Permission summary for user on $resource :") -foregroundcolor green
                Get-MailboxFolderPermission -Identity "${resource}:\calendar" -User $User@$Domain
            }
        }
    }
else
{
    foreach ($User in $Users)
    {
        if($Allflag -eq $true)
        {
            Write-host "You are about to grant $User $permission access to ALL RESOURCES, YOU PROBABLY SHOULDN'T DO THIS, PLEASE CTRL+C TO CANCEL NOW OR WAIT 15 SECONDS TO PROCEED"
            Start-sleep -s 5
            Write-host "NO PERMISSIONS APPLIED"
            exit #I have left this in so juniors won't accidently do this.
        }

        Write-Host("`nUpdating Permissions for $User")

        foreach ($resource in $resources)
        {
            $currentpermissions = Get-MailboxFolderPermission -Identity "${resource}:\calendar" -User $User@$Domain -ErrorAction SilentlyContinue

            if($currentpermissions -ne $null)
            {
                $UserInput = 0
                $perms = $currentpermissions | select -expand AccessRights
                Write-Host "$User currently has" -foregroundcolor green -NoNewline
                Write-Host " $perms " -foregroundcolor Magenta -NoNewline
                Write-Host "rights for this resource, remove current permissions and overright with" -foregroundcolor green -NoNewline
                Write-Host " $permission " -foregroundcolor Magenta -NoNewline
                Write-Host "rights? Y/N : " -foregroundcolor green -NoNewline
                $UserInput = read-Host
                If(($UserInput -eq "Y") -or ($UserInput -eq "Y") -or ($UserInput -eq "Yes") -or ($UserInput -eq "yes"))
                {
                    Write-Host("`nRemoving Permissions for $User on $resource")
                    Remove-MailboxFolderPermission -Identity "${resource}:\calendar" -User $User@$Domain
                    Write-Host("`nAdding $permission Permissions for $User on $resource")
                    Add-MailboxFolderPermission -Identity "${resource}:\calendar" -User $User@$Domain -AccessRights $permission
                    Write-Host ("`nPermission summary for user on resource :") -foregroundcolor green
                    Get-MailboxFolderPermission -Identity "${resource}:\calendar" -User $User@$Domain
                }
                else
                {
                    Write-Host ("`nNothing has changed. Permission summary for user on resource :") -foregroundcolor green
                    Get-MailboxFolderPermission -Identity "${resource}:\calendar" -User $User@$Domain
                }
            }
            else
            {
                Write-Host("`nAdding $permission Permissions for $User on $resource")
                Add-MailboxFolderPermission -Identity "${resource}:\calendar" -User $User@$Domain -AccessRights $permission
                Write-Host ("`nPermission summary for user on resource :`n") -foregroundcolor green
                Get-MailboxFolderPermission -Identity "${resource}:\calendar" -User $User@$Domain
            }
        }
    }
}
