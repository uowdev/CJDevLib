# Description: This can be used at the beginning of a script to capture Names with the option to Save or Load to/from a .txt document for later use or for batch operations
# Requirements: A C:/PS Folder OR set a new location
# Created by: James Rule (@jgrrule)
# Note: This is a fairly simple, but useful, script. Written for $Users in its current state but could be used for anything that needs a list of inputs. It's worth noting that there are many other ways to do this- but this format alllows for beginners to easily read and understand what's happening.

$Users = @() #The array which will contain all the inputs by the end.
$Done = $false
$UserInput = 0
Write-host "Enter full names of users (Firstname.LastName) OR enter "l" to load the last used list from C:\PS\Userlist.txt `nWhen finished adding users, enter "c" to continue OR When finished enter "s" to save userlist.txt (will wipe previous entries) and then continue"

while($Done -eq $false)
    {
        $UserInput=read-host
        If(($UserInput -eq "C") -or ($UserInput -eq "c"))
        {   $Done = $true   }
        If(($UserInput -eq "S") -or ($UserInput -eq "s"))
        {
            $Users | Out-File "C:\PS\Userlist.txt"
            Write-host "User List Exported"
            $Done = $true
        }
        If(($UserInput -eq "L") -or ($UserInput -eq "l"))
        {
            Write-host "Using Saved Users"
            $Users = Get-Content "C:\PS\Userlist.txt"
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
#This is where you can now use the $User array to do whatever you want
