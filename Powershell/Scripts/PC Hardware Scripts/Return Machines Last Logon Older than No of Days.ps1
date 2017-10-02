# Title: Return Machines Last Logon Older than a certain number of Days
# Description: Checks all Computers across every region and finds machines which haven't been logged on since a certain date
# Requirements: Run as Admin, specify a date you want to check against ($DaysInactive)
# Created by : ? I don't think I made this one. I think I found this and modified it. No credit here.

# Gets time stamps for all computers in the domain that have not logged in since after specified date
import-module activedirectory
$domain = "domain.com"
$DaysInactive = 30
$time = (Get-Date).Adddays(-($DaysInactive))

# Get all AD computers with lastLogonTimestamp less than our time
Get-ADComputer -Filter {LastLogonTimeStamp -lt $time} -Properties LastLogonTimeStamp |

# Output hostname and lastLogonTimestamp into CSV. You may need to change the filepath to be in your user folder.
Select-Object Name,@{Name="Stamp"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}} | export-csv D:\OLD_Computer_30.csv -notypeinformation
