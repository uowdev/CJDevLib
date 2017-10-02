# Title: Return Machines Last Logon Older than Specify Date
# Description: Checks all Computers across every region and finds machines which haven't been logged on since a certain date
# Requirements: You will need to run this from an account with admin permissions for each computer.
# Created by: James Rule (@jgrrule)

#----Clear Console
clear

#----Specify Time
$time = get-date ('15/04/2017')

#Specify an OU to collect computers, you could also create another CSV for this.
$OU = 'OU=Sales,OU=Clients,OU=Computers,OU=root,DC=domain,DC=domain'

#----Define an ouput file
$Out = 'E:\aspac_computers_lastlogon_gt_1month.csv'

$ComComputers = Get-ADObject -SearchBase $OU -Filter {LastLogonTimeStamp -lt $time} -Properties LastLogonTimeStamp

Foreach ($Computer in $ComComputers) {
    #Print the list into CSV
    $Time = [DateTime]$Computer.LastLogonTimeStamp
    [pscustomobject]@{'Computer' = $Computer.Name; 'LastLogon' = $Time} | Export-CSV -Path $Out -Append -NoTypeInformation
   }
