# Title: Return Desktops CPUINFO from List of Computer Names
# Description: This script takes a list of computers, either from CSV or Array. It searches for and returns those PC's CPU information.
# Requirements: You will need to run this from an account with admin permissions for each computer.
# Created by: James Rule (@jgrrule)

#----Define CSV Location
$PClist = 'C:\PS\unknowncomputers.csv'

#----Put Computers into an array
$Computers = Get-Content $PClist

#----Define an ouput file
$CPUlist = 'C:\PS\knowncomputers.csv'

#----Read each computer
Foreach ($Computer in $Computers) {

    #Specify Command Parameters
    $Params = @{'Computername' = $Computer; 'Class' = 'Win32_Processor'}

    #Query each computer
    $Processor = (Get-WmiObject @Params).Name

    #Print the list into CSV
    [pscustomobject]@{'CPU' = $Processor; 'Computer' = $Computer} | Export-CSV -Path $CPUlist -Append -NoTypeInformation
   }
