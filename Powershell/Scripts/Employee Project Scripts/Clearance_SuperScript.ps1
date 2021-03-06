﻿<#----------------------------------------------------------------------------------------NOTES------------------------------------------------------------------------->
Requirements before running
o365- Follow steps One, Two and Three before running: https://technet.microsoft.com/en-us/library/dn975125.aspx
Skype- Ensure you have installed the Skype For Business Management Software: SoftwareStore\Microsoft.com\Skype for Business 2015 Server

LOD 30-Mar-17 09:52: Started Script
LOD 30-Mar-17 13:49: Email Redirection functionality added
LOD 03-Apr-17 15:03: AD User disabled, Groups removed, {AD User moved to Disabled OU (COMMENTED OUT FOR FUNCTIONALITY TO BE ADDED)}
LOD 04-Apr-17 10:24: Email Hidden from exchange list
LOD 05-Apr-17 13:50: Email OffBoarding Migration Added (NEED TO TEST WITH A USER THAT HAS INFORMATION IN THEIR MAILBOX, COMMENTED OUT FOR FURTHER TESTING)
JMR 08-May-17 16:43: Added all the functionality LOD couldn't do.
JMR 08-May-17 16:43: Added error handling for Migrations and Skype removal. Removed Lachlan from Users list. (RIP) Now Version 1.7

 
<-----------------------------------------------------------------------------------------------------------------------------------------------------------------------#>
<#------------------------------------------------------------------------------------Known bugs------------------------------------------------------------------------->


 
<-----------------------------------------------------------------------------------------------------------------------------------------------------------------------#>

#Clears Console
clear

Write-Host "------------------CLEARANCE SCRIPT V1.7------------------" -foregroundcolor Blue -backgroundcolor yellow

#Clears active PSSessions and removes old Proxy Settings
Write-Host "---Clearing Active Sessions---" -foregroundcolor green -backgroundcolor gray
Get-PSSession | Remove-PSSession
Remove-Variable -name proxysettings -EA SilentlyContinue


#Initialise Variables
$correct = $False
$count = 0
$login =$false
$attemps = $false
$quit = 'n'
$complete = $false

#----------------------------------------------------------------------------------Select User START--------------------------------------------------------------------->
$input = read-host "Please Select an operating User `n 1 - James `n 2 - Keane `n"

switch ($input){

    1 {$FullyUser = "James.Rule@wisetechglobal.com"
       $3LetterUser = "CORP\jmr"
       $CORPORATE = "CORPORATE\james.rule"
       $name = "James.Rule"}

    2 {$FullyUser = "Keane.Zhang@wisetechglobal.com"
       $3LetterUser = "CORP\kez"
       $CORPORATE = "CORPORATE\keane.zhang"
       $name = "Keane.Zhang"}
    
    default {$FullyUser = $null 
             $3LetterUser = $null
             $CORPORATE = $null}
}
#----------------------------------------------------------------------------------Select User END----------------------------------------------------------------------->

#----------------------------------------------------------------------------------Gather Credentials START-------------------------------------------------------------->

    #Onpremise Cedentials

    Write-Host "---Exchange Onpremise Credentials---" -foregroundcolor green -backgroundcolor gray
    $correctOnpremise = $False
    $count = 0

    do{
        $OnPremiseCred = Get-Credential -Message 'Please provide Exchange Onpremise credentials' -UserName $3LetterUser
        
        try{
            $OnpremiseSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sydco-smai-1.wtg.zone/PowerShell/ -Authentication Kerberos -Credential $OnPremiseCred
            Import-PSSession $OnpremiseSession
            Write-Host "---Exchange On Premise Session started---" -foregroundcolor green -backgroundcolor gray
            $correctOnpremise = $true
        }

        Catch {
            Write-Host "Account details incorrect, did you forget to include domain?" -foregroundcolor green -backgroundcolor red
            $count++
            if ($count -eq 3){
                Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
                exit
            } 
        }             
    } Until($correctOnpremise -eq $true)

    #Online Cedentials

    Write-Host "---Exchange Online Credentials---" -foregroundcolor green -backgroundcolor gray
    $correctOnline =$False
    $count = 0

    do{ 
        $OnlineCred = Get-Credential -Message 'Please provide @Wisetechglobal.com credentials' -UserName $FullyUser

        try{
            Connect-MsolService -Credential $OnlineCred -Verbose
            Write-Host "---MsolService Session started---" -foregroundcolor green -backgroundcolor gray
            $proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
            $OnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $OnlineCred -Authentication Basic -AllowRedirection -SessionOption $proxysettings
            Import-PSSession $OnlineSession -Prefix ExO
            Write-Host "---Exchange Online Session started---" -foregroundcolor green -backgroundcolor gray
            $correctOnline = $true
        }

        Catch {
            Write-Host "Account details incorrect, did you forget to include domain?" -foregroundcolor green -backgroundcolor red
            $count++
            if ($count -eq 3){
                Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
                exit
            }
        }     
    } Until($correctOnline -eq $true)

    #Corporate Cedentials

    Write-Host "---CORPORATE Credentials---" -foregroundcolor green -backgroundcolor gray
    $correctCorporate =$False
    $count = 0
   
    do{
    
      $CORPORATEcred = Get-Credential -Message 'Please provide CORPORATE\ credentials' -UserName $CORPORATE
      $DomainNetBIOS = $CORPORATEcred.username.Split("{\}")[0]
      $UserName = $CORPORATEcred.username.Split("{\}")[1]
      $Password = $CORPORATEcred.GetNetworkCredential().password
      $DomainFQDN = (Get-ADDomain $DomainNetBIOS).DNSRoot

        try{           
            $DomainObj = "LDAP://" + $DomainFQDN
            $DomainBind = New-Object System.DirectoryServices.DirectoryEntry($DomainObj,$UserName,$Password)
            $DomainName = $DomainBind.distinguishedName            
            Write-Host "---CORPORATE Credentials Validated---" -foregroundcolor green -backgroundcolor gray
            $correctCorporate = $true
        }

        Catch {
            Write-Host "Account details incorrect, did you forget to include domain?" -foregroundcolor green -backgroundcolor red
            $count++
            if ($count -eq 3){
                Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
                exit
            }
        }     
    } Until($correctCorporate -eq $true)
    
    

       
#----------------------------------------------------------------------------------Gather Credentials START-------------------------------------------------------------------------------->



#----------------------------------------------------------------------------------Mutliple User Creation START---------------------------------------------------------------------------->

do{
    :loop do{

        $complete = $false

        try{

    #----------------------------------------------------------------------------------Collect User Data START----------------------------------------------------------------------------->
            Write-Host "---Retrieve User Details---" -foregroundcolor green -backgroundcolor gray
            $count = 0
            $UserWTG = $null
            $UserCorporate = $null
            $correct = $False

            do{
                $attemps = $false
                
                $Firstname = Read-Host 'What is the users first name? (Ensure use of Capitals)'
                $Lastname = Read-Host 'What is the users last name? (Ensure use of Capitals)'
                $Username = $Firstname + "." + $Lastname
                $UsernoDot = $Firstname + " " + $Lastname
                
                    try{                
                        Write-Host "Checking User Information" -foregroundcolor green                     
                        if(Get-ADUser $Username){
                            $UserWTG = Get-ADUser $Username
                            Write-Host "Found user $Username in wtg.zone" -foregroundcolor green
                        }
                        else{
                            $UserWTG = $null
                            catch
                        }
                    }

                    catch{
                        Write-Host "User not found in wtg.zone" -foregroundcolor green -backgroundcolor red
                        
                    }
                    
                    try{                    
                        If(Get-ADUser $Username  -Server corporate.cargowise.com -ea SilentlyContinue -Credential $CORPORATEcred){
                            $UserCorporate = Get-ADUser $Username  -Server corporate.cargowise.com -Credential $CORPORATEcred
                            Write-Host "Found user $Username in corporate.cargowise.com" -foregroundcolor green
                        }
                        else{
                            $UserCorporate = $null
                            catch
                        }                      
                    }

                    catch{                        
                        $UserCorporate = $null
                        Write-Host "User not found in corporate.cargowise.com" -foregroundcolor green -backgroundcolor red
                    }

                    if(($UserWTG -ne $null)){
                        $correct = $true
                    }
                    else{
                        $count++
                        Write-Host "User not found in either domain, Trying again" -foregroundcolor green -backgroundcolor red
                        if ($count -eq 3){ 
                            Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red     
                            break :loop
                        }
                    }                      
            } Until($correct -eq $true) 

    #----------------------------------------------------------------------------------Collect User Data END------------------------------------------------------------------------------------>

    #----------------------------------------------------------------------------------Redirect Email START------------------------------------------------------------------------------------->
            
            Write-Host "---Redirect User Details---" -foregroundcolor green -backgroundcolor gray
                     
            $Redirect = 'n'
            $correct = $False

            [Validateset('Y','N', IgnoreCase)]$Redirect = (Read-Host "Do emails need to be redirected? (Y/N)")
            if($Redirect -eq 'Y'){ 
                $RedirectEmail = Read-Host "Please enter User with a dot for mail to be redirected to"
                
                if($RedirectEmail -eq "me"){
                    $RedirectEmail = $name
                }
                                
                #Hide From Exchange List

                Set-RemoteMailbox -Identity $Username -HiddenFromAddressListsEnabled $true

                #Check for Valid Redirect Email Adress
                do{                    

                    try{
                        if(Get-RemoteMailbox "$RedirectEmail"){
                            Write-Host "Mailbox $RedirectEmail found" -foregroundcolor green
                            Set-RemoteMailbox $Username -PrimarySmtpAddress "$Username@wtg.zone" -EmailAddressPolicyEnabled $False
                            Set-RemoteMailbox $Username -EmailAddresses @{remove="$Username@wisetechglobal.com"} -EmailAddressPolicyEnabled $False
                            Write-Host "Primary SMPT for $Username Removed" -foregroundcolor green
                            Start-Sleep -s 10
                            Set-RemoteMailbox $RedirectEmail -EmailAddresses @{add="$Username@wisetechglobal.com"}
                            Write-Host "$Username's email added to $RedirectEmail" -foregroundcolor green
                            $correct = $true
                        }
                        else{
                            Catch
                        }                        
                    }

                    Catch {
                        Write-Host "Mailbox $RedirectEmail  not found" -foregroundcolor green -backgroundcolor red
                        $count++
                        if ($count -eq 3){ 
                            Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
                            $attemps = $true
                            break loop
                        }
                    }                     
                } Until($correct -eq $true)
            }          
                                
    #----------------------------------------------------------------------------------Redirect Email END---------------------------------------------------------------------------------------->

    #----------------------------------------------------------------------------------AD Group Start-------------------------------------------------------------------------------------------->

        Write-Host "---Removing AD User from Groups---" -foregroundcolor green -backgroundcolor gray
        $correct = $False
        $count = 0
        $Groups = $null
        $group = $null

        do{
                
                try{
                    $Groups = Get-ADPrincipalGroupMembership $Username
                    foreach($group in $Groups){
                        if($group.name -notcontains "Domain Users"){
                            Remove-ADGroupMember -Identity $group -Members $Username -Confirm:$false | Where-Object{$group.name -notcontains "Domain Users"}
                        }
                    }
                    Write-Host "AD User Removed from all Groups" -foregroundcolor green
                    $correct = $true
                }

                Catch{
                    Write-Host "AD User could not be removed from groups, Trying again" -foregroundcolor green -backgroundcolor red
                    $count++
                    if ($count -eq 3){ 
                        Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
                        $correct = $true
                    }
                }
        
        } Until($correct -eq $true)
    
    #----------------------------------------------------------------------------------AD Group End---------------------------------------------------------------------------------------------->
    
    #----------------------------------------------------------------------------------Skype Removal START--------------------------------------------------------------------------------------->

        Write-Host "---Removing User from Skype Server---" -foregroundcolor green -backgroundcolor gray

        $correct = $False
        $test = $null
        $count = 0
        
        do{   
            start-sleep -s 3 
            try{                    
                    $SkypePresent = Get-Csuser $Username -ErrorAction SilentlyContinue                                                    
                    if($SkypePresent -ne $null)
                    {                       
                    Disable-CsUser -identity "$Username@wisetechglobal.com" -Verbose
                    Write-Host "User Removed from Skype Server" -foregroundcolor green
                    $correct = $true                      
                    }              
                    
                    else { 
                    Write-Host "User not found on Skype Server" -foregroundcolor green
                    $correct = $true                 
                    }
            }
                     
            catch{            
            $count++ 
            Write-Host "Skype removal failed, waiting and trying again" -foregroundcolor green -backgroundcolor Red
             if (($count -eq 3)){ 
                Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
                $correct = $true                
            }
          }
        }Until($correct -eq $true)        

    #----------------------------------------------------------------------------------Skype Removal END------------------------------------------------------------------------------------------->

    #----------------------------------------------------------------------------------Migration START-------------------------------------------------------------------------------------------->

         Write-Host "---Starting OffBoarding Migration---" -foregroundcolor green -backgroundcolor gray

         $correct = $False
         $count = 0

        do{           
            try{
                
                $MailboxPresent = Get-ExOMailbox $UserName

                if($MailboxPresent -ne $null){
                    New-ExOMoveRequest -Identity "$Username@wtg.zone" -Outbound -RemoteTargetDatabase ExchangeDB1 -RemoteHostName mail.wtg.zone -RemoteCredential $OnPremiseCred -TargetDeliveryDomain wtg.zone 
                    Write-Host "Migration Started" -foregroundcolor green -backgroundcolor Magenta                
                    $countmigrations = 1;                    
                    while (get-ExOmoverequest $Username | fl status -ne Completed)
                    {
                       $timeelapsed = $countmigrations*10
                       Get-ExOMoveRequest $UserName | fl status
                       Start-Sleep -s -600
                       Write-Host "---OffBoarding In Progress, $timeelapsed minutes elapsed---" -foregroundcolor green -backgroundcolor Magenta
                       $countmigrations++
                    }

                    if (get-ExOmoverequest $Username | fl status -eq Completed) {
                        $correct = $true
                        get-ExOmoverequest $ Username | fl 
                    }
                }
                else {
                    Write-Host "Mailbox not found" -foregroundcolor green
                    [Validateset('Y','N', IgnoreCase)]$Continue = (Read-Host "No Mailbox to migrate, continue? (Y/N)")
                    if($Continue -eq 'Y'){ 
                    $correct = $true
                    }
                    if($Continue -eq 'N'){ 
                    $count = 'cancel'
                    catch
                }                
            }
            
        }
            Catch{
                if ($count = 'cancel'){                
                    break loop
                    }
                else{
                    Write-Host "$Username Migration Could not start, Trying again" -foregroundcolor green -backgroundcolor red
                    $count++
                    if ($count -eq 3){ 
                        Write-Host "Could not Start Off Boarding Migration" -foregroundcolor green -backgroundcolor red 
                        $correct = $true
                        break loop
                    }
                }
            }        
        } Until($correct -eq $true)       
    
   #----------------------------------------------------------------------------------Migration END--------------------------------------------------------------------------------------------->

   #---------------------------------------------------------------------------------Office365 License START-------------------------------------------------------------------------------------->

            Write-Host "---Removing Office 365 License(s)---" -foregroundcolor green -backgroundcolor gray

            $correct = $False
            $count = 0

            do{       

                try{
                    if(Get-MsolUser -UserPrincipalName "$Username@Wisetechglobal.com" | Where-Object { $_.isLicensed -eq "TRUE" }){
                        Set-MsoluserLicense -UserPrincipalName "$Username@Wisetechglobal.com" -Removelicenses  Wisetechglobal:ENTERPRISEPACK -verbose:$False
                        Write-Host "o365 License Removed" -foregroundcolor yellow -backgroundcolor blue
                        $correct = $true
                    }
                    else{
                        Write-Host "License already Removed" -foregroundcolor yellow -backgroundcolor blue
                        $correct = $true
                    }
                }

                Catch{
                    Write-Host "License Removal failed, waiting and trying again" -foregroundcolor green -backgroundcolor red
                    Start-Sleep -s 10
                    $count++
                    if ($count -eq 3){ 
                        Write-Host "License Removal failed 3 times, Manually Remove License" -foregroundcolor green -backgroundcolor red 
                        $correct = $true
                    }
                }     
            } Until($correct -eq $true)        

   #----------------------------------------------------------------------------------Office365 License END-------------------------------------------------------------------------------------->


   #----------------------------------------------------------------------------------AD Disable Start------------------------------------------------------------------------------------------>
        
         
            Write-Host "---Checking with Operator Before Proceeding---" -foregroundcolor green -backgroundcolor Magenta
            [Validateset('Y','N', IgnoreCase)]$Proceed = (Read-Host "Proceed with User Clearance? (y/n)")
            if($Proceed -eq 'Y'){
            
              
           
                  $correct = $False
                  $count = 0
    
                  do{

                    try{                    
                        Disable-ADAccount $Username
                        Write-Host "WTG AD User Disabled" -foregroundcolor green
                        $correct = $true
                    }

                    Catch{                    
                        Write-Host "Diabling WTG AD User Failed" -foregroundcolor green -backgroundcolor red
                        $count++
                        if ($count -eq 3){ 
                            Write-Host "Too many attempts - Action not completed" -foregroundcolor green -backgroundcolor red
                            $correct = $true
                        }
                    }
                    } Until ($correct -eq $True)
                    
                  $correct = $False
                  $count = 0
    
                  do{

                    try{ 
                        if($UserCorporate -ne $null){
                            Disable-ADAccount $UserCorporate  -Server corporate.cargowise.com -ea SilentlyContinue -Credential $CORPORATEcred
                            Write-Host "CORPORATE AD User Disabled" -foregroundcolor green
                            $correct = $true
                       }
                       else{$correct = $true}
                       }

                 Catch{
                        Write-Host "Diabling CORPORATE AD User Failed" -foregroundcolor green -backgroundcolor red
                        $count++
                        if ($count -eq 3){ 
                            Write-Host "Too many attempts - Action not completed" -foregroundcolor green -backgroundcolor red 
                            $correct = $true
                    }
                 }
             } Until ($correct -eq $true)

    

    #----------------------------------------------------------------------------------AD Disable END--------------------------------------------------------------------------------------------#>

    #------------------------------------------------------------------------------Move To Disabled OU START------------------------------------------------------------------------------------->
  
        $correct = $False
        $count = 0
    
        Write-Host "---Moving User To Disabled OU---" -foregroundcolor green -backgroundcolor gray

        do{      

            try{                
                move-adobject $UserWTG.DistinguishedName -TargetPath "OU=Disabled,OU=Accounts,OU=root,DC=wtg,DC=zone" -Credential $OnPremiseCred
                $correct = $true
                Write-Host "User moved to Disabled OU" -foregroundcolor green

                Write-Host "---$UsernoDot has been succesfully offboarded---" -foregroundcolor Blue -backgroundcolor yellow
                $complete = $true # LAST COMMAND              

            }

            Catch{
                Start-Sleep -s 10
                Write-Host "AD OU Move Failed, waiting and trying again" -foregroundcolor green -backgroundcolor red
                $count++
                if ($count -eq 3){
                    Write-Host "AD OU Move Failed 3 times, cancelling user creation" -foregroundcolor green -backgroundcolor red 
                    break loop
                }
            }
        } Until($correct -eq $true)
        }
        else {$complete = $true}


        
    #------------------------------------------------------------------------------Move To Disabled OU END--------------------------------------------------------------------------------------->
    
            Write-Host "---Sequence Escape---" -foregroundcolor Blue -backgroundcolor yellow 
            $complete = $true #If any of the nested Try-Catch loops breakout- the program ends. 
        }
        Catch{
            break loop
        }

    }Until($complete -eq $true)

    [Validateset('Y','N', IgnoreCase)]$quit = read-host "Would you like to run again? (Y/N)"

}Until($quit -eq 'n')

exit
      
#------------------------------------------------------------------------------------------Contributors----------------------------------------------------------------------------------------->
# James Rule
# Lachlan O'Dea
