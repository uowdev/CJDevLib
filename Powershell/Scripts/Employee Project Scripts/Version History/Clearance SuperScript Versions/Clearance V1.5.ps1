<#----------------------------------------------------------------------------------------NOTES------------------------------------------------------------------------->
Requirements before running
o365- Follow steps One, Two and Three before running: https://technet.microsoft.com/en-us/library/dn975125.aspx
Skype- Ensure you have installed the Skype For Business Management Software: SoftwareStore\Microsoft.com\Skype for Business 2015 Server

LOD 30-Mar-17 09:52: Started Script
LOD 30-Mar-17 13:49: Email Redirection functionality added
LOD 03-Apr-17 15:03: AD User disabled, Groups removed, {AD User moved to Disabled OU (COMMENTED OUT FOR FUNCTIONALITY TO BE ADDED)}
LOD 04-Apr-17 10:24: Email Hidden from exchange list
LOD 05-Apr-17 13:50: Email OffBoarding Migration Added (NEED TO TEST WITH A USER THAT HAS INFORMATION IN THEIR MAILBOX, COMMENTED OUT FOR FURTHER TESTING)
 
<-----------------------------------------------------------------------------------------------------------------------------------------------------------------------#>
<#------------------------------------------------------------------------------------Known bugs------------------------------------------------------------------------->


 
<-----------------------------------------------------------------------------------------------------------------------------------------------------------------------#>

#Clears Console
clear

Write-Host "------------------CLEARANCE SCRIPT V1.5------------------" -foregroundcolor Blue -backgroundcolor yellow

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
$input = read-host "Please Select an operating User `n 1 - James `n 2 - Lachlan `n 3 - Keane `n"

switch ($input){

    1 {$FullyUser = "James.Rule@wisetechglobal.com"
       $3LetterUser = "CORP\jmr"
       $CORPORATE = "CORPORATE\james.rule"
       $name = "James.Rule"}

    2 {$FullyUser = "Lachlan.Odea@wisetechglobal.com" 
       $3LetterUser = "CORP\lod"
       $CORPORATE = "CORPORATE\lachlan.odea"
       $name = "Lachlan.Odea"}

    3 {$FullyUser = "Keane.Zhang@wisetechglobal.com"
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

    <#Write-Host "---CORPORATE Credentials---" -foregroundcolor green -backgroundcolor gray
    $correctCorporate =$False
    $count = 0
    $validName= Get-ADUser $name
    $validPath = Get-ADUser $name -Server corporate.cargowise.com -Properties distinguishedname,cn | select @{n='ParentContainer';e={$_.distinguishedname -replace "CN=$($_.cn),",''}}
    Get-PSSession | Remove-PSSession

    Remove-Variable -name proxysettings -EA SilentlyContinue
    do{
 
        $CORPORATEcred = Get-Credential -Message 'Please provide CORPORATE\ credentials' -UserName $CORPORATE

        try{

            move-adobject $validName.distinguishedName -server corporate.cargowise.com -TargetPath $validPath -Credential $CORPORATEcred
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
    } Until($correctCorporate -eq $true)#>
    
    

       
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

                        #Get-PSSession | Remove-PSSession

                    try{
                    
                        If(Get-ADUser $Username  -Server corporate.cargowise.com -ea SilentlyContinue){
                            $UserCorporate = Get-ADUser $Username  -Server corporate.cargowise.com
                            Write-Host "Found user $Username in corporate.cargowise.com" -foregroundcolor green
                        }

                        else{
                            $UserCorporate = $null
                            catch
                        }

                      
                    }

                    catch{
                        
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
            
            Get-PSSession | Remove-PSSession
            Remove-Variable -name proxysettings -EA SilentlyContinue

            $Redirect = 'n'
            $correct = $False

            [Validateset('Y','N', IgnoreCase)]$Redirect = (Read-Host "Do emails need to be redirected? (Y/N)")
            if($Redirect -eq 'Y'){ 
                $RedirectEmail = Read-Host "Please enter User with a dot for mail to be redirected to"
                
                if($RedirectEmail -eq "me"){
                    $RedirectEmail = $name
                }


                #Connect to OnPremise
                $OnpremiseSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sydco-smai-1.wtg.zone/PowerShell/ -Authentication Kerberos -Credential $OnPremiseCred
                Import-PSSession $OnpremiseSession
                
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

    <#----------------------------------------------------------------------------------Migration START---------------------------------------------------------------------------------------->

        $correct = $False
        $count = 0

        Get-PSSession | Remove-PSSession
        Remove-Variable -name proxysettings -EA SilentlyContinue

        $proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
        $OnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $OnlineCred -Authentication Basic -AllowRedirection -SessionOption $proxysettings
        Import-PSSession $OnlineSession

        Write-Host "---Starting OffBoarding Migration---" -foregroundcolor green -backgroundcolor gray
        if(test-path C:\Batch){
            Remove-Item C:\Batch -Recurse
        }
        mkdir C:\Batch -Verbose:$False
        Add-Content C:\Batch\OffBoardingEmail.csv "EmailAddress"
        Add-Content C:\Batch\OffBoardingEmail.csv "$Username@wtg.zone"

        do{
           
            try{
                New-MigrationBatch -Name "$Username.2OnPremise" -TargetEndpoint mail.wtg.zone2 -TargetDeliveryDomain wtg.zone -TargetDatabase 'ExchangeDB1' -CSVData ([System.IO.File]::ReadAllBytes("C:\Batch\OffBoardingEmail.csv")) -NotificationEmails $FullyUser -AutoStart -AutoComplete -TargetArchiveDatabases "ExchangeDB1" -BadItemLimit 50 -LargeItemLimit 50
                
                #Start-MigrationBatch "$Username.2OnPremise" 
                $correct = $true
            }

            Catch{
                Write-Host "$Username Migration Could not start, Trying again" -foregroundcolor green -backgroundcolor red
                $count++
                if ($count -eq 3){ 
                    Write-Host "Could not Start Off Boarding Migration" -foregroundcolor green -backgroundcolor red 
                    $correct = $true
                }
            }
        
        } Until($correct -eq $true)

        Remove-Item C:\Batch -Recurse
    #----------------------------------------------------------------------------------Migration END------------------------------------------------------------------------------------------>#>

    #----------------------------------------------------------------------------------AD Disable Start------------------------------------------------------------------------------------------>
        
        $correct = $False
        $count = 0

        Write-Host "---Disabling AD User---" -foregroundcolor green -backgroundcolor gray

        do{
                
                try{
                    Disable-ADAccount $Username
                    Write-Host "AD User Disabled" -foregroundcolor green
                    $correct = $true
                }

                Catch{
                    Write-Host "Diabling User Failed" -foregroundcolor green -backgroundcolor red
                    $count++
                    if ($count -eq 3){ 
                        Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
                        $correct = $true
                    }
                }

        } Until($correct -eq $true)



    #----------------------------------------------------------------------------------AD Disable END-------------------------------------------------------------------------------------------->

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
            
            if(Get-CsUser $Username -Verbose:$False){
                Disable-CsUser $Username
                Write-Host "User Removed From Skype Server" -foregroundcolor green
                $correct = $true
            }
            else{
                Write-Host "User not found on Skype Server, Trying again" -foregroundcolor green
                $count++
            }

            if (($count -eq 3) -and ($correct -ne $true)){ 
                Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
                $correct = $true                
            }
            
        }Until($correct -eq $true)

        

    #----------------------------------------------------------------------------------Skype Removal END----------------------------------------------------------------------------------------->

    #---------------------------------------------------------------------------------Office365 License START--------------------------------------------------------------------------------------->

            Write-Host "---Removing Office 365 License(s)---" -foregroundcolor green -backgroundcolor gray

            $attemps = $False
            $correct = $False
            $count = 0

            do{       

                try{
                    if(Get-MsolUser -UserPrincipalName "$Username@Wisetechglobal.com"){
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
                        $attemps = $true
                        $correct = $true
                    }
                }     
            } Until($correct -eq $true)

        

    #----------------------------------------------------------------------------------Office365 License END------------------------------------------------------------------------------------------->

    <#------------------------------------------------------------------------------Move To Disabled OU START------------------------------------------------------------------------------------->
    
        $attemps = $false
        $correct = $False
        $count = 0
    
        Write-Host "---Moving User To Disabled OU---" -foregroundcolor green -backgroundcolor gray

        do{      

            try{
            
                move-adobject $UserWTG.DistinguishedName -TargetPath "OU=Disabled,OU=Accounts,OU=root,DC=wtg,DC=zone" -Credential $OnPremiseCred
                $correct = $true
                Write-Host "User moved to Disabled OU" -foregroundcolor green
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

    #------------------------------------------------------------------------------Move To Disabled OU END--------------------------------------------------------------------------------------->#>
    
            $complete = $true
        }
        Catch{
            break loop
        }

    }Until($complete -eq $true)

    [Validateset('Y','N', IgnoreCase)]$quit = read-host "Would you like to run again? (y/n)"

}Until($quit -eq 'n')

exit
      
#------------------------------------------------------------------------------------------Contributors----------------------------------------------------------------------------------------->
# James Rule
# Lachlan O'Dea#>
