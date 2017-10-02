<#----------------------------------------------------------------------------------------NOTES------------------------------------------------------------------------->
Requirements before running
o365- Follow steps One, Two and Three before running: https://technet.microsoft.com/en-us/library/dn975125.aspx
Skype- Ensure you have installed the Skype For Business Management Software: SoftwareStore\Microsoft.com\Skype for Business 2015 Server

LOD 16-Mar-17: Added Multi user functionality, self poputlating credential fields, 
selection for new user region locations

JMR 16-Mar-17: Added function for Ivka to be granted access to mailbox IF Developer, Added module headers/footers

JMR 22-Mar-17: Added sleeps in catches. This way if it fails due to attempting too quickly- it will wait and try again (rather than pointlessly try 3 times)
Also commented out references to Skype phone number. Removed all references to $Alertemail
 
<-----------------------------------------------------------------------------------------------------------------------------------------------------------------------#>
<#------------------------------------------------------------------------------------Known bugs------------------------------------------------------------------------->

-  IF a Skype account already exists for the user- no message or notification informs the user so. Instead- it returns an error and loops in attempting to create the user
 
<-----------------------------------------------------------------------------------------------------------------------------------------------------------------------#>


#Clears Console
clear

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

#----------------------------------------------------------------------------------Select User START--------------------------------------------------------------------->
$input = read-host "Please Select an operating User `n 1 - James `n 2 - Lachlan `n 3 - Keane `n"

switch ($input){

    1 {$FullyUser = "James.Rule@wisetechglobal.com"
       $3LetterUser = "CORP\jmr"}

    2 {$FullyUser = "Lachlan.Odea@wisetechglobal.com" 
       $3LetterUser = "CORP\lod"}

    3 {$FullyUser = "Keane.Zhang@wisetechglobal.com"
       $3LetterUser = "CORP\kez"}

    default {$FullyUser = $null 
             $3LetterUser = $null}
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

            $OnpremiseSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sydco-smai-1.wtg.zone/PowerShell/ -Authentication Kerberos -Credential $OnPremiseCred -Verbose
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

#----------------------------------------------------------------------------------Gather Credentials START------------------------------------------------------------------->
        
#----------------------------------------------------------------------------------Mutliple User Creation START--------------------------------------------------------------->

do{
    :loop do{

        $complete = $false

        try{

    #----------------------------------------------------------------------------------Collect User Data START---------------------------------------------------------------->
            Write-Host "---Retrieve User Details---" -foregroundcolor green -backgroundcolor gray
    
            do{
                $attemps = $false

                do{
                    try{

                            $Firstname = Read-Host 'What is the users first name? (Ensure use of Capitals)'

                            
                            }
                            catch{
                            }
                            until ()

                            $Lastname = Read-Host 'What is the users last name? (Ensure use of Capitals)'


                $Username = $Firstname + "." + $Lastname
                $UsernoDot = $Firstname + " " + $Lastname

                try{

                    Write-Host "Check User Information" -foregroundcolor green
                    $User = Get-ADUser $Username 
                    Write-Host "Found user $User" -foregroundcolor green
                    Rename-ADObject $User -NewName $UsernoDot -verbose
                    Set-ADuser  $User -GivenName $Firstname -Surname $Lastname -verbose

                    $correct = $true
                }

                Catch {

                    Write-Host "User not found in AD" -foregroundcolor green -backgroundcolor red
                    $count++
                    if ($count -eq 3){ 
                        Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
                        $attemps = $true
                        break loop
                    }
                } 
            } Until($correct -eq $true) 

    #----------------------------------------------------------------------------------Collect User Data END--------------------------------------------------------------------->

    #----------------------------------------------------------------------------------Regional Selection START------------------------------------------------------------------>
        
        #Select New Users Working region

        $count = 0
        do{
            $location = $null
            $skypeServer = $null
            $OUPath = $null
        
            $input = read-host "Please select a Geographical Region that the new user will be assigned to `n 1 - ASPAC `n 2 - EMEA `n 3 - AMERICAS`n"

            switch ($input){

            1 {
                    $input = read-host "Please Select User Office Region within ASPAC `n 1 - Auckland `n 2 - Brisbane `n 3 - Melbourne `n 4 - Nanjing `n 5 - Shanghai `n 6 - Shenzhen `n 7 - Singapore `n 8 - Sydney `n"
                    switch ($input){
                    
                    1 {$location = "NZ" 
                       $skypeServer = "sydcosfb.wtg.zone"
                       $OUPath= "OU=Auckland,OU=ASPAC,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }

                    2 {$location = "AU"
                       $skypeServer = "sydcosfb.wtg.zone"
                       $OUPath= "OU=Brisbane,OU=ASPAC,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }

                    3 {$location = "AU"
                       $skypeServer = "sydcosfb.wtg.zone"
                       $OUPath= "OU=Melbourne,OU=ASPAC,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }

                    4 {$location = "CN"
                       $skypeServer = "sydcosfb.wtg.zone"
                       $OUPath= "OU=Nanjing,OU=ASPAC,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }
                    5 {$location = "CN"
                       $skypeServer = "szxcosfb.wtg.zone"
                       $OUPath= "OU=Shanghai,OU=ASPAC,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }
                    6 {$location = "CN"
                       $skypeServer = "szxcosfb.wtg.zone"
                       $OUPath= "OU=Shenzhen,OU=ASPAC,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }
                    7 {$location = "CN"
                       $skypeServer = "SINCO-SSBA-1.wtg.zone"
                       $OUPath= "OU=Singapore,OU=ASPAC,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }
                    8 {$location = "AU"
                       $skypeServer = "sydcosfb.wtg.zone"
                       $OUPath= "OU=Sydney,OU=ASPAC,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }
                    }
             }
           
            2 {
                    $input = read-host "Please Select User Office Region within EMEA `n 1 - Cape Town `n 2 - Durban `n 3 - Johannesburg `n 4 - Milano `n 5 - Milton Keynes `n 6 - Russia `n 7 - Wiesbaden `n"
                    switch ($input){
                    
                    1 {$location = "ZA" 
                       $skypeServer = "jnbcosfb.wtg.zone"
                       $OUPath= "OU=Cape Town,OU=EMEA,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }

                    2 {$location = "ZA"
                       $skypeServer = "jnbcosfb.wtg.zone"
                       $OUPath= "OU=Durban,OU=EMEA,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }

                    3 {$location = "ZA"
                       $skypeServer = "jnbcosfb.wtg.zone"
                       $OUPath= "OU=Johannesburg,OU=EMEA,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }

                    4 {$location = "IT"
                       $skypeServer = "loncosfb.wtg.zone"
                       $OUPath= "OU=Milano,OU=EMEA,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }
                    5 {$location = "GB"
                       $skypeServer = "loncosfb.wtg.zone"
                       $OUPath= "OU=Milton Keynes,OU=EMEA,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }
                    6 {$location = "CN"
                       $skypeServer = ""
                       $OUPath= "OU=Russia,OU=EMEA,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }
                    7 {$location = "CN"
                       $skypeServer = "loncosfb.wtg.zone"
                       $OUPath= "OU=Wiesbaden,OU=EMEA,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }
                    }
              }

            3 {$location = "US" 
               $skypeServer = "ordcosfb.wtg.zone"
               $OUPath= "OU=Chicago,OU=AMERICAS,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                
            }

        }
        
            if($location -eq $null){
                write-host "Please Select a location to continue"
                $count++
            }

            if($count -eq 3){
                $attemps = $true
                break loop
            }

        }until($location -ne $null -and $attemps -ne $true)
        
    #----------------------------------------------------------------------------------Regional Selection END----------------------------------------------------------------------------->

    #----------------------------------------------------------------------------------User Options END----------------------------------------------------------------------------------->

           # $numReq = 'n' ---Functionality still needs to be added
            $isDev = 'n'

           # [Validateset('Y','N', IgnoreCase)]$numReq = (Read-Host "Does the user need to dial out? (Y/N)") ---Functionality still needs to be added
            [Validateset('Y','N', IgnoreCase)]$isDev = (Read-Host "Is the User a Developer? (Y/N)")

    #----------------------------------------------------------------------------------User Options END------------------------------------------------------------------------------------>

    #----------------------------------------------------------------------------------AD Start-------------------------------------------------------------------------------------------->
            $attemps = $false
            $correct = $False
            $count = 0
    
            Write-Host "---Moving User To Correct Region OU---" -foregroundcolor green -backgroundcolor gray

            do{      

                try{
            
                    move-adobject $User.DistinguishedName -TargetPath $OUPath -Credential $OnPremiseCred
                    $correct = $true
                }

                Catch{
                    Start-Sleep -s 10
                    Write-Host "AD OU Move Failed, waiting and trying again" -foregroundcolor green -backgroundcolor red
                    $count++
                    if ($count -eq 3){
                        Write-Host "AD OU Move Failed 3 times, cancelling user creation" -foregroundcolor green -backgroundcolor red 
                        $attemps = $true
                        break loop
                    }
                } 

            } Until($correct -eq $true) 

    #----------------------------------------------------------------------------------AD End-------------------------------------------------------------------------------------------------->

    #----------------------------------------------------------------------------------Mailbox Creation START---------------------------------------------------------------------------------->
 
            $SendFrom = $OnlineCred.UserName -replace ("0x5c", "")

            Write-Host "---Enabling Remote Mailbox---" -foregroundcolor green -backgroundcolor gray

            $attemps = $false
            $correct = $False
            $count = 0

            do{      

                try{

                    Enable-RemoteMailbox $Username -RemoteRoutingAddress "$Username@WiseTechGlobal.mail.onmicrosoft.com" -Verbose:$false
                    $correct = $true
                }

                Catch{

                    Write-Host "Remote Mailbox creation failed, waiting and trying again" -foregroundcolor green -backgroundcolor red 
                    Start-Sleep -s 10
                    $count++
                    if ($count -eq 3){
                        Write-Host "Remote Mailbox creation failed 3 times, cancelling user creation" -foregroundcolor green -backgroundcolor red 
                        $attemps = $true
                        break loop
                    }
                } 
            } Until($correct -eq $true) 

          <# if($isDev = 'Y'){

                $proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
                $OnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $OnlineCred -Authentication Basic -AllowRedirection -SessionOption $proxysettings
                Import-PSSession $OnlineSession -Prefix 'o365'


                Add-MailboxPermission -Identity "$Username@Wisetechglobal.com" -User "Ivka.Novokmet@wisetechglobal.com" -AccessRights FullAccess
                 Write-Host "---Mailbox Permission Added---" -foregroundcolor green -backgroundcolor gray

            }#>

    #----------------------------------------------------------------------------------Mailbox Creation END----------------------------------------------------------------------------------------->
   
    #---------------------------------------------------------------------------------Office365 License START--------------------------------------------------------------------------------------->

            Write-Host "---Assigning Office 365 License(s)---" -foregroundcolor green -backgroundcolor gray

            $attemps = $False
            $correct = $False
            $count = 0

            do{       

                try{

                    Set-Msoluser -UserPrincipalName "$Username@Wisetechglobal.com" -Usagelocation $location -verbose:$False
                    Set-MsoluserLicense -UserPrincipalName "$Username@Wisetechglobal.com" -Addlicenses  Wisetechglobal:ENTERPRISEPACK -verbose:$False
                    $correct = $true
                }

                Catch{

                    Write-Host "License Assignment failed, waiting and trying again" -foregroundcolor green -backgroundcolor red
                    Start-Sleep -s 10
                    $count++
                    if ($count -eq 3){ 
                        Write-Host "License Assignment failed 3 times, cancelling user creation" -foregroundcolor green -backgroundcolor red 
                        $attemps = $true
                        break loop
                    }
                }     
            } Until($correct -eq $true)

        

    #----------------------------------------------------------------------------------Office365 License END------------------------------------------------------------------------------------------->

    #----------------------------------------------------------------------------------Skype For Business START---------------------------------------------------------------------------------------->

            Write-Host "---Enabling Skype For Business Account---" -foregroundcolor green -backgroundcolor gray

            $attemps = $false
            $correct = $False
            $count = 0 
        
            do{       

                try{
            
                    enable-csuser -identity "$Username@wisetechglobal.com" -RegistrarPool $skypeServer -sipaddress "sip:$Username@wisetechglobal.com" -verbose:$false
                    set-csuser -Identity "$Username@wisetechglobal.com" -Enterprisevoiceenabled $true -verbose:$false
                    $correct = $true
                     Write-Host "---Skype Account Succesful---" -foregroundcolor green -backgroundcolor gray
                }

                Catch{

                    Write-Host "Skype creation failed, waiting and trying again" -foregroundcolor green -backgroundcolor red
                    Start-Sleep -s 10
                    $count++
                    if ($count -eq 3){
                        Write-Host "Remote Skype creation failed 3 times, cancelling user creation" -foregroundcolor green -backgroundcolor red 
                        $attemps = $true
                        break loop
                    }
                } 
            } Until($correct -eq $true) 

            $complete = $true

    #----------------------------------------------------------------------------------Skype For Business End------------------------------------------------------------------------------------------->
        }

        catch{
            break loop
        }

    }until($complete -eq $true)

    [Validateset('Y','N', IgnoreCase)]$quit = read-host "Would you like to run again? (y/n)"

}until($quit -eq 'n')

#----------------------------------------------------------------------------------Mutliple User Creation END-------------------------------------------------------------------------------------------->
Exit

#------------------------------------------------------------------------------------------Contributors----------------------------------------------------------------------------------------------->
# James Rule
# Lachlan O'dea