<#--------------------------------NOTES---------------------------------
Requirements before running
o365- Follow steps one, two and 3 before running: https://technet.microsoft.com/en-us/library/dn975125.aspx

LOD 16-Mar-17 11:28: Added Multi user functionality, self poputlating credential fields, 
selection for new user region locations

----------------------------------------------------------------------#>


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
$quit = $false

#select Running User
$input = read-host "Please Select a User `n 1 - James `n 2 - Lachlan `n 3 - Keane `n 4 - Other `n"

switch ($input){

    1 {$FullyUser = "James.Rule@wisetechglobal.com"
       $3LetterUser = "CORP\jmr"
       $AlertEmail = $FullyUser}

    2 {$FullyUser = "Lachlan.Odea@wisetechglobal.com" 
       $3LetterUser = "CORP\lod"
       $AlertEmail = $FullyUser}

    3 {$FullyUser = "Keane.Zhang@wisetechglobal.com"
       $3LetterUser = "CORP\kez"}

    default {$FullyUser = $null 
             $3LetterUser = $null
             $AlertEmail = $FullyUser}
}


#Gathering Credentials
do{
    
    #Gathering OnPremise Creds and Logging in

    Write-Host "---Exchange Onpremise Credentials---" -foregroundcolor green -backgroundcolor gray

    $correctOnpremise = $False
    $correctOnline =$False
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


    #Gathering Online Creds and Logging in

    Write-Host "---Exchange Online Credentials---" -foregroundcolor green -backgroundcolor gray

    $count = 0

    do{
 
        $OnlineCred = Get-Credential -Message 'Please provide MsolService credentials' -UserName $FullyUser

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


    if($correctOnpremise -eq $true -and $correctOnline -eq $true){
        $login = $true
    }

}until($login -eq $true)   
  
        
#Start Multiple User Loop ------------------------------------------------------------------------------------------------------------------------------------------------
do{
    
    #Finding user in AD
    Write-Host "---Retrieve User Details---" -foregroundcolor green -backgroundcolor gray
    
    do{
        $attemps = $false
        $Firstname = Read-Host 'What is the users first name?'

        $Lastname = Read-Host 'What is the users last name?'

        $Username = $Firstname + "." + $Lastname
        $UsernoDot = $Firstname + " " + $Lastname

        try{

            Write-Host "Check User Information" -foregroundcolor green
            $User = Get-ADUser $Username 
            Write-Host "Found user $User" -foregroundcolor green
            $correct = $true

        }
        Catch {

            Write-Host "User not found from AD" -foregroundcolor green -backgroundcolor red
            $count++
            if ($count -eq 3){ 
                Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
                $attemps = $true
                break
            }

        } 

    } Until($correct -eq $true) 

    if($attemps -eq $true){
        break
    }

    #select New Users Working region
    do{
        $location = $null
        $skypeServer = $null
        $OUPath = $null
        $count = 0
        $input = read-host "Please Select a region that the new user will be assigned to `n 1 - ASPAC `n 2 - EMEA `n 3 - AMERICAS`n"
        switch ($input){

            1 {
           
                    $input = read-host "Please Select User Country within ASPAC `n 1 - Australia `n 2 - China `n 3 - Singapore `n 4 - New Zealand`n"
                    switch ($input){
                    
                    1 {$location = "AU" 
                       $skypeServer = "sydcosfb.wtg.zone"
                       $OUPath= "OU=Sydney,OU=ASPAC,OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone"
                       }

                    2 {$location = "CN"}

                    3 {$location = "SG"}

                    4 {$location = "NZ"}
                    }
              }

            1 {$location = "AU"}

            2 {$location = "CN"}

            3 {$location = "US"}

            4 {$location = "ZA"}

        }
        
        if($location -eq $null){
            write-host "Please Select a location to continue"
            $count++
        }

        if($count -eq 3){
            $attemps = $true
            break
        }

    }until($location -ne $null -and $attemps -ne $true)

    if($attemps -eq $true){
        break
    }
        
    #User Options

    $numReq = 'n'
    $isDev = 'n'

    [Validateset('Y','N', IgnoreCase)]$numReq = (Read-Host "Does the user need to dial out? (Y/N)")
    #waiting for boolean convert
    [Validateset('Y','N', IgnoreCase)]$isDev = (Read-Host "Is the User a Developer? (Y/N)")

 #----------------------------------------------------------------------------------AD Start------------------------------------------------------------------------------------------------------

    $attemps = $false
    $correct = $False
    $count = 0
        
    do{

        start-sleep -s 1
        Write-Host "---Moving User To Correct Region OU---" -foregroundcolor green -backgroundcolor gray

        try{
            
            move-adobject $User.DistinguishedName -TargetPath $OUPath -Credential $OnPremiseCred -Verbose   
       
            $correct = $true
        }
        Catch{

            Write-Host "AD OU Move Failed" -foregroundcolor green -backgroundcolor red
            $count++
            if ($count -eq 3){
                Write-Host "AD OU Move Failed 3 times, cancelling user creation" -foregroundcolor green -backgroundcolor red 
                $attemps = $true
                break
            }
        } 

    } Until($correct -eq $true) 

    #----------------------------------------------------------------------------------AD End--------------------------------------------------------------------------------------------------------

             
    #Creating Maibox Locally and enabling Online Access
    $SendFrom = $OnlineCred.UserName -replace ("0x5c", "")

    Write-Host "---Enabling Remote Mailbox---" -foregroundcolor green -backgroundcolor gray

    $attemps = $false
    $correct = $False
    $count = 0

    do{

        start-sleep -s 1

        try{

            Enable-RemoteMailbox $UsernoDot -RemoteRoutingAddress "$Username@WiseTechGlobal.mail.onmicrosoft.com" -Verbose:$false
            $correct = $true

        }
        Catch{

            Write-Host "Remote Mailbox creation failed, waiting and trying again" -foregroundcolor green -backgroundcolor red 
            $count++
            if ($count -eq 3){
                Write-Host "Remote Mailbox creation failed 3 times, cancelling user creation" -foregroundcolor green -backgroundcolor red 
                $attemps = $true
                break
            }
        } 

    } Until($correct -eq $true) 

    if($attemps -eq $true){
        break
    }

    #Assigns o365 License to user

    Write-Host "---Assigning Office 365 License(s)---" -foregroundcolor green -backgroundcolor gray

    $attemps = $false
    $correct = $False
    $count = 0

    do{

        start-sleep -s 2

        try{

            Set-Msoluser -UserPrincipalName "$Username@Wisetechglobal.com" -Usagelocation $location -verbose:$False
            Set-MsoluserLicense -UserPrincipalName "$Username@Wisetechglobal.com" -Addlicenses  Wisetechglobal:ENTERPRISEPACK
            $correct = $true

        }
        Catch{

            Write-Host "License Assignment failed, waiting and trying again" -foregroundcolor green -backgroundcolor red
            $count++
            if ($count -eq 3){ 
                Write-Host "License Assignment failed 3 times, cancelling user creation" -foregroundcolor green -backgroundcolor red 
                $attemps = $true
            }
        }
     
    } Until($correct -eq $true)

    if($attemps -eq $true){
        break
    }

    #----------------------------------------------------------------------------------SKYPE Start------------------------------------------------------------------------------------------------------
   
    $attemps = $false
    $correct = $False
    $count = 0

    do{

        start-sleep -s 1
        Write-Host "---Enabling Skype For Business Account---" -foregroundcolor green -backgroundcolor gray

        try{
            
            enable-csuser -identity "$Username@wisetechglobal.com" -RegistrarPool $skypeServer -sipaddress "sip:$Username@wisetechglobal.com" -Credential $OnPremiseCred 
            set-csuser -Identity "$Username@wisetechglobal.com" -Enterprisevoiceenabled $true -Credential $OnPremiseCred 
            $correct = $true
        }
        Catch{

            Write-Host "Skype creation failed, waiting and trying again" -foregroundcolor green -backgroundcolor red
            $count++
            if ($count -eq 3){
                Write-Host "Remote Skype creation failed 3 times, cancelling user creation" -foregroundcolor green -backgroundcolor red 
                $attemps = $true
                break
            }
        } 

    } Until($correct -eq $true) 

    #----------------------------------------------------------------------------------SKYPE End--------------------------------------------------------------------------------------------------------

    #Ask for Re-run   
    $quit = read-host "Would you like to run again? (y/n)"


}until($quit -notlike 'y*' -or $attemps -eq $true) 


#Script Completion
Exit