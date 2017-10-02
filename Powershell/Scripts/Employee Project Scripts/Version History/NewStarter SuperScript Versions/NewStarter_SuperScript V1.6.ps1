<#----------------------------------------------------------------------------------------NOTES------------------------------------------------------------------------->
Requirements before running
o365- Follow steps One, Two and Three before running: https://technet.microsoft.com/en-us/library/dn975125.aspx
Skype- Ensure you have installed the Skype For Business Management Software: SoftwareStore\Microsoft.com\Skype for Business 2015 Server

LOD 16-Mar-17: Added Multi user functionality, self poputlating credential fields, 
selection for new user region locations

JMR 16-Mar-17: Added function for Ivka to be granted access to mailbox IF Developer, Added module headers/footers

JMR 22-Mar-17: Added sleeps in catches. This way if it fails due to attempting too quickly- it will wait and try again (rather than pointlessly try 3 times)
Also commented out references to Skype phone number. Removed all references to $Alertemail

JMR 27-Mar-17: Changed method for new mailboxes (Now Enables Remote Mailbox, no more migration) 

LOD 30-Mar-17 14:58: Kind of FIXED Skype and License Assignment

LOD 30-Mar-17 15:26: License Module moved above Remote mailbox creation after investigation Online Mailbox Cannot be created with Office Online o365 License

JMR 05-May-17 15:06: Fixed Skype and License Assignment, removed Lachlan from Users list. (RIP) 
 
<-----------------------------------------------------------------------------------------------------------------------------------------------------------------------#>
<#------------------------------------------------------------------------------------Known bugs------------------------------------------------------------------------->


 
<-----------------------------------------------------------------------------------------------------------------------------------------------------------------------#>


#Clears Console
clear
Write-Host "------------------NewStarter SCRIPT V1.6------------------" -foregroundcolor Blue -backgroundcolor yellow
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
$input = read-host "Please Select an operating User `n 1 - James `n 2 - Keane `n 3 - Ivka `n"

switch ($input){

    1 {$FullyUser = "James.Rule@wisetechglobal.com"
       $3LetterUser = "CORP\jmr"
       $CORPORATE = "CORPORATE\james.rule"
       $name = "James.Rule"}

    2 {$FullyUser = "Keane.Zhang@wisetechglobal.com"
       $3LetterUser = "CORP\kez"
       $CORPORATE = "CORPORATE\keane.zhang"
       $name = "Keane.Zhang"}

    3 {$FullyUser = "Ivka.Novokmet@wisetechglobal.com"
       $3LetterUser = "CORP\ivn"
       $CORPORATE = "CORPORATE\Ivka.Novokmet"
       $name = "Ivka.Novokmet"}
    
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

                      $Firstname = Read-Host 'What is the users first name? (Ensure use of Capitals)'


                            $Lastname = Read-Host 'What is the users last name? (Ensure use of Capitals)'


                $Username = $Firstname + "." + $Lastname
                $UsernoDot = $Firstname + " " + $Lastname

                try{

                    Write-Host "Check User Information" -foregroundcolor green
                    $User = Get-ADUser $Username 
                    Write-Host "Found user $User" -foregroundcolor green
                    Rename-ADObject $User -NewName $UsernoDot -Credential $OnPremiseCred
                    Set-ADuser  $User -GivenName $Firstname -Surname $Lastname -Credential $OnPremiseCred

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
        $attempts = $false
        
        do{

            Write-Host "Capture User OU Location" -foregroundcolor green -backgroundcolor gray

            try{

                #Region
                $selection=0
                $input=$null
                $region=$null
                $office=$null        
                $team=$null        
                $OUPath=$null
                $teams=$null
                $offices=$null
                $regions=$null
                $locationset=$false
        
                [array]$regions=Get-adorganizationalunit -SearchBase 'OU=Users,OU=Accounts,OU=root,DC=wtg,DC=zone' -SearchScope Onelevel -filter * | Select-Object DistinguishedName,Name
        
                foreach($region in $regions){

                    write-host $selection. $region.Name `n
                    $selection++
                }

                $input=read-host 'Region Selection ' 

                $region=$regions[$input]
                
                Write-host `n 'Region=' $region `n | Select-Object Name

                #Office
                $selection=0
                $input=$null                
                 
                [array]$offices=get-adorganizationalunit -SearchBase $region.distinguishedname -SearchScope Onelevel -filter * 
        
                if($offices.length -eq 1){
                    $office=$offices
                }

                else{
                    foreach($office in $offices){

                        write-host $selection. $office.Name `n
                        $selection++
                    }

                    $input =(read-host "Office Selection")
                    $office=$offices[$input] 
                    Write-host `n 'Office =' $office `n | Select-Object Name 
                }
              
        

                #Team
                $selection=0
                $input=$null
               
     
        
                [array]$teams= get-adorganizationalunit -SearchBase $office.DistinguishedName -SearchScope Onelevel -filter * 
        
                if($teams.length -eq 0){
                    $OUPath=$office
                    $locationset=$true
                }

                else{

                    foreach($team in $teams){

                        write-host $selection. $team.Name `n
                        $selection++
                    }
                    
                    $input=read-host 'Team Selection, if no team required; enter "n"'
                    
                    if($input -eq 'n'){
                        $OUPath=$office
                        $locationset=$true
                        Write-host `n 'Location =' $OUPath | Select-Object Name
                    }
                    elseif($input -notin 0..($teams.Length-1) -and ($input -ne "n")){
                        catch
                    }
                    else{
                        $OUPath=$teams[$input]
                        $locationset=$true
                        Write-host `n 'Location =' $OUPath | Select-Object Name
                    }
                }     
            }
            catch{
                $count++
                if($count -eq 3){
                    $attempts = $true
                    break :loop
                }
            }
        }until(($locationset -eq $true) -or ($attempts -eq $true))
        
    #----------------------------------------------------------------------------------Regional Selection END----------------------------------------------------------------------------->

    #----------------------------------------------------------------------------------User Options END----------------------------------------------------------------------------------->

           # $numReq = 'n' ---Functionality still needs to be added
            $isDev = 'n'

           # [Validateset('Y','N', IgnoreCase)]$numReq = (Read-Host "Does the user need to dial out? (Y/N)") ---Functionality still needs to be added
            [Validateset('Y','N', IgnoreCase)]$isDev = (Read-Host "Is the User a Developer? (Y/N)")

    #----------------------------------------------------------------------------------User Options END------------------------------------------------------------------------------------>

    #----------------------------------------------------------------------------------AD Start-------------------------------------------------------------------------------------------->
           
           #ADGroup

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

            } Until(($correct -eq $true) -or ($attempts = $true)) 


            <#ADName

          #  Write-Host "---Renaming AD User---" -foregroundcolor green -backgroundcolor gray

            do{      

                try{
            
                    
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

            } Until(($correct -eq $true) -or ($attempts = $true)) #>

    #----------------------------------------------------------------------------------AD End-------------------------------------------------------------------------------------------------->

    <#----------------------------------------------------------------------------------Mailbox Creation START---------------------------------------------------------------------------------->
 
            $SendFrom = $OnlineCred.UserName -replace ("0x5c", "")

            Write-Host "---Enabling Remote Mailbox---" -foregroundcolor green -backgroundcolor gray

            $attemps = $false
            $correct = $False
            $count = 0

            do{      

                try{
                    if(Get-RemoteMailbox $Username){
                        Write-Host "Mailbox is already enabled" -foregroundcolor yellow -backgroundcolor blue
                        $correct = $true
                    }
                    Else{
                        Enable-RemoteMailbox $Username -RemoteRoutingAddress "$Username@WiseTechGlobal.mail.onmicrosoft.com" -Verbose:$false
                        $correct = $true
                    }
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

    #----------------------------------------------------------------------------------Mailbox Creation END-----------------------------------------------------------------------------------------#>
    
    <#---------------------------------------------------------------------------------Office365 License START--------------------------------------------------------------------------------------->

            Write-Host "---Assigning Office 365 License(s)---" -foregroundcolor green -backgroundcolor gray

            $attemps = $False
            $correct = $False
            $count = 0

            $input = Read-Host "Country the user will be working in"

            do{       

                try{
                        Get-MsolUser -UserPrincipalName "$Username@Wisetechglobal.com"                    
                        Set-Msoluser -UserPrincipalName "$Username@Wisetechglobal.com" -Usagelocation $input -verbose:$False
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
                             

    #----------------------------------------------------------------------------------Office365 License END-------------------------------------------------------------------------------------------#>

    #----------------------------------------------------------------------------------Skype For Business START---------------------------------------------------------------------------------------->

            Write-Host "---Enabling Skype For Business Account---" -foregroundcolor green -backgroundcolor gray

            $attemps = $false
            $correct = $False
            $count = 0
            $selection = 0
            [array]$pools = Get-CSPool | Where-Object {$_.Identity -like "*sfb.wtg.zone"} 
            
            foreach($pool in $pools){
            write-host $selection. $pool.Identity `n
            $selection++
            }

            $input =(read-host "Skype Pool Selection")
            $pool=$pools[$input] 
            Write-host `n 'Skype Pool =' $pool.Identity

            do{       
                Start-Sleep -s 5
                try{
                    
                     enable-csuser -identity "$Username@wisetechglobal.com" -RegistrarPool $pool.Identity -sipaddress "sip:$Username@wisetechglobal.com" -verbose:$false
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
# Lachlan O'Dea