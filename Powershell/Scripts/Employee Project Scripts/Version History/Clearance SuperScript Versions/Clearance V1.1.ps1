<#----------------------------------------------------------------------------------------NOTES------------------------------------------------------------------------->
Requirements before running
o365- Follow steps One, Two and Three before running: https://technet.microsoft.com/en-us/library/dn975125.aspx
Skype- Ensure you have installed the Skype For Business Management Software: SoftwareStore\Microsoft.com\Skype for Business 2015 Server

LOD 30-Mar-17 09:52: Started Script
LOD 30-Mar-17 13:49: Email Redirection functionality added
 
<-----------------------------------------------------------------------------------------------------------------------------------------------------------------------#>
<#------------------------------------------------------------------------------------Known bugs------------------------------------------------------------------------->


 
<-----------------------------------------------------------------------------------------------------------------------------------------------------------------------#>

#Clears Console
clear

Write-Host "------------------CLEARANCE SCRIPT V1.0------------------" -foregroundcolor Blue -backgroundcolor yellow

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
            do{
                $attemps = $false
                
                $Firstname = Read-Host 'What is the users first name? (Ensure use of Capitals)'
                $Lastname = Read-Host 'What is the users last name? (Ensure use of Capitals)'

                $Username = $Firstname + "." + $Lastname
                $UsernoDot = $Firstname + " " + $Lastname

                

                    Write-Host "Checking User Information" -foregroundcolor green
                    $UserWTG = Get-ADUser $Username 
                    if($UserWTG -ne $null){
                        Write-Host "Found user $Username in wtg.zone" -foregroundcolor green
                    }
                    else{
                        Write-Host "User not found in wtg.zone" -foregroundcolor green -backgroundcolor red
                        $UserWTG = $null
                    }

                   <# Get-PSSession | Remove-PSSession

                    $UserCorporate = Get-ADUser $Username  -Server corporate.cargowise.com
                    If($UserCorporate -ne $null){
                        Write-Host "Found user $Username in corporate.cargowise.com" -foregroundcolor green
                    }
                    else{
                        Write-Host "User not found in corporate.cargowise.com" -foregroundcolor green -backgroundcolor red
                        $UserCorporate = $null
                    }#>

                    if(($UserWTG -ne $null) -or ($UserCorporate -ne $null)){
                        $correct = $true
                    }
                    else{
                        $count++
                    }

                
                    if ($count -eq 3){ 
                        Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
                        #break loop
                    }
                 
            } Until($correct -eq $true) 

    #----------------------------------------------------------------------------------Collect User Data END------------------------------------------------------------------------------------>

   

    #----------------------------------------------------------------------------------User Options END----------------------------------------------------------------------------------------->
            
            Write-Host "---Redirect User Details---" -foregroundcolor green -backgroundcolor gray
            
            Get-PSSession | Remove-PSSession
            Remove-Variable -name proxysettings -EA SilentlyContinue

            $Redirect = 'n'
            $correct = $False

            [Validateset('Y','N', IgnoreCase)]$Redirect = (Read-Host "Do emails need to be redirected? (Y/N)")
            if($Redirect -eq 'Y'){ 
                $RedirectEmail = Read-Host "Please enter User with a dot for mail to be redirected to"
                
                #Connect to OnPremise
                $OnpremiseSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sydco-smai-1.wtg.zone/PowerShell/ -Authentication Kerberos -Credential $OnPremiseCred
                Import-PSSession $OnpremiseSession
                

                #Check for Valid Redirect Email Adress
                do{
                    

                    try{
                        if(Get-RemoteMailbox "$RedirectEmail"){
                            Write-Host "Mailbox $RedirectEmail found" -foregroundcolor green
                            Set-RemoteMailbox $Username -PrimarySmtpAddress "$Username@wtg.zone" -EmailAddressPolicyEnabled $False
                            Set-RemoteMailbox $Username -EmailAddresses @{remove="$Username@wisetechglobal.com"} -EmailAddressPolicyEnabled $False
                            Write-Host "Primary SMPT for $Username Removed" -foregroundcolor green
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
                                
    #----------------------------------------------------------------------------------User Options END------------------------------------------------------------------------------------------>

    #----------------------------------------------------------------------------------Skype Removal START--------------------------------------------------------------------------------------->

        Write-Host "---Removing User from Skype Server---" -foregroundcolor green -backgroundcolor gray

        $correct = $False
        $test = $null
        do{
            try{
                $test = Get-CsUser $Username
                if($test -ne $null){
                    Disable-CsUser $Username
                    Write-Host "User Removed From Skype Server" -foregroundcolor green
                    $correct = $true
                }
                else{
                    catch
                }
            }
            Catch{
                Write-Host "User could not be removed from Skype, Retrying" -foregroundcolor green -backgroundcolor red
                $count++
                if ($count -eq 3){ 
                    Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
                    $attemps = $true
                    break loop
                }
            }
            
        }Until($correct -eq $true)

        

    #----------------------------------------------------------------------------------Skype Removal END----------------------------------------------------------------------------------------->

            $complete = $true
        }
        Catch{
            break loop
        }

    }Until($complete -eq $true)

}Until($complete -eq $true)

exit
      
#------------------------------------------------------------------------------------------Contributors----------------------------------------------------------------------------------------->
# James Rule
# Lachlan O'Dea#>
