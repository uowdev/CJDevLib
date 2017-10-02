# Description: THis script takes a list of users, either from User input or from .txt, and runs clearance operations on those users. Currently features: 
#   1) AD CORP + Corporate - Set as disabled
#   2) Exchange OnPremise and O365 Online - Forward emails + hide from exchange list
#   3) Skype for Business
# Requirements: Online and Onpremise sessions (Included in script- with prefix), OnPremise and Online admin account credentials, Ran as 3 letter admin (Skype Credentials), Msolservice permissions. 
# Created by : JMR on 4/9/17
# Current Version Notes: 

#---------------------------------------------------------------------------------------------

#--Batch operation - Check Account Names

$UserList = Get-Content "C:\PS\UserList.txt"
Foreach ($User in $UserList) {
    try {
        $check = 1 
        Write-Host "Checking AD for $User " -foregroundcolor yellow
        $check = get-aduser $User    
        if ($check)
        {Write-Host "User Found : $check" -foregroundcolor green}            
    }
    catch {
        Write-Host $_.Exception |format-list -force 
    }    
}

#--Staff Clearance - Batch operation - Disable CORP Account--

$UserList = Get-Content "C:\PS\UserList.txt"
Foreach ($User in $UserList) {
    try { 
        Write-Host "Disabling $User in CORP" -foregroundcolor yellow
        Disable-ADAccount $User -Verbose
        Write-Host "$User Disabled" -foregroundcolor green
    }
    catch {
        Write-Host $_.Exception |format-list -force 
    }    
}

#--Staff Clearance - Batch operation - - AD Wipe Group Membership--

$UserList = Get-Content "C:\PS\UserList.txt"
Foreach ($User in $UserList) {
    try { 
        $Groups = Get-ADPrincipalGroupMembership $User
        foreach ($group in $Groups) {
            if ($group.name -notcontains "Domain Users") {
                Remove-ADGroupMember -Identity $group -Members $User -Confirm:$false 
            }
        }
        Write-Host "$User Removed from all Groups" -foregroundcolor green
    }
    catch {
        Write-Host $_.Exception |format-list -force 
    }    
}

#--Staff Clearance - Batch operation - Remove Skype for Business Account--

$UserList = Get-Content "C:\PS\UserList.txt"
Foreach ($User in $UserList) {
    try { 
        Write-Host "Removing $User from Skype for Business" -foregroundcolor yellow
        Disable-CsUser -identity "$User@wisetechglobal.com" -Verbose
        Write-Host "$User Removed from Skype Server" -foregroundcolor green
    }
    catch {
        Write-Host $_.Exception |format-list -force 
    }    
}

#--end--
#--Staff Clearance - Batch operation - Disable CORPORATE Account--

Write-Host "---CORPORATE Credentials Required---" -foregroundcolor green -backgroundcolor gray
$correctCorporate = $False
$count = 0

do {
    $CORPORATEcred = Get-Credential -Message 'Please provide CORPORATE\ credentials' 
    $DomainNetBIOS = $CORPORATEcred.username.Split("{\}")[0]
    $UserName = $CORPORATEcred.username.Split("{\}")[1]
    $Password = $CORPORATEcred.GetNetworkCredential().password
    $DomainFQDN = (Get-ADDomain $DomainNetBIOS).DNSRoot

    try {           
        $DomainObj = "LDAP://" + $DomainFQDN
        New-Object System.DirectoryServices.DirectoryEntry($DomainObj, $UserName, $Password)
        Write-Host "---CORPORATE Credentials Validated---" -foregroundcolor green -backgroundcolor gray
        $correctCorporate = $true
    }

    Catch {
        Write-Host "Account details incorrect, did you forget to include domain?" -foregroundcolor green -backgroundcolor red
        $count++
        if ($count -eq 3) {
            Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
            exit
        }
    }     
} Until($correctCorporate -eq $true)

$UserList = Get-Content "C:\PS\UserList.txt"
Foreach ($User in $UserList) {
    try { 
        Write-Host "Disabling $User In CORPORATE" -foregroundcolor yellow
        Disable-ADAccount $User -Server corporate.cargowise.com  -Credential $CORPORATEcred
        Write-Host "$User CORPORATE Account Disabled" -foregroundcolor green        
    }
    catch {
        Write-Host $_.Exception |format-list -force 
    }    
}

#--end--
#--Staff Clearance - Batch operation - Forward Emails and/or Hide From Exchange List --

Clear-Host
get-pssession | remove-pssession
Write-Host "---Exchange Onpremise Credentials---" -foregroundcolor green -backgroundcolor gray
$correctOnpremise = $False
$count = 0

do {
    $OnPremiseCred = Get-Credential -Message 'Please provide Exchange Onpremise credentials' 

    try {
        $OnpremiseSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://sydco-smai-1.wtg.zone/PowerShell/ -Authentication Kerberos -Credential $OnPremiseCred
        Import-PSSession $OnpremiseSession
        Write-Host "---Exchange On Premise Session started---" -foregroundcolor green -backgroundcolor gray
        $correctOnpremise = $true
    }

    Catch {
        Write-Host "Account details incorrect, did you forget to include domain?" -foregroundcolor green -backgroundcolor red
        $count++
        if ($count -eq 3) {
            Write-Host "Too many attempts" -foregroundcolor green -backgroundcolor red 
        } 
    }             
} Until($correctOnpremise -eq $true)

Write-Host "---Redirect User Details---" -foregroundcolor green -backgroundcolor gray

$UserList = Get-Content "C:\PS\UserList.txt"
Foreach ($User in $UserList) {
    do {
        try { 
            $Redirect = 'n'
            $correct = $False

            [Validateset('Y', 'N', IgnoreCase)]$Redirect = (Read-Host "Does $User email need to be redirected? (Y/N)")
            
            if ($Redirect -eq 'Y') { 
                $RedirectEmail = Read-Host "Please enter User with a dot for mail to be redirected to" 
            
                #Hide From Exchange List
                Set-RemoteMailbox $User -HiddenFromAddressListsEnabled $true

                #Check for Valid Redirect Email Adress
            
                if (Get-RemoteMailbox "$RedirectEmail") {
                    Write-Host "Mailbox $RedirectEmail found" -foregroundcolor green
                    Set-RemoteMailbox $User -PrimarySmtpAddress "$User@wtg.zone" -EmailAddressPolicyEnabled $False
                    Set-RemoteMailbox $User -EmailAddresses @{remove = "$User@wisetechglobal.com"} -EmailAddressPolicyEnabled $False
                    Write-Host "Primary SMPT for $User Removed, waiting 20 seconds" -foregroundcolor green
                    Start-Sleep -s 20
                    Set-RemoteMailbox $RedirectEmail -EmailAddresses @{add = "$User@wisetechglobal.com"}
                    Write-Host "$User email added to $RedirectEmail" -foregroundcolor green
                    $correct = $true
                }
                else {
                    Write-Host "Mailbox $RedirectEmail not found" -foregroundcolor red
                    $correct = $true
                }
            }
            else {
                Set-RemoteMailbox $User -HiddenFromAddressListsEnabled $true
                $correct = $true                
            }
        }
        catch {
            Write-Host $_.Exception |format-list -force 
        }            
    } Until($correct -eq $true)             
}
      
#Edit LOG 
# 4/9/17 (JMR): Created 
# 11/9/17 #2 (JMR): Updated aliases to full commands. 

     
