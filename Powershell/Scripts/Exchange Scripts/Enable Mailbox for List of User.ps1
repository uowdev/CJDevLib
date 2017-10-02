# Description: Will change a specified list of Users o365 License. Useful when users have had a custom license applied and need a full license.
# Requirements: Onpremise credentials
# Created by: James Rule (Twitter: @Jgrrule Github: ClickyJimmy).
# Note: I haven't made this as generic as some others. You may need to spend some time working on this

Write-Host "---Creating Session---" -foregroundcolor green -backgroundcolor gray

Get-PSSession | Remove-PSSession

$OnPremiseCred = Get-Credential -Message 'OnPremise Credentials'
$OnpremiseSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri #Enteryouronpremismailboxaddresshere/PowerShell/ -Authentication Kerberos -Credential $OnPremiseCred
Import-PSSession $OnpremiseSession

Write-Host "---Enabling Remote Mailbox---" -foregroundcolor green -backgroundcolor gray
$attemps = $false
$correct = $False
$count = 0

$Users = Get-ADUser -SearchBase "OU=SaoPaulo,OU=AMERICAS,OU=Users,OU=Accounts,OU=root,DC=,DC=" -Filter *
do {
    foreach ($Username in $Users) {
        $Username = $Username.Samaccountname
        try {
            if (Get-RemoteMailbox $Username -ErrorAction SilentlyContinue) {
                Write-Host "Mailbox is already enabled" -foregroundcolor yellow -backgroundcolor blue
                $correct = $true
            }
            Else {
                Write-Host "Enabling Mailbox for $Username" -foregroundcolor yellow -backgroundcolor blue
                Enable-RemoteMailbox $Username -RemoteRoutingAddress "$Username@WBIZNIZ.mail.onmicrosoft.com"
                $correct = $true
            }
        }

        Catch {
            Write-Host "Remote Mailbox creation failed, waiting and trying again" -foregroundcolor green -backgroundcolor red
            Start-Sleep -s 10
            $count++
            if ($count -eq 3) {
                Write-Host "Remote Mailbox creation failed 3 times, cancelling process" -foregroundcolor green -backgroundcolor red
                $attemps = $true
                break loop
            }
        }
    }
} Until($correct -eq $true)
