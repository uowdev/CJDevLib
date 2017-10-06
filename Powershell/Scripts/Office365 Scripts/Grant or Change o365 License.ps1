# Description: Will change a specified list of Users o365 License. Useful when users have had a custom license applied and need a full license.
# Requirements: Msol session (Included)
# Created by : James Rule (Twitter: @jgrrule Github: ClickyJimmy)

$Credentials = Get-Credentials
Connect-MsolService -Credential $Credentials

#You will need to change the AccountSkuID
$AccountSkuId = 'BIZNIZ:ENTERPRISEPACK'
$attemps = $false
$correct = $False
$count = 0
#I am retreiving the User list from a text document, but you could change this to look at an OU
$Users =  Get-Content "c:\temp\users.txt"

#This is an example of using License Options to only grant certain permissions. In this example, we are granint everything but SkypeforBusiness
$SkypeOnlineDisabled = New-MsolLicenseOptions -AccountSkuId "$AccountSkuId" -DisabledPlans "MCOSTANDARD"

do {
    foreach ($Username in $Users) {
        $Username = $Username.Samaccountname + "@BIZNIZtechglobal.com"
        try {
            Write-Host "Changing license for $Username" -foregroundcolor yellow -backgroundcolor blue
            #You need to assign the correct Usagelocation when assigning a license. Comment out if you don't want to change.
            Set-MsolUser -UserPrincipalName $Userupn -UsageLocation BR
            #This REMOVES the existing license. Comment out if you are assigning license to unlicensed user.
            Set-MsolUserLicense -UserPrincipalName $Username -RemoveLicenses "$AccountSkuId"
            #Adds the License. You may want to comment out -Licenseoptions if you just want to assign the full license.
            Set-MsolUserLicense -UserPrincipalName $Username -AddLicenses "$AccountSkuId" -LicenseOptions $SkypeOnlineDisabled
        }
        Catch { #I use this catch fairly often - sometimes when working with different types of sessions, I have found that some operations fail the first time, perhaps due
                # to some quirk of the PSsession, but work the second time. Either way you can change this if you wanted.
            Write-Host "Failed, waiting and trying again" -foregroundcolor green -backgroundcolor red
            Start-Sleep -s 10
            $count++
            if ($count -eq 3) {
                Write-Host "Failed 3 times, cancelling process" -foregroundcolor green -backgroundcolor red
                $attemps = $true
                break loop
            }
        }
    }
    $correct = $True
} Until($correct -eq $true)
