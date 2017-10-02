# Title: Check OU Users for Mailbox and License
# Description: Checks all users within a specific OU if they have a mailbox and are licensed.
# Requirements: Run as Local Mailbox admin. Requires Office365 Admin (Session Commands Included)
# Created by: James Rule (Twitter: @Jgrrule Github: ClickyJimmy). Credit to Lachlan O'Dea for the Capture OU segment.
# Note: I haven't made this as generic as some others. You may need to spend some time working on this

clear
Get-PSSession | Remove-PSSession
$export= 'E:\hasmailboxorlicense.csv'

#Login
Write-Host "---Exchange Online Credentials---" -foregroundcolor green -backgroundcolor gray
    $correctOnline =$False
    $count = 0

    do{
        $OnlineCred = Get-Credential -Message 'Please provide Exchangeonline credentials'
        try{
            Connect-MsolService -Credential $OnlineCred -Verbose
            Write-Host "---MsolService Session started---" -foregroundcolor green -backgroundcolor gray
            $proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
            $OnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $OnlineCred -Authentication Basic -AllowRedirection -SessionOption $proxysettings
            Import-PSSession $OnlineSession
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

#Select Region
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

                [array]$regions=Get-adorganizationalunit -SearchBase 'OU=Users,OU=Accounts,OU=root,DC=OU,DC=zone' -SearchScope Onelevel -filter * | Select-Object DistinguishedName,Name

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

#Test if user has mailbox and license

$users = Get-ADUser -SearchBase $OUPath -Filter *

ForEach($user in $users)
{
    $mailboxtest = $False
    $licensetest = $False
    $m = ''
    $l = ''
    do{
        try{
            $MailboxPresent = Get-Mailbox $user.UserPrincipalName

            if($MailboxPresent -ne $null){
            $m = 'Yes'
            $mailboxtest = $true
            }

            else {
            $m = 'No'
            $mailboxtest = $true
            }
        }
        catch{
        $m = 'Error'
        $mailboxtest = $true
    }
    } Until($mailboxtest -eq $true)

    do{
        try{
            if(Get-MsolUser -UserPrincipalName $user.UserPrincipalName | Where-Object { $_.isLicensed -eq "TRUE" }){
            $l = 'Yes'
            $licensetest = $true
            }
            else{
            $l = 'No'
            $licensetest = $true
            }
        }
        Catch{
           $l = 'Error'
           $licensetest = $true

        }
    } Until( $licensetest = $true)

#Write to csv
    [pscustomobject]@{'User' = $user; 'UserPrincipal' = $user.UserPrincipalName; 'Has Mailbox' = $m; 'Has Office License' = $l;} | Export-CSV -Path $export -Append -NoTypeInformation
  }
