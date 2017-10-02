#Batch Operations
#To be used when onboarding many users. 
#Create a .txt called "UserList" in C:\PS\
#IE : C:\PS\UserList.txt\
#Each line of UserList should be a users account name 
#IE : 
#james.rule
#daniel.marshall

#-
$DotUsers =@()
$UserList = Get-Content "C:\PS\UserList.txt"
Foreach ($User in $UserList){
   $User = $User.replace(' ','.') 
   $DotUsers += $User
}
$DotUsers | Out-File "C:\PS\UserList.txt"

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#--Batch operation - Check Account Names--
$UserCount = 0
$UserList = Get-Content "C:\PS\UserList.txt"

Foreach ($User in $UserList){
    try{     
        $check = 1 
        Write-Host "`nChecking AD for $User " -foregroundcolor yellow
        $check = get-aduser $User    
        if ($check) {
        
            Write-Host "User Found : $check" -foregroundcolor green
            $UserCount++
        }            
    }
    catch{
        echo $_.Exception |format-list -force 
    }    
}
Write-Host "`nConfirmed $UserCount users`n" 

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#-- Staff Onboarding - Batch operation - Assign Specific Licenses
#ENSURE YOU CHANGE THE UsageLocation BEFORE USING THIS SCRIPT 
Connect-Msolservice
$SkypeOnlineDisabled = New-MsolLicenseOptions -AccountSkuId "WiseTechGlobal:ENTERPRISEPACK" -DisabledPlans "MCOSTANDARD"
$SkypeAndTeamsDisabled = New-MsolLicenseOptions -AccountSkuId "WiseTechGlobal:ENTERPRISEPACK" -DisabledPlans "MCOSTANDARD","TEAMS1"
$SkypeAndExchangeDisabled = New-MsolLicenseOptions -AccountSkuId "WiseTechGlobal:ENTERPRISEPACK" -DisabledPlans "MCOSTANDARD","EXCHANGE_S_ENTERPRISE"

$UserList = Get-Content "C:\PS\UserList.txt"

Foreach ($User in $UserList){
    $Userupn = "$User" + "@wisetechglobal.com"
    Write-host "$Userupn being licensed..."
    Set-MsolUser -UserPrincipalName $Userupn -UsageLocation TW
    Set-MsolUserLicense -UserPrincipalName $Userupn -AddLicenses "WiseTechGlobal:ENTERPRISEPACK" -LicenseOptions $SkypeAndExchangeDisabled
    write-host "$User Licensed"
    }

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#-- Staff Onboarding - Batch operation - Select and move all Users to specific OU

do{
    $attempts = $false
    $locationset=$false    
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

            $input = read-host 'Office Selection, if no office required; enter "n" to use $region OU'
            if($input -eq 'n'){
                $OUPath=$region.distinguishedname
                $locationset=$true
                Write-host `n 'Location =' $OUPath | Select-Object Name
                break
            }
            
            else {
                $office=$offices[$input]
                Write-host `n 'Office =' $office `n | Select-Object Name 
            }
        
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

            $input=read-host 'Team Selection, if no team required; enter "n" to use $office OU'            
            
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

$UserList = Get-Content "C:\PS\UserList.txt"
Foreach ($User in $UserList){
    try{        
        $ADUser = get-aduser $User    
        write-host "$User moving to $OUPath"
        move-adobject $ADUser -TargetPath $OUPath 
        write-host "`n$User moved to $OUPath"
    }
    catch{
        echo $_.Exception |format-list -force 
    }    
}

#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#-- Staff Onboarding - Batch operation - Enable on Skype For Business Server
#Attempt 1 : WILL SHOW AS "PERMISSION DENIED" BUT ACTUALLY DOES ENABLE THE USER 
#Attempt 2: WILL CONFIRM IF MOVED "OBject with Identity ... was not changed"

$attemps = $false
$correct = $false
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
        $UserList = Get-Content "C:\PS\UserList.txt"
        Foreach ($User in $UserList){
            try{        
                 enable-csuser -identity "$User@wisetechglobal.com" -RegistrarPool $pool.Identity -sipaddress "sip:$User@wisetechglobal.com" -verbose:$false
                 set-csuser -Identity "$User@wisetechglobal.com" -Enterprisevoiceenabled $true -verbose:$false
                 $correct = $true
                 Write-Host "---Skype Account Succesful---" -foregroundcolor green -backgroundcolor gray
            }
            catch{
            echo $_.Exception |format-list -force 
            }    
        }
    }

    Catch{
        Write-Host "Skype creation failed, waiting and trying again" -foregroundcolor green -backgroundcolor red
        Start-Sleep -s 10
        $count++
        if ($count -eq 3){
            Write-Host "Remote Skype creation failed 3 times" -foregroundcolor green -backgroundcolor red 
            $attemps = $true
            break loop
        }
    } 
} Until($correct -eq $true) 

$complete = $true