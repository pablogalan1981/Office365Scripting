  
<#
.SYNOPSIS
    This script will reset all user mailbox password (except for the admin) to a random value and will export them to a CSV file
    
.DESCRIPTION    

.NOTES
	Author		    Pablo Galan Sabugo <pablogalan1981@gmail.com> 
	Date		      Nov/2018
	Disclaimer: 	This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
  Version: 1.1
#>
######################################################################################################################################################
# Helper functions
######################################################################################################################################################

function Helper-GenerateRandomTempFilename([string]$identifier)
{
    $filename = $env:temp + "\MigrationWiz-"
    if($identifier -ne $null -and $identifier.Length -ge 1)
    {
        $filename += $identifier + "-"
    }
    $filename += (Get-Date).ToString("yyyyMMddHHmmss")
    $filename += ".csv"

    return $filename
}

function Helper-WriteDebug([string]$line)
{
    if($debug)
    {
        Write-Host -Object  ("DEBUG: $line")
    }
}


function Helper-PromptConfirmation([string]$prompt)
{
    while($true)
    {
        $confirm = Read-Host -Prompt ($prompt + " [Y]es or [N]o")

        if($confirm -eq "Y")
        {
            return $true
        }

        if($confirm -eq "N")
        {
            return $false
        }
    }
}

function Helper-GeneratePassword()
{
    $upperCaseChars = "ABCDEFGHIJKLMNPQRSTUVWXYZ"
    $lowerCaseChars = "abcdefghijkmnopqrstuvwxyz"
    $numericChars = "23456789"
    $symbolChars = "-=!@#$%^&*()_+"

    $password = ""

    $rand = New-Object -TypeName System.Random
    1..1 | ForEach-Object -Process { $password = $password + $upperCaseChars[$rand.next(0,$upperCaseChars.Length-1)] }
    1..7 | ForEach-Object -Process { $password = $password + $lowerCaseChars[$rand.next(0,$lowerCaseChars.Length-1)] }
    1..3 | ForEach-Object -Process { $password = $password + $numericChars[$rand.next(0,$numericChars.Length-1)] }
    1..1 | ForEach-Object -Process { $password = $password + $symbolChars[$rand.next(0,$symbolChars.Length-1)] }

    return $password
}

######################################################################################################################################################
# Generate random password for all users
######################################################################################################################################################

function Action-Office365SetUserPasswordsRandom
{
    $count = 0
    $filename = Helper-GenerateRandomTempFilename -identifier "Office365UserPasswords"

    write-host 
    $msg = "#####################################################################################################################################################`
                       CONNECTION TO OFFICE 365 TENANT             `
#####################################################################################################################################################"
    Write-Host $msg

    Connect-MsolService -Credential (Office365Helper-GetCredentials)
    $adminUpn = $script:o365Creds.UserName

    Write-Host
    $tenantDomain = Office365Helper-GetTenantDomain -Credentials $script:o365Creds
    $tenantDomainName = $tenantDomain.replace(".onmicrosoft.com","")
    $msg = "SUCCESS: Connection to  Office 365 '$tenantDomain' Remote PowerShell."
    Write-Host -ForegroundColor Green  $msg

        write-host 
    $msg = "#####################################################################################################################################################`
                       RESET OFFICE 365 USERS PASSWORDS      `
#####################################################################################################################################################"
    Write-Host $msg

    Write-Host
    $forceChange = (Helper-PromptConfirmation -prompt "Would you like to user to be forced to change the password on first login?")

    Write-Host
    Write-Host -Object ("Changing all Office 365 user passwords to something random except $adminUpn")
    
    Write-Host
    Write-Host -Object ("Passwords will be saved to $filename")

    $csv = "UserPrincipalName,Password`r`n"
    $file = New-Item -Path $filename -ItemType file -force -value $csv

    $users = @(Get-MsolUser -All)
    if($users -ne $null)
    {
        foreach($user in $users)
        {
            $count++
            Write-Progress -Activity ("Setting Office 365 user password (" + $count + "/" + $users.Length + ")") -Status $user.DisplayName -PercentComplete ($count/$users.Length*100)

            if($user.UserPrincipalName.ToLower() -ne $adminUpn.ToLower())
            {
                $userPrincipalName = $user.UserPrincipalName
                $password = Helper-GeneratePassword
                Helper-WriteDebug -line ("UPN = $userPrincipalName, Password = $password")

                $result = Set-MsolUserPassword -ObjectId $user.ObjectId.ToString() -NewPassword $password –ForceChangePassword $forceChange

                $csv = ""
                $csv += '"' + $userPrincipalName + '"' + ','		# UserPrincipalName
                $csv += '"' + $password + '"'						# Password

                Add-Content -Path $filename -Value $csv
            }
        }
    }

    if($filename) {
        try {
            Start-Process -FilePath $filename
        }catch {
            $msg = "ERROR: Failed to find the CSV file '$filename'."    
            Write-Host -ForegroundColor Red  $msg
            return
        }   
    }
}

######################################################################################################################################################
# Office 365 helper functions
######################################################################################################################################################

function Office365Helper-GetCredentials()
{
    if($script:o365Creds -eq $null)
    {
        # prompt for credentials
        $script:o365Creds = $host.ui.PromptForCredential("Office 365 Credentials", "Enter your Office 365 administrative user name and password", "", "")
    }

    return $script:o365Creds
}

function Helper-LoadOffice365Module()
{
    if((Get-Module -Name MSOnline) -eq $null)
    {
        Helper-WriteDebug -line ("Office 365 PowerShell module was not loaded")

        Import-Module -Name MSOnline -ErrorAction SilentlyContinue
        if((Get-Module -Name MSOnline) -eq $null)
        {
            Helper-WriteDebug -line ("Office 365 PowerShell module was not found")
            Start-Process -FilePath $o365PowerShellDownloadLink
            throw ("The Office 365 PowerShell module was not found.  Download and install the Microsoft Online Services Sign-In Assistant and the Microsoft Online Services Module for Windows PowerShell from " + $o365PowerShellDownloadLink)
        }
        else
        {
            Helper-WriteDebug -line ("Office 365 PowerShell module successfully loaded")
        }
    }
    else
    {
        Helper-WriteDebug -line ("Office 365 PowerShell module is already loaded")
    }
}

# Function to get the tenant domain
Function Office365Helper-GetTenantDomain {
    param 
    (      
        [parameter(Mandatory=$true)] [Object]$Credentials

    )

    try {
	    Connect-MsolService -Credential $Credentials -ErrorAction Stop
	    
        $tenantDomain = @((Get-MsolDomain |?{$_.Name -match '.onmicrosoft.com' -and $_.Name -notmatch '.mail.'}).Name)

        if($tenantDomain.Count -le 2) {
            $geoLocations = @("APC";"AUS";"CAN";"EUR";"FRA";"IND";"JPN";"KOR";"NAM";"ZAF";"ARE";"GBR")

            foreach ($domain in $tenantDomain) {
                foreach ($geoLocation in $geoLocations) {
                    if ($domain -match $geoLocation) {
                        switch ($geoLocation) {
                            "APC" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in Asia-Pacific"
                            }
                            "AUS" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in Australia"
                            }
                            "CAN" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in Canada"
                            }
                            "EUR" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in Europe"
                            }
                            "FRA" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in France"
                            }
                            "IND" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in India"
                            }
                            "JPN" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in Japan"
                            }
                            "KOR" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in Korea"
                            }
                            "NAM" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in North America"
                                $tenantDomain = $tenantDomain | Where-Object { $_ –ne $domain }
                            }
                            "ZAF" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in South Africa"
                            }
                            "ARE" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in United Arab Emirates"
                            }
                            "GBR" {
                                write-host -ForegroundColor Yellow "WARNING: Office 365 tenant Multi-Geo. Domain '$domain' in United Kingdom"
                            }                           
                        }
                    }
                }      
            }

            write-host "INFO: Main tenant domain '$tenantDomain'" 
        }        
    }
    catch {
	    $msg = "ERROR: Failed to connect to Azure Active Directory to get the tenant domain."
        Write-Host $msg -ForegroundColor Red
        Log-Write -Message $msg 
        
        do {
            $tenantDomain = Read-Host -Prompt ("Enter tenant domain or [C] to cancel")
        } while ($tenantDomain -ne "C" -and $tenantDomain -eq "")

        if ($tenantDomain -eq "C") {
            Exit
        }
	}

    Return $tenantDomain
}

######################################################################################################################################################
# Main menu
######################################################################################################################################################

Write-Host
if(Helper-PromptConfirmation -prompt "Are you sure you want to change all Office 365 user passwords (except for admin) to something random?")
{
    Helper-LoadOffice365Module
    Action-Office365SetUserPasswordsRandom
}
