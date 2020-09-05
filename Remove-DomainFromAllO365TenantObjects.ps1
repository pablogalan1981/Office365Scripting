<#
.SYNOPSIS
    This script that removes a validated domain from all the objects of an Ofice 365 tenant and after that 
    tries to delete the domain from the tenant.

.DESCRIPTION
    The script will
        - Prompt for the source Office 365 global admin credentials to connect to Office 365 and Azure Active Directory.
        - Display all the verified domains in the source Office 365 tenant for you to select the one you want to delete. Only one at a time.​
        - Ask for the new default domain in the source tenant after the domain removal.​
        - Rename the UserPrincipalNames of Msol users to the selected domain, usually the onmicrosoft.com tenant domain.
        - Remove the EmailAddresses with the domain to delete from all Exchange Online recipients.
	
.NOTES
	Author			Pablo Galan Sabugo <pablogalan1981@gmail.com>
	Date		    May/2019
	Disclaimer: 	This script is provided 'AS IS'. No warrantee is provided either expressed or implied. 
    BitTitan cannot be held responsible for any misuse of the script.
    Version: 1.1
#>

#######################################################################################################################
#                                               FUNCTIONS
#######################################################################################################################

# Function to check the AzureAD and MSOnline module
Function Import-PowerShellModules{
    if (!(((Get-Module -Name "MSOnline") -ne $null) -or ((Get-InstalledModule -Name "MSOnline" -ErrorAction SilentlyContinue) -ne $null))) {
        Write-Host
        $msg = "INFO: MSOnline PowerShell module not installed."
        Write-Host $msg     
        $msg = "INFO: Installing MSOnline PowerShell module."
        Write-Host $msg

        Sleep 5
    
        try{
            Install-Module -Name MSOnline -force -ErrorAction Stop
        }
        catch{
            $msg = "ERROR: Failed to install MSOnline module. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Write-Host
            $msg = "ACTION: Run this script 'As administrator' to intall the MSOnline module."
            Write-Host -ForegroundColor Yellow $msg
            Exit
        }
        Import-Module MSOnline
    }

    if (!(((Get-Module -Name "AzureAD") -ne $null) -or ((Get-InstalledModule -Name "AzureAD" -ErrorAction SilentlyContinue) -ne $null))) {
        Write-Host
        $msg = "INFO: AzureAD PowerShell module not installed."
        Write-Host $msg     
        $msg = "INFO: Installing AzureAD PowerShell module."
        Write-Host $msg

        Sleep 5
    
        try{
            Install-Module -Name AzureAD -force -ErrorAction Stop
        }
        catch{
            $msg = "ERROR: Failed to install AzureAD module. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Write-Host
            $msg = "ACTION: Run this script 'As administrator' to intall the AzureAD module."
            Write-Host -ForegroundColor Yellow $msg
            Exit
        }
        Import-Module AzureAD
    }
}

# Function to create the working and log directories
Function Create-Working-Directory {    
    param 
    (
        [CmdletBinding()]
        [parameter(Mandatory=$true)] [string]$workingDir,
        [parameter(Mandatory=$true)] [string]$logDir
    )
    if ( !(Test-Path -Path $workingDir)) {
		try {
			$suppressOutput = New-Item -ItemType Directory -Path $workingDir -Force -ErrorAction Stop
            $msg = "SUCCESS: Folder '$($workingDir)' for CSV files has been created."
            Write-Host -ForegroundColor Green $msg
		}
		catch {
            $msg = "ERROR: Failed to create '$workingDir'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
		}
    }
    if ( !(Test-Path -Path $logDir)) {
        try {
            $suppressOutput = New-Item -ItemType Directory -Path $logDir -Force -ErrorAction Stop      

            $msg = "SUCCESS: Folder '$($logDir)' for log files has been created."
            Write-Host -ForegroundColor Green $msg 
        }
        catch {
            $msg = "ERROR: Failed to create log directory '$($logDir)'. Script will abort."
            Write-Host -ForegroundColor Red $msg
            Exit
        } 
    }
}

# Function to write information to the Log File
Function Log-Write {
    param
    (
        [Parameter(Mandatory=$true)]    [string]$Message
    )
    $lineItem = "[$(Get-Date -Format "dd-MMM-yyyy HH:mm:ss") | PID:$($pid) | $($env:username) ] " + $Message
	Add-Content -Path $logFile -Value $lineItem
}

# Function to create EXO PowerShell session
Function Connect-O365Tenant {
    
    #Prompt for Office 365 global admin Credentials
    $msg = "INFO: Connecting to Exchange Online."
    Write-Host $msg
    Log-Write -Message $msg 

    if (!($o365Session.State)) {
        try {
            $loginAttempts = 0
            do {
                $loginAttempts++

                # Connect to Source Exchange Online via MSPC endpoint
                if($useMspcEndpoints) {

                    #Select source endpoint
                    $exportEndpointId = Select-MSPC_Endpoint -CustomerOrganizationId $customerOrganizationId -ExportOrImport "source" -EndpointType "ExchangeOnline2"
                    #Get source endpoint credentials
                    [PSObject]$exportEndpointData = Get-MSPC_EndpointData -CustomerOrganizationId $customerOrganizationId -EndpointId $exportEndpointId 

                    #Create a PSCredential object to connect to source Office 365 tenant
                    $srcAdministrativeUsername = $exportEndpointData.AdministrativeUsername
                    $srcAdministrativePassword = ConvertTo-SecureString -String $($exportEndpointData.AdministrativePassword) -AsPlainText -Force
                    $O365Creds = New-Object System.Management.Automation.PSCredential ($srcAdministrativeUsername, $srcAdministrativePassword)
                }
                # Connect to Source Exchange Online via manual credentials entry
                else {
                    # Connect to Exchange Online
                    $O365Creds = Get-Credential -Message "Enter Your Office 365 Admin Credentials."
                    if (!($O365Creds)) {
                         $msg = "ERROR: Cancel button or ESC was pressed while asking for Credentials. Script will abort."
                         Write-Host -ForegroundColor Red  $msg
                         Log-Write -Message $msg 
                         Exit
                    }
                 }

                $o365Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365Creds -Authentication Basic -AllowRedirection -ErrorAction Stop -WarningAction SilentlyContinue
                $result = Import-PSSession -Session $o365Session -AllowClobber -ErrorAction Stop -WarningAction silentlyContinue -DisableNameChecking
                
                $msg = "SUCCESS: Connection to Exchange Online."
                Write-Host -ForegroundColor Green  $msg
                Log-Write -Message $msg 
            }
            until (($loginAttempts -ge 3) -or ($($o365Session.State) -eq "Opened"))

            # Only 3 attempts allowed
            if($loginAttempts -ge 3) {
                $msg = "ERROR: Failed to connect to the Office 365. Review your Office 365 admin credentials and try again."
                Write-Host $msg -ForegroundColor Red
                Log-Write -Message $msg 
                Start-Sleep -Seconds 5
                Exit
            }
        }
        catch {
            $msg = "ERROR: Failed to connect to Office 365."
            Write-Host -ForegroundColor Red $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Get-PSSession | Remove-PSSession
            Exit
        }

        try{
            # Connect to Azure Active Directory (AD) using the Microsoft Azure Active Directory Module for Windows PowerShell module
            $msg = "INFO: Connecting to MsolService."
            Write-Host $msg
            Log-Write -Message $msg 

            Connect-MsolService -Credential $O365Creds    
            
            $msg = "SUCCESS: Connection to MsolService."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg  
        }
        catch {
            $msg = "ERROR: Failed to connect to MsolService."
            Write-Host -ForegroundColor Red $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message
            Get-PSSession | Remove-PSSession
            Exit
        }

        try{
            # Connect to Azure Active Directory (AD) using the Azure Active Directory PowerShell for Graph module.
            $msg = "INFO: Connecting to Azure Active Directory."
            Write-Host $msg
            Log-Write -Message $msg 

            Connect-AzureAD -Credential $O365Creds    
            
            $msg = "SUCCESS: Connection to Azure Active Directory."
            Write-Host -ForegroundColor Green  $msg
            Log-Write -Message $msg      
        }
        catch {
            $msg = "ERROR: Failed to connect to Azure Active Directory."
            Write-Host -ForegroundColor Red $msg
            Log-Write -Message $msg 
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $_.Exception.Message                     
        }
        
        return $O365Creds 
    } 
    else {
        Get-PSSession | Remove-PSSession
    }
}

# Function to create EXO PowerShell session
Function Connect-DestinationO365Tenant {
    
    #Prompt for Office 365 global admin Credentials
     try{

         # Connect to destination MsolService
         $O365Creds = Get-Credential -Message "Enter Your DESTINATION Office 365 Admin Credentials."
         if (!($O365Creds)) {
              $msg = "ERROR: Cancel button or ESC was pressed while asking for Credentials. Script will abort."
              Write-Host -ForegroundColor Red  $msg
              Log-Write -Message $msg 
              Exit
         }
         
         Write-Host
         # Connect to destination Azure Active Directory (AD) using the Microsoft Azure Active Directory Module for Windows PowerShell module
         $msg = "INFO: Connecting to destination MsolService."
         Write-Host $msg
         Log-Write -Message $msg 
 
         Connect-MsolService -Credential $O365Creds    
     
         $msg = "SUCCESS: Connection to MsolService."
         Write-Host -ForegroundColor Green  $msg
         Log-Write -Message $msg  
     }
     catch {
         $msg = "ERROR: Failed to connect to MsolService."
         Write-Host -ForegroundColor Red $msg
         Log-Write -Message $msg 
         Write-Host -ForegroundColor Red $_.Exception.Message
         Log-Write -Message $_.Exception.Message
         Get-PSSession | Remove-PSSession
         Exit
     }
    
     return $O365Creds 

}

# Function to get the tenant domain
 Function Get-TenantDomain {
    param 
    (      
        [parameter(Mandatory=$true)] [Object]$Credentials

    )

    try {  
        $tenantDomain = @((Get-MsolDomain |?{$_.Name -match 'onmicrosoft.com'}).Name)[0]

    }
    catch {
	    $msg = "ERROR: Failed to connect to Azure Active Directory."
        Write-Host $msg -ForegroundColor Red
        Log-Write -Message $msg 
        Start-Sleep -Seconds 5
        Exit
	}

    Return $tenantDomain
}

# Function to get all validated domains in the tenant
Function Get-VanityDomains {
    param 
    (      
        [parameter(Mandatory=$true)] [Object]$Credentials

    )

    try {  
        $tenantVanityDomains = @(Get-MsolDomain -Status Verified |? {$_.Name -notmatch 'onmicrosoft.com'}).Name
    }
    catch {
	    $msg = "ERROR: Failed to connect to Azure Active Directory."
        Write-Host $msg -ForegroundColor Red
        Log-Write -Message $msg 
        Start-Sleep -Seconds 5
        Exit
	}

    Return $tenantVanityDomains
}

# Function to select the domain to delete
Function Select-Domain {
    param 
    (      
        [parameter(Mandatory=$true)] [Object]$Credentials,
        [parameter(Mandatory=$false)] [Boolean]$DisplayAll

    )

    $tenantDomain = Get-TenantDomain -Credentials $Credentials
    $vanityDomains = @(Get-VanityDomains -Credentials $Credentials)
    $domainLength = $vanityDomains.Length

    #######################################
    # {Prompt for the domain to delete
    #######################################
    
    if($vanityDomains -ne $null) {
        if(!$all) {
            Write-Host
            Write-Host -ForegroundColor Yellow -Object "ACTION: Select a verified domain to delete:" 
        }
        else {
            Write-Host
            Write-Host -Object "INFO: Current domains added to the Office 365 tenant:" 
        }


        for ($i=0; $i -lt $domainLength; $i++) {
            $vanityDomain = $vanityDomains[$i]
            Write-Host -Object $i,"-",$vanityDomain
        }

        if(!$all) {
            Write-Host -Object "a - Add domain to Office 365 tenant"
            Write-Host -Object "x - Exit"
            Write-Host

            do {
                if($domainLength -eq 1) {
                    $result = Read-Host -Prompt ("Select 0, a or x")
                }
                else {
                    $result = Read-Host -Prompt ("Select 0-" + ($domainLength-1) + ", a or x")
                }

                if($result -eq "a") {
                    Add-Office365Domain
                }
                elseif($result -eq "x") {
                    Exit
                }
                elseif(($result -match "^\d+$") -and ([int]$result -ge 0) -and ([int]$result -lt $domainLength)) {
                    $vanityDomain = $vanityDomains[$result]
                    Return $vanityDomain
                }
            }
            while($true)
        }
    }
    else{
        Write-Host
        Write-Host -ForegroundColor Red "INFO: There is no domain attached to the Office 365 tenant. The default domain is $tenantDomain." 
        Write-Host
        Exit
    }
}

# Function to select the domain to delete
Function Select-NewDefaultDomain {
    param 
    (      
        [parameter(Mandatory=$true)] [Object]$Credentials,
        [parameter(Mandatory=$true)] [Object]$DomainToDelete
    )

    $tenantDomain = Get-TenantDomain -Credentials $Credentials
    $domains = @(Get-VanityDomains -Credentials $Credentials | Where-Object {$_ -notmatch $DomainToDelete})
    $domains += $tenantDomain
    $domainsCount = $domains.Count

    #######################################
    # {Prompt for the domain to delete
    #######################################
    
    if($domains -ne $null) {
        Write-Host
        Write-Host -ForegroundColor Yellow -Object "ACTION: Select the domain to set as default:" 

        for ($i=0; $i -lt $domainsCount; $i++) {
            $newDefaultDomain = $domains[$i]
            Write-Host -Object $i,"-",$newDefaultDomain
        }

        Write-Host

        do {
            if($domainsCount -eq 0) {
                Return $tenantDomain
            }
            else {
                $result = Read-Host -Prompt ("Select 0-" + ($domainsCount))

                 if(([int]$result -ge 0) -and ([int]$result -lt $domainsCount)) {
                    $newDefaultDomain = $domains[$result]
                    Return $newDefaultDomain
                }
            }
        }
        while($true)
    }
    else{
        Write-Host
        Write-Host -ForegroundColor Red "INFO: There is no other domain attached to the Office 365 tenant. The default domain will be $tenantDomain." 
        Return $tenantDomain
    }
}

Function Check-MsolUsersWithDomain {
    param 
    (      
        [parameter(Mandatory=$false)] [Object]$TenantDomain,
        [parameter(Mandatory=$false)] [Object]$Domain,
        [parameter(Mandatory=$true)] [Object]$Credentials

    )

    if($TenantDomain -eq $null) {
        $TenantDomain = Get-TenantDomain -Credentials $Credentials
    }

    if($Domain -eq $null) {
        $Domain = Select-Domain -Credentials $Credentials
    }

    $adminUserName = $Credentials.UserName
    $adminUserNamePrefix = $Credentials.UserName.Split("@")[0]
    
    $msolUsersWithDomain = @()
    $msolUsersProxyAddressesWithDomain = @()
    $dataout = 0

    Write-Host
    Write-Host "INFO: Exporting Msol users from Azure Active Directory with domain '$domain'." 
    try{
        $msolUsersWithDomain = @(Get-MsolUser -All | Where-Object {$_.UserPrincipalName -match $domain -and $_.UserPrincipalName -notmatch $adminUserNamePrefix} | Select-Object DisplayName, UserPrincipalName, ProxyAddresses, objectId | Sort-Object -Property DisplayName)
        $msolUsersProxyAddressesWithDomain = @(Get-MsolUser -All | Where-Object {$_.ProxyAddresses -match $domain -and $_.ProxyAddresses -notmatch $adminUserNamePrefix} | Select-Object DisplayName, UserPrincipalName, ProxyAddresses, objectId | Sort-Object -Property DisplayName)

        $msolUsersWithDomainCount = $msolUsersWithDomain.Count
        $msolUsersProxyAddressesWithDomainCount = $msolUsersProxyAddressesWithDomain.Count

        if($msolUsersWithDomainCount -gt 0 -or $msolUsersProxyAddressesWithDomainCount -gt 0) {

            if($msolUsersWithDomainCount -gt 0 -and $msolUsersProxyAddressesWithDomainCount -gt 0) {
                Write-Host -ForegroundColor Green "SUCCESS: $msolUsersWithDomainCount Msol users have been found with domain '$domain' in UserPrincipalName and ProxyAddresses."
            }
            elseif($msolUsersWithDomainCount -gt 0 -and $msolUsersProxyAddressesWithDomainCount -eq 0) {
                Write-Host -ForegroundColor Green "SUCCESS: $msolUsersWithDomainCount Msol users have been found with domain '$domain' in UserPrincipalName."
            }
            elseif($msolUsersWithDomainCount -eq 0 -and $msolUsersProxyAddressesWithDomainCount -gt 0) {
                Write-Host -ForegroundColor Green "SUCCESS: $msolUsersProxyAddressesWithDomainCount Msol users have been found with domain '$domain' in ProxyAddresses."
            }
                        
            #$msolUsersWithDomain | Format-Table DisplayName, UserPrincipalName, ProxyAddresses
            
            #Export msol users with domain to CSV file
            do {
                try {
                    if($msolUsersWithDomainCount -gt 0) {
                        $msolUsersWithDomain | Select-Object DisplayName, UserPrincipalName, @{ n='ProxyAddresses'; e={ $_.ProxyAddresses -join ';' } } | Export-Csv -Path $workingDir\msolUsersWithDomain-$domain.csv -NoTypeInformation -force 
                    }
                    elseif($msolUsersWithDomainCount -eq 0 -and $msolUsersProxyAddressesWithDomainCount -gt 0) {
                        $msolUsersProxyAddressesWithDomain | Select-Object DisplayName, UserPrincipalName, @{ n='ProxyAddresses'; e={ $_.ProxyAddresses -join ';' } } | Export-Csv -Path $workingDir\msolUsersWithDomain-$domain.csv -NoTypeInformation -force 
                    }
                    $msg = "SUCCESS: CSV file '$workingDir\msolUsersWithDomain-$domain.csv' processed, exported and open."
                    Write-Host -ForegroundColor Green $msg 
                    Log-Write -Message $msg

                    Break
                }
                catch {
                    $msg = "WARNING: Close opened CSV file '$workingDir\msolUsersWithDomain-$domain.csv'."
                    Write-Host -ForegroundColor Yellow $msg
                    Log-Write -Message $msg 
                    Write-Host

                    Start-Sleep 5
                }
            } while ($true) 

            try {
                Start-Process -FilePath $workingDir\msolUsersWithDomain-$domain.csv
            }catch {
                $msg = "ERROR: Failed to find the CSV file '$workingDir\msolUsersWithDomain-$domain.csv'."    
                Write-Host -ForegroundColor Red  $msg
                Exit
            }  

            do {
                $confirm = (Read-Host "ACTION: If you have reviewed the CSV file then press [C] to continue" ) 
            } while($confirm -ne "C")

        }
        else {
            $msg = "INFO: No Msol users have been found with UserPrincipalName domain '$domain'."
            Write-Host $msg -ForegroundColor Red
            Log-Write -Message $msg   
            $Global:MsolUserDomain = $true
        }
    }
    catch {
        $msg = "ERROR: Failed to get the MsolUsers with domain '$domain'."
        Write-Host -ForegroundColor Red  $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $msg
        Log-Write -Message $_.Exception.Message   
    }
    
    if($msolUsersWithDomainCount -gt 0 -or $msolUsersProxyAddressesWithDomainCount -gt 0) {

        $script:newDefaultDomain = Select-NewDefaultDomain -Credentials $Credentials -DomainToDelete $domain
        
        if($msolUsersWithDomainCount -gt 0) {
            do {
                    Write-Host
                    $msg = "INFO: MsolUsers with domain '$domain' will be renamed to '$newDefaultDomain'."
                    Write-Host  $msg
                    Log-Write -Message $msg

                $confirm = (Read-Host "ACTION: Do you want to remove domain '$domain' from Msol UserPrincipalNames?  [Y]es or [N]o" ) 
            } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))
        }

        if($confirm -eq "Y") {
            try {
                #removing domain from UserMailbox
                $msolUsersWithDomain | % {
                    Set-MsolUserPrincipalName -ObjectId $_.objectId -NewUserPrincipalName ($_.UserPrincipalName.Split("@")[0]+"@"+$newDefaultDomain); 
                    
                    $dataout += 1 ; 

                    Write-Progress -Activity ("Processing Msol Users") -Status "DisplayName: $($_.DisplayName) UserPrincipalName: $($_.UserPrincipalName)."

                }
                #removing domain from User

                Write-Progress -Activity " " -Completed

                if($dataout -ne 0) {
                    $msg = "SUCCESS: $dataout MsolUsers with domain '$domain' renamed to '$TenantDomain'."
                    Write-Host $msg -ForegroundColor Green
                    Log-Write -Message $msg
                } 
            }
            catch{
                $msg = "INFO: Failed to remove domain '$domain' from  Msol UserPrincipalNames."
                Write-Host $msg -ForegroundColor Red
                Log-Write -Message $msg   
                Exit
            }
            
            Write-Host
            $msg = "INFO: Waiting 5 minutes for change replication."
            Write-Host $msg -ForegroundColor Yellow
            Log-Write -Message $msg
            Start-Sleep -Seconds 300
        }

        $Global:MsolUserDomain = $true
    }
}

Function Check-MsolGroupsWithDomain {
    param 
    (      
        [parameter(Mandatory=$false)] [Object]$TenantDomain,
        [parameter(Mandatory=$false)] [Object]$Domain,
        [parameter(Mandatory=$true)] [Object]$Credentials

    )

    if($TenantDomain -eq $null) {
        $TenantDomain = Get-TenantDomain -Credentials $Credentials
    }

    if($Domain -eq $null) {
        $Domain = Select-Domain -Credentials $Credentials
    }

    $adminUserName = $Credentials.UserName
    $adminUserNamePrefix = $Credentials.UserName.Split("@")[0]
    
    $msolGroupsWithDomain = @()
    $dataout = 0

    Write-Host
    Write-Host "INFO: Exporting Msol groups from Azure Active Directory with domain '$domain'." 
    try{
        $msolGroupsWithDomain = @(Get-MsolGroup -All | Where-Object {$_.EmailAddress -match $domain -or $_.ProxyAddresses -match $domain} | Select-Object DisplayName, EmailAddress, ProxyAddresses, GroupType, objectId | Sort-Object -Property DisplayName)

        $msolGroupsWithDomainCount = $msolGroupsWithDomain.Count

        if($msolGroupsWithDomainCount -gt 0) {
            Write-Host -ForegroundColor Green "SUCCESS: $msolGroupsWithDomainCount Msol groups have been found with domain '$domain' in EmailAddress or ProxyAddresses."
     
            #$msolUsersWithDomain | Format-Table DisplayName, UserPrincipalName, ProxyAddresses

            #Export msol groups with domain to CSV file
            do {
                try {
                    $msolGroupsWithDomain | Select-Object DisplayName, EmailAddress, @{ n='ProxyAddresses'; e={ $_.ProxyAddresses -join ';' }}, GroupType, objectId | Export-Csv -Path $workingDir\msolGroupsWithDomain-$domain.csv -NoTypeInformation -force 

                    $msg = "SUCCESS: CSV file '$workingDir\msolGroupsWithDomain.csv' processed, exported and open."
                    Write-Host -ForegroundColor Green $msg 
                    Log-Write -Message $msg

                    Break
                }
                catch {
                    $msg = "WARNING: Close opened CSV file '$workingDir\msolGroupsWithDomain.csv'."
                    Write-Host -ForegroundColor Yellow $msg
                    Log-Write -Message $msg 
                    Write-Host

                    Start-Sleep 5
                }
            } while ($true)      

            try {
                Start-Process -FilePath $workingDir\msolGroupsWithDomain-$domain.csv
            }catch {
                $msg = "ERROR: Failed to find the CSV file '$workingDir\msolGroupsWithDomain-$domain.csv'."    
                Write-Host -ForegroundColor Red  $msg
                Exit
            }  

        }
        else {
            $msg = "INFO: No Msol groups have been found with domain '$domain'."
            Write-Host $msg -ForegroundColor Red
            Log-Write -Message $msg   
            $Global:MsolGroupDomain = $true
        }
    }
    catch {
        $msg = "ERROR: Failed to get the Msol groups with domain '$domain'."
        Write-Host -ForegroundColor Red  $msg
        Write-Host -ForegroundColor Red $_.Exception.Message
        Log-Write -Message $msg
        Log-Write -Message $_.Exception.Message   
    }
}

Function Check-ExchangeOlineRecipientsWithDomain {
    param 
    (      
        [parameter(Mandatory=$false)] [Object]$TenantDomain,
        [parameter(Mandatory=$false)] [Object]$Domain,
        [parameter(Mandatory=$true)] [Object]$Credentials

    )

    if($TenantDomain -eq $null) {
        $TenantDomain = Get-TenantDomain -Credentials $Credentials
    }

    if($Domain -eq $null) {
        $Domain = Select-Domain -Credentials $Credentials
    }

    $adminUserName = $Credentials.UserName
    $adminUserNamePrefix = $Credentials.UserName.Split("@")[0]

    $exoRecipientsWithDomain = @()
    Write-Host
    Write-Host "INFO: Exporting recipients from Exchange Online with domain '$domain'." 
    try{
        $exoRecipientsWithDomain = @(Get-Recipient -ResultSize Unlimited -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Where-Object {$_.PrimarySmtpAddress -match "$Domain" -or $_.EmailAddresses -match "$Domain"} | Select-Object PrimarySmtpAddress,EmailAddresses,RecipientType,RecipientTypedetails,ExchangeObjectId | Sort-Object -Property RecipientType,RecipientTypedetails)

        $exoRecipientsWithDomainCount = $exoRecipientsWithDomain.Count
        if($exoRecipientsWithDomainCount -ne 0) {
            Write-Host -ForegroundColor Green "SUCCESS: $exoRecipientsWithDomainCount Exchange Online recpients have been found with domain '$domain'."
            
            #$exoRecipientsWithDomain | Format-Table PrimarySmtpAddress,EmailAddresses,RecipientType,RecipientTypedetails -wrap
            
            #Export msol users with domain to CSV file
            do {
                try {
                    $exoRecipientsWithDomain | Export-Csv -Path $workingDir\exoRecipientsWithDomain-$domain.csv -NoTypeInformation -force
                    $msg = "SUCCESS: CSV file '$workingDir\exoRecipientsWithDomain-$domain.csv' processed, exported and open."
                    Write-Host -ForegroundColor Green $msg 
                    Log-Write -Message $msg

                    Break
                }
                catch {
                    $msg = "WARNING: Close opened CSV file '$workingDir\exoRecipientsWithDomain-$domain.csv'."
                    Write-Host -ForegroundColor Yellow $msg
                    Log-Write -Message $msg 
                    Write-Host

                    Start-Sleep 5
                }
            } while ($true) 

            try {
                Start-Process -FilePath $workingDir\exoRecipientsWithDomain-$domain.csv
            }catch {
                $msg = "ERROR: Failed to find the CSV file '$workingDir\exoRecipientsWithDomain-$domain.csv'."    
                Write-Host -ForegroundColor Red  $msg
                Exit
            }  

            do {
                $confirm = (Read-Host "ACTION:  If you have reviewed the CSV file then press [C] to continue" ) 
            } while($confirm -ne "C")

        }
        else {
            $msg = "INFO: No Exchange Online recipients have been found with domain '$domain'."
            Write-Host $msg -ForegroundColor Red
            Log-Write -Message $msg   
            $Global:ExoDomain = $true
        }
    }
    catch {
        $msg = "ERROR: Failed to get the Exchange Online recipients with domain '$domain'."
        Write-Host $msg -ForegroundColor Red
        Log-Write -Message $msg    
    }
    
    if($exoRecipientsWithDomainCount -ne 0) {
        do {
             Write-Host
             $msg = "INFO: Exchange Online recipients with domain '$domain' will be renamed to '$newDefaultDomain'."
             Write-Host  $msg
             Log-Write -Message $msg

            $confirm = (Read-Host "ACTION: Do you want to remove domain '$domain' from Exchange Online recipients?  [Y]es or [N]o" ) 
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        if($confirm -eq "Y") {

            $dataout = 0

            foreach($exoRecipientWithDomain in $exoRecipientsWithDomain) {

                $ExchangeObjectId = $exoRecipientWithDomain.ExchangeObjectId

                $PrimarySmtpAddress = $exoRecipientWithDomain.PrimarySmtpAddress
                if($exoRecipientWithDomain.EmailAddresses){$EmailAddresses = @($exoRecipientWithDomain.EmailAddresses.split(' ').split(",").split(";").trim());}

                $filteredEmailAddresses = ''
                foreach($emailAddress in $EmailAddresses) { 
                    if(($emailAddress -cmatch "smtp:" -or $emailAddress -cmatch "SMTP:") -and ($emailAddress -match ".onmicrosoft.com" -and $emailAddress -notmatch ".mail.onmicrosoft.com")) {
                        $filteredEmailAddresses += $emailAddress.replace("SMTP:","").replace("smtp:","")
                    }
                } 

                $tenantEmailAddress = @($filteredEmailAddresses | Select-Object -Unique)[0]
                if($tenantEmailAddress) {
                    $srcTenantEmailAddress = $tenantEmailAddress
                }

                if($newDefaultDomain -eq $TenantDomain) {
                    $newPrimarySmtpAddress = $srcTenantEmailAddress
                }    
                else{
                    $newPrimarySmtpAddress = ($primarySmtpAddress -split "@")[0]+"@"+$newDefaultDomain
                }

                $filteredEmailAddresses = ''
                foreach($emailAddress in $EmailAddresses) { 
                    if(($emailAddress -cmatch "smtp:" -or $emailAddress -cmatch "SMTP:") -and ($emailAddress -match "$Domain")) {
                        $filteredEmailAddresses += $emailAddress.replace("SMTP:","").replace("smtp:","")
                    }
                } 

                $EmailaddressWithDomain = @($filteredEmailAddresses | Select-Object -Unique)[0]
                if($EmailaddressWithDomain) {
                    $EmailaddressToDelete = $EmailaddressWithDomain
                }
                                                             
                $recipientType = $exoRecipientWithDomain.RecipientType
                $recipientTypeDetails = $exoRecipientWithDomain.RecipientTypeDetails
   
                Write-Progress -Activity ("Processing Exchange Online Recipient") -Status "$PrimarySmtpAddress {$recipientType, $RecipientTypeDetails}."

                try {
                    switch($recipientType) {
                       "UserMailbox" {
                            $ifEmailAddresses = Get-Mailbox $PrimarySmtpAddress | Select-Object EmailAddresses
                            if($ifEmailAddresses) {
                                Set-Mailbox $PrimarySmtpAddress -EmailAddresses  @{remove="SMTP:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                                Set-Mailbox $PrimarySmtpAddress -EmailAddresses  @{remove="smtp:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                                Set-Mailbox $PrimarySmtpAddress -EmailAddresses  @{remove="SIP:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue                                
                                Set-Mailbox $PrimarySmtpAddress -EmailAddresses  @{remove="sip:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue

                                Set-Mailbox $PrimarySmtpAddress -EmailAddresses  @{add="SMTP:$newPrimarySmtpAddress"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                                Set-Mailbox $PrimarySmtpAddress -EmailAddresses  @{add="SIP:$newPrimarySmtpAddress"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue

                                $dataout += 1
                            }
                       }                  
                       "MailUniversalDistributionGroup" {
                            if($recipientTypeDetails -eq "GroupMailbox") {
                                #Set Default Email Policy for Office 365 Groups to force new Groups to use new default domain
                                $emailAddressPolicy = Get-EmailAddressPolicy | Where-Object {$_.name -eq "groups"}
                                if(!$emailAddressPolicy ) {
                                    New-EmailAddressPolicy -Name Groups -IncludeUnifiedGroupRecipients -EnabledEmailAddressTemplates "SMTP:@$newDefaultDomain" -Priority 1
                                }
                                else {
                                    Set-EmailAddressPolicy -Identity $emailAddressPolicy.Identity -EnabledEmailAddressTemplates "SMTP:@$newDefaultDomain" -Priority 1
                                }

                                Set-UnifiedGroup -identity $PrimarySmtpAddress -PrimarySmtpAddress $newPrimarySmtpAddress -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                                Set-UnifiedGroup -identity $PrimarySmtpAddress -EmailAddresses @{remove="smtp:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                                Set-UnifiedGroup -identity $PrimarySmtpAddress -EmailAddresses @{remove="SMTP:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue 
                                
                                $dataout += 1                            
                            }
                            else {
                                Set-DistributionGroup -identity $PrimarySmtpAddress -PrimarySmtpAddress $newPrimarySmtpAddress -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                                Set-DistributionGroup -identity $PrimarySmtpAddress -EmailAddresses @{remove="smtp:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                                Set-DistributionGroup -identity $PrimarySmtpAddress -EmailAddresses @{remove="SMTP:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue

                                $dataout += 1
                            }                    
                       }
                       "MailUniversalSecurityGroup" {
                            Set-DistributionGroup -identity $PrimarySmtpAddress –PrimarySmtpAddress $newPrimarySmtpAddress -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                            Set-DistributionGroup -identity $PrimarySmtpAddress -EmailAddresses @{remove="smtp:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                            Set-DistributionGroup -identity $PrimarySmtpAddress -EmailAddresses @{remove="SMTP:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue

                            $dataout += 1
                       }
                       "DynamicDistributionGroup" {
                            Set-DistributionGroup -identity $PrimarySmtpAddress –PrimarySmtpAddress $newPrimarySmtpAddress -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                            Set-DistributionGroup -identity $PrimarySmtpAddress -EmailAddresses @{remove="smtp:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                            Set-DistributionGroup -identity $PrimarySmtpAddress -EmailAddresses @{remove="SMTP:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue

                            $dataout += 1
                       }
                       "MailUser" {
                            Set-MailUser -identity $PrimarySmtpAddress –PrimarySmtpAddress $newPrimarySmtpAddress -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                            Set-MailUser -identity $PrimarySmtpAddress -EmailAddresses @{remove="smtp:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                            Set-MailUser -identity $PrimarySmtpAddress -EmailAddresses @{remove="SMTP:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue

                            $dataout += 1
                       }
                       "PublicFolder" {
                            Set-MailPublicFolder -identity $PrimarySmtpAddress –PrimarySmtpAddress $newPrimarySmtpAddress -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                            Set-MailPublicFolder -identity $PrimarySmtpAddress -EmailAddresses @{remove="smtp:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
                            Set-MailPublicFolder -identity $PrimarySmtpAddress -EmailAddresses @{remove="SMTP:$EmailaddressToDelete"} -WarningAction SilentlyContinue -ErrorAction SilentlyContinue

                            $dataout += 1
                       }
                    }                      
                }
                catch {
                    $msg = "INFO: Failed to remove domain '$domain' from $PrimarySmtpAddress Exchange Online Recipient {$recipientType, $RecipientTypeDetails}."
                    Write-Host $msg -ForegroundColor Red
                    Log-Write -Message $msg  
                    Write-Host -ForegroundColor Red $($_.Exception.Message)
                    Log-Write -Message $($_.Exception.Message)  
                }  
                
                Write-Progress -Activity " " -Completed                         
            }
            
            if($dataOut) {
                $msg = "SUCCESS: $dataOut Exchange Online Recipients with domain '$domain' have been renamed to '$TenantDomain'."
                Write-Host $msg -ForegroundColor Green
                Log-Write -Message $msg  

                Write-Host
                $msg = "INFO: Waiting 5 minutes for change replication."
                Write-Host $msg -ForegroundColor Yellow
                Log-Write -Message $msg
                Start-Sleep -Seconds 300
            }

            $Global:ExoDomain = $true
        }
    }

}

Function Check-AzureADDomain {
    param 
    (      
        [parameter(Mandatory=$false)] [Object]$Domain,
        [parameter(Mandatory=$true)] [Object]$Credentials

    )

    if($Domain -eq $null) {
        $Domain = Select-Domain -Credentials $Credentials
    }
    
    Write-Host
    Write-Host "INFO: Exporting applications from Azure Active Directory with domain $domain." 
    try{
        $azureAppsWithDomain = Get-AzureADApplication | Where-Object {$_.Homepage -match "$domain" -or $_.IdentifierUris -match "$domain" -or $_.ReplyUrls -match "$domain"} | Select-Object DisplayName, Homepage, IdentifierUris, ReplyUrls, AppId

        $azureAppsWithDomainCount = $azureAppsWithDomain.Count
        if($azureAppsWithDomainCount -ne 0) {
            Write-Host -ForegroundColor Green "SUCCESS: $azureAppsWithDomainCount Azure Apps have been found with domain '$domain'."
            $azureAppsWithDomain | Format-Table DisplayName, Homepage, IdentifierUris, ReplyUrls, AppId
            
            #Export applications from Azure Active Directory with domain to CSV file
            do {
                try {
                    $azureAppsWithDomain | Export-Csv -Path $workingDir\azureAppsWithDomain-$domain.csv -NoTypeInformation -force
                    
                    $msg = "SUCCESS: CSV file '$workingDir\azureAppsWithDomain-$domain.csv' processed, exported and open."
                    Write-Host -ForegroundColor Green $msg 
                    Log-Write -Message $msg

                    Break
                }
                catch {
                    $msg = "WARNING: Close opened CSV file '$workingDir\azureAppsWithDomain-$domain.csv'."
                    Write-Host -ForegroundColor Yellow $msg
                    Log-Write -Message $msg 
                    Write-Host

                    Start-Sleep 5
                }
            } while ($true)   

            do {
                $confirm = (Read-Host "ACTION: If you have reviewed the CSV file then press [C] to continue" ) 
            } while($confirm -ne "C")

            if($msolUsersWithDomainCount -ge 1) {
                Write-Host
                do {
                    $confirm = (Read-Host "ACTION: Do you want to remove the Azure Apps?  [Y]es or [N]o" ) 
                } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

                if($confirm -eq "Y") {
            
                    foreach($azureAppWithDomain in $azureAppsWithDomain) {
                        try {
                            if($azureAppWithDomain.AvailableToOtherTenants -eq $True) {
                                Set-AzureADApplication -ObjectId $azureAppWithDomain.ObjectId -AvailableToOtherTenants $False
                                Remove-AzureADApplication -ObjectId $azureAppWithDomain.ObjectId
                            }
                            else {
                                Remove-AzureADApplication -ObjectId $azureAppWithDomain.ObjectId
                            }
                        }
                        catch{
            
                        }
                    }
                    $Global:AzureDomain = $true
                }
            }

        }
        else {
            $msg = "INFO: No Azure Apps have been found with domain '$domain'."
            Write-Host $msg -ForegroundColor Red
            Log-Write -Message $msg   
            $Global:AzureDomain = $true
        }
    }
    catch {
        $msg = "ERROR: Failed to get the Azure Apps with domain '$domain'."
        Write-Host $msg -ForegroundColor Red
        Log-Write -Message $msg    
    }   
}

Function Remove-Office365Domain {
    param 
    (      
        [parameter(Mandatory=$false)] [Object]$TenantDomain,
        [parameter(Mandatory=$false)] [Object]$Domain,
        [parameter(Mandatory=$true)] [Object]$Credentials
    )

    if($AzureDomain -eq $true -and $MsolUserDomain -eq $true -and $MsolGroupDomain -eq $true -and $ExoDomain -eq $true) {
        do {
            $confirm = (Read-Host "ACTION: Do you want to remove domain '$domain' from Office 365 tenant?  [Y]es or [N]o" ) 
        } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

        if($confirm -eq "Y") {
            try {
                $domainToDelete = Get-MsolDomain -DomainName $domain
                if($domainToDelete.IsDefault) {

                    Set-MsolDomain -Name $newDefaultDomain -IsDefault

                    $msg = "INFO: '$domain' was the default domain. '$newDefaultDomain' is now the default one."
                    Write-Host $msg -ForegroundColor Green
                    Log-Write -Message $msg
                 }
            }
            catch {
                $msg = "ERROR: Failed to remove domain '$domain' from Office 365 tenant."
                Write-Host $msg -ForegroundColor Red
                Log-Write -Message $msg  
            }

            Remove-MsolDomain -DomainName $domain -force -ErrorAction SilentlyContinue

            if($error[0].toString()  -match 'Unable to remove this domain.*') {   
                $msolUsersWithDomain = @(Get-MsolUser -DomainName "thesociety.com.mx")     
                $msolUsersWithDomainCount = $msolUsersWithDomain.Count       
                
                $msolGroupsWithDomain = @(Get-MsolGroup -All | Where-Object {$_.EmailAddress -match $domain -or $_.ProxyAddresses -match $domain})
                $msolGroupsWithDomainCount = $msolGroupsWithDomain.Count

                if($msolUsersWithDomain -and !$msolGroupsWithDomain) {
                    $msg = "ERROR: Failed to remove domain '$domain' from Office 365 tenant. $msolUsersWithDomainCount MsolUsers ProxyAddresses are still blocking removal: '$($msolUsersWithDomain.DisplayName -join (";"))'"
                }    
                elseif($msolUsersWithDomain -and $msolGroupsWithDomain) {
                    $msg = "ERROR: Failed to remove domain '$domain' from Office 365 tenant. $msolUsersWithDomainCount MsolUsers ProxyAddresses and/or $msolGroupsWithDomainCount MsolGroups are still blocking removal: '$($msolGroupsWithDomain.DisplayName -join (";"))'"
                }
                elseif(!$msolUsersWithDomain -and $msolGroupsWithDomain) {
                    
                    $msg = "ERROR: Failed to remove domain '$domain' from Office 365 tenant. $msolGroupsWithDomainCount MsolGroups are still blocking removal."
                }
                else{
                    $msg = "ERROR: Failed to remove domain '$domain' from Office 365 tenant."  
                }
                Write-Host $msg -ForegroundColor Red
                Log-Write -Message $msg    
                
                $msg = "ACTION: Go to 'Microsoft 365 admin center->Setup->Domains' and you can now manually remove domain '$domain' from Office 365 tenant."
                Write-Host $msg -ForegroundColor Yellow
                Log-Write -Message $msg   
                Exit
            }
            else {
                $msg = "SUCCESS: Domain '$domain' removed from Office 365 tenant."
                Write-Host $msg -ForegroundColor Green
                Log-Write -Message $msg
                $domain = Select-Domain -Credentials $Credentials -DisplayAll $true
            }
        }
        else {
            Return
        }
    }
}

Function Add-Office365Domain {
    Write-Host
    do {
        $confirm = (Read-Host "ACTION: Do you want to add a new domain to the new Office 365 tenant?  [Y]es or [N]o" ) 
    } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

    if($confirm -eq "Y") {
        try {

            do {
                $domain = (Read-Host "ACTION: Enter the new domain you want to add to the new Office 365 tenant" ) 
            } while($domain.ToLower() -eq "")

            do {
                $confirm = (Read-Host "ACTION: Do you want to add a new domain '$domain' to the new Office 365 tenant?  [Y]es or [N]o" ) 
            } while(($confirm.ToLower() -ne "y") -and ($confirm.ToLower() -ne "n"))

            if($confirm -eq "Y") {

                #$Credentials = Connect-DestinationO365Tenant
                $Credentials = Connect-DestinationExchangeOnline

                $tenantId = (Get-MSOLCompanyInformation | Select-Object objectID).ObjectId.Guid
                New-MsolDomain  -TenantId $tenantId -Name $domain
                
                $msg = "SUCCESS: Domain '$domain' added to Office 365 tenant."
                Write-Host $msg -ForegroundColor Green
                Log-Write -Message $msg
                Write-Host
                
                $domainVerificationKey = Get-MsolDomainVerificationDNS -TenantId $tenantId  -DomainName $domain | Select-Object Label,Ttl
                $msg = "ACTION: Add a TXT record 'MS=$($domainVerificationKey.label.replace(".$domain",''))' with TTL '$($domainVerificationKey.Ttl)' in  your DNS  and wait for replication."
                Write-Host $msg -ForegroundColor Yellow
                Log-Write -Message $msg
            }
            else {
                Return
            }
        }
        catch {
            $msg = "ERROR: Failed to add domain '$domain' to Office 365 tenant."
            Write-Host -ForegroundColor Red  $msg
            Write-Host -ForegroundColor Red $_.Exception.Message
            Log-Write -Message $msg
            Log-Write -Message $_.Exception.Message
            Exit
        }

        do {
            $confirm = (Read-Host "ACTION: If you have added the TXT record to your DNS server then press [C] to continue" ) 
        } while($confirm -ne "C")

        Write-Host
        $msg = "INFO: Check domain validation $domain every 1 minute. Press [Ctrl] + [C] to cancel."
        Write-Host $msg 
        Log-Write -Message $msg
        
        do {
       
            $Status = Confirm-MsolDomain -TenantId $tenantId -DomainName $domain -ErrorAction SilentlyContinue      
            
            if($Status -eq $null) {
                $msg = "ERROR: Failed to confirm domain '$domain' in DNS."
                Write-Host -ForegroundColor Red  $msg
                Write-Host -ForegroundColor Red  $error[0].ToString()
                Log-Write -Message $msg
                Log-Write -Message  $error[0].ToString()

                $msg = "INFO: Check '$domain' domain validation every 1 minute. Press [Ctrl] + [C] to cancel."
                Write-Host $msg 
                Log-Write -Message $msg
            }
            
            Start-Sleep -Seconds 60                 
             
        }while($Status -eq $null)   

        if($Status){
            $msg = "SUCCESS: Domain $domain verified."
            Write-Host $msg -ForegroundColor Green
            Log-Write -Message $msg
            Write-Host
            Exit
        }  
    }
    else {
        Return
    }
}

#######################################################################################################################
#                                               MAIN PROGRAM
#######################################################################################################################

#Check PowerShell modules
Import-PowerShellModules

#Working Directory
$global:workingDir = "C:\scripts"

#Logs directory
$logDirName = "LOGS"
$logDir = "$workingDir\$logDirName"

#Log file
$logFileName = "$(Get-Date -Format yyyyMMdd)_Remove-DomainFromAllO365TenantObjects.log"
$logFile = "$logDir\$logFileName"

Create-Working-Directory -workingDir $workingDir -logDir $logDir

Write-Host
Write-Host "BitTitan Office 365 Domain migration tool"
Write-Host 
Write-Host -ForegroundColor Yellow "WARNING: Minimal output will appear on the screen." 
Write-Host -ForegroundColor Yellow "         Please look at the log file '$($logFile)'."
Write-Host -ForegroundColor Yellow "         Generated CSV file will be in folder '$($workingDir)'."
Write-Host 
Start-Sleep -Seconds 1

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT STARTED ++++++++++++++++++++++++++++++++++++++++"
Log-Write -Message $msg 


write-host 
$msg = "####################################################################################################`
                       CONNECTION TO SOURCE OFFICE 365 TENANT             `
####################################################################################################"
Write-Host $msg
Log-Write -Message $msg
Write-Host

$Credentials = Connect-O365Tenant

write-host 
$msg = "####################################################################################################`
                       SELECT DOMAIN TO BE REMOVED FROM SOURCE OFFICE 365 TENANT           `
####################################################################################################"
Write-Host $msg
Log-Write -Message $msg
Write-Host

$tenantDomain = Get-TenantDomain -Credentials $Credentials
$domain = Select-Domain -Credentials $Credentials

write-host 
$msg = "####################################################################################################`
                       REMOVE DOMAIN FROM MSOL USERS           `
####################################################################################################"
Write-Host $msg
Log-Write -Message $msg
Write-Host

Check-MsolUsersWithDomain -Domain $domain -tenantDomain $tenantDomain -Credentials $Credentials

write-host 
$msg = "####################################################################################################`
                       REMOVE DOMAIN FROM EXCHANGE ONLINE RECIPIENTS           `
####################################################################################################"
Write-Host $msg
Log-Write -Message $msg
Write-Host

Check-ExchangeOlineRecipientsWithDomain -Domain $domain -tenantDomain $tenantDomain -Credentials $Credentials

Check-MsolGroupsWithDomain -Domain $domain -tenantDomain $tenantDomain -Credentials $Credentials

#Check-AzureADDomain -Domain $domain -Credentials $Credentials

write-host 
$msg = "####################################################################################################`
                       TRY TO REMOVE DOMAIN SOURCE OFFICE 365 TENANT            `
####################################################################################################"
Write-Host $msg
Log-Write -Message $msg
Write-Host
Write-Host
Remove-Office365Domain -Domain $domain -tenantDomain $tenantDomain -Credentials $Credentials

Write-Host
$msg = "ACTION: Go to 'Microsoft 365 admin center->Setup->Domains' and you can now try to manually remove domain '$domain' from Office 365 tenant."
Write-Host $msg -ForegroundColor Yellow
Log-Write -Message $msg   
Write-Host

#END SCRIPT 
Write-Host

$msg = "++++++++++++++++++++++++++++++++++++++++ SCRIPT FINISHED ++++++++++++++++++++++++++++++++++++++++`n"
Log-Write -Message $msg

if($o365Session) {

    try {
        Write-Host "INFO: Opening directory $workingDir where you will find all the generated CSV files."
        Invoke-Item $workingDir
        Write-Host
    }
    catch{
        $msg = "ERROR: Failed to open directory '$workingDir'. Script will abort."
        Write-Host -ForegroundColor Red $msg
        Exit
    }

    Remove-PSSession $o365Session
}

Exit



