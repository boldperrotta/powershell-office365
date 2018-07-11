function Show-Menu
{
     param (
           [string]$Title = 'Office 365 Investigation Tool'
     )
     cls
     Write-Host "================ $Title ================"
     
     Write-Host "1: Press '1' to connect to Office365 and AzureAD Services."
     Write-Host "2: Press '2' to enable Auditing for all Mailboxes and set 6 month retention"
     Write-Host "3: Press '3' to gather information regarding the Mailbox and Tenant"
     Write-Host "4: Press '4' to reset the users password"
     Write-Host "5: Press '5' to run a message trace and geolocate sent messages"
     Write-Host "6: Press '6' to close all connections from compromised user to O365/Azure"
     Write-Host "A: Press 'A' to Install Azure AD Powershell Module (Required)"
     Write-Host "Q: Press 'Q' to quit."
}

do
{
     Show-Menu
     $input = Read-Host "Please make a selection"
     switch ($input)
     {
           '1' {
                cls
                'You chose option #1, please enter admin credentials for Office 365'
                #Get Credentials and Connect to Office 365 and Azure AD
                Import-Module AzureAD
                $UserCredential = Get-Credential
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
                Import-PSSession $Session
                Connect-MsolService -Credential $UserCredential
                Connect-AzureAD -Credential $UserCredential
           } '2' {
                cls
                'You chose option #2, auditing will be enabled'
                #Enable Mailbox auditing for the Organization
                Get-mailbox -Filter {(RecipientTypeDetails -eq 'UserMailbox')} | ForEach {Set-Mailbox $_.Identity -AuditEnabled $true -AuditAdmin Copy,Create,FolderBind,HardDelete,MessageBind,Move,MoveToDeletedItems,SendAs,SendOnBehalf,SoftDelete,Update -AuditDelegate Create,FolderBind,HardDelete,Move,MoveToDeletedItems,SendAs,SendOnBehalf,SoftDelete,Update -AuditOwner Create,HardDelete,MailboxLogin,Move,MoveToDeletedItems,SoftDelete,Update -AuditLogAgeLimit 180}

                #Enable Unified Audit Log for the Organization
                Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true
           } '3' {
                cls
                'You chose option #3, output is save to the folder CompromisedAccount which is created where you ran the script from'
                #Make directory to collect logs and Set Location
                New-Item -Path .\CompromisedAccount -ItemType directory | Out-Null
                Set-Location -Path .\CompromisedAccount | Out-Null

                #Collect email in question
                if (!$Email) { $Email = Read-Host 'Enter the email address of the compromised account' }

                #Collect Mailbox attributes
                Get-Mailbox -Identity $Email | fl > $email-MailboxAttributes.txt

                #Collect password set date for all users
                Get-MsolUser -All | select DisplayName, LastPasswordChangeTimeStamp > PasswordSetDate.txt
                
                #Get Inbox rules for Compromised email
                Get-InboxRule -Mailbox $Email | fl > $Email-InboxRules.txt

                #Collect Mail forwards for Compromised Email and remove
                Get-Mailbox -Identity $Email | Select Name, DeliverToMailboxAndForward, ForwardingSmtpAddress > $Email-MailForwards.txt
                #Set-Mailbox -Identity $Email -DeliverToMailboxAndForward $false -ForwardingSmtpAddress $null

                #Collect Mailbox delegates
                Get-MailboxPermission -Identity $Email > $Email-MailboxPermissions.txt

           } '4' {
                cls
                'You chose option #4, the users random password will be saved to Email-NewPassword.txt in the compromised account folder'
                #Reset Password to account
                if (!$Email) { $Email = Read-Host 'Enter the email address of the compromised account' }
                $ascii=$NULL;For ($a=33;$a –le 126;$a++) {$ascii+=,[char][byte]$a }
                Function Get-TempPassword() {
                Param(
                [int]$length=10,
                [string[]]$sourcedata
                )
                For ($loop=1; $loop –le $length; $loop++) {
                            $TempPassword+=($sourcedata | GET-RANDOM)
                         }
                return $TempPassword
                }
                $NewPassword = Get-TempPassword –length 12 –sourcedata $ascii
                $NewPassword > $Email-NewPassword.txt
                $NewPassword
                Set-MsolUserPassword -UserPrincipalName "$Email" -NewPassword $NewPassword -ForceChangePassword $false
           } '5' {
                cls
                'You chose option #5, the message trace will be exported into the compromised account folder'
                #Run Message Trace on account, GeoLocate Sent Emails
                if (!$Email) { $Email = Read-Host 'Enter the email address of the compromised account' }
                $EndDate = Get-Date
                $StartDate = $EndDate.AddDays(-30)
                $IPList = Get-MessageTrace -SenderAddress $Email -StartDate $StartDate -EndDate $EndDate | select FromIP,Received
                $TableColumn = "Time,IP Address,CountryCode,Country,RegionCode,Region,City,Zip Code,TimeZone,Latitude,Longitude,MetroCode"
                $TableColumn | Out-File MessageTraceIP.txt -Append
                Foreach ($IPAddress in $IPList) {
                    $Received = $IPAddress.Received
                    $API = "5d90da1fdb8655af8b21ca307aaaf339"
                    $URI = "api.ipstack.com/$($IPAddress.FromIP)?access_key=$API&fields=ip,continent_name,country_name,region_name,city"
                    $Geo = Invoke-RestMethod $URI 
                    "$Received, $Geo" | Out-File $Email-MessageTraceIP.txt -Append
    
                    }
                #Import-CSV $Email-MessageTraceIP.txt -Delimiter "," | Export-Csv $Email-MessageTraceIP.csv
                Get-MessageTrace -SenderAddress $Email -StartDate $StartDate -EndDate $EndDate | fl > $Email-Sent.txt
                Get-MessageTrace -RecipientAddress $Email -StartDate $StartDate -EndDate $EndDate | fl > $Email-Received.txt
           } '6' {
                cls
                'You chose option #6, all connections to azure for the compromised user will be terminated'
                if (!$Email) { $Email = Read-Host 'Enter the email address of the compromised account' }
                Get-AzureADUser -SearchString $Email | Revoke-AzureADUserAllRefreshToken
           } 'a' {
                cls
                'You selected option A, please follow and agree to all prompts  to install the AzureAD Module'
                Install-Module AzureAD
           } 'q' {
                cls
                Remove-PSSession $Session
                return
           }
     }
     pause
}
until ($input -eq 'q')