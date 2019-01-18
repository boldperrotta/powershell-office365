function Show-Menu
{
     param (
           [string]$Title = 'Office 365 Best Practices Script'
     )
     cls
     Write-Host "================ $Title ================"
     
     Write-Host "1: Press '1' to connect to Office365 and AzureAD Services."
     Write-Host "2: Press '2' to enable Auditing for all Admin Activity, Mailboxes and set 6 month retention"
     Write-Host "3: Press '3' to block countries with high reputation of spam"
     Write-Host "4: Press '4' to Set advanced spam options"
     Write-Host "5: Press '5' to Disable IMAP/POP for all existing accounts and any new accounts"
     Write-Host "5: Press '6' to Block Mail Rules that Autoforward Mail"
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
                'You chose option #2, mailbox and admin auditing will be enabled'
                #Enable Mailbox Auditing and set 180 day retention
                Get-mailbox -Filter {(RecipientTypeDetails -eq 'UserMailbox')} | ForEach {Set-Mailbox $_.Identity -AuditEnabled $true -AuditLogAgeLimit 180 -AuditAdmin Copy,Create,FolderBind,HardDelete,MailItemAccessed,Move,MoveToDeletedItems,SendAs,SendOnBehalf,SoftDelete,Update -AuditDelegate Create,FolderBind,HardDelete,Move,MoveToDeletedItems,SendAs,SendOnBehalf,SoftDelete,Update -AuditOwner Create,HardDelete,MailboxLogin,Move,MoveToDeletedItems,SoftDelete,Update} 
                
                #Enable Unified Audit Logging
                Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true 
           } '3' {
                cls
                'You chose option #3, Email from outside countries with high reputation of spam will be sent to junk'
                #Set countries to send to spam to default spam policy
                Set-HostedContentFilterPolicy -Identity Default -EnableRegionBlockList $true -RegionBlockList CN,RU,UA,JP,IN,HK,GB,BR,DE
           } '4' {
                cls
                'You chose option #4, Advanced Spam Options will be set for default mail policy'
                #Set advanced spam options
                Set-HostedContentFilterPolicy -Identity Default -IncreaseScoreWithNumericIps On -IncreaseScoreWithRedirectToOtherPort On -MarkAsSpamJavaScriptInHtml On -MarkAsSpamNdrBackscatter On
           } '5' {
                cls
                'You chose option #4, Disable IMAP and POP for all account and new accounts going forward'
                #Set advanced spam options
                $Mailboxes = Get-Mailbox -ResultSize Unlimited
                ForEach ($Mailbox in $Mailboxes) {$Mailbox | Set-CASMailbox -PopEnabled $False -ImapEnabled $False }
           } '6' {
               #Turn off autoforwarding for the Domain
               Set-RemoteDomain Default -AutoForwardEnabled $false
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
