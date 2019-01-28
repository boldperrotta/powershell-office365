function Show-Menu
{
     param (
           [string]$Title = 'Office 365 Best Practices Script'
     )
     cls
     Write-Host "================ $Title ================"
     
     Write-Host "1: Press '1' to connect to Office365 and AzureAD Services."
     Write-Host "2: Press '2' to enable Auditing for all Admin Activity, Mailboxes and set 6 month retention"
     Write-Host "3: Press '3' to block countries with high reputation of spam and all of Africa and South America"
     Write-Host "4: Press '4' to Set advanced spam options"
     Write-Host "5: Press '5' to Disable IMAP/POP for all existing accounts and any new accounts"
     Write-Host "6: Press '6' to Block Mail Rules that Autoforward Mail"
     Write-Host "6: Press '7' to check if external email has matching display name with internal user"
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
                'You chose option #3, Email from outside countries with high reputation of spam will be sent to junk, plus all email from Africa and South America'
                #Set countries to send to spam to default spam policy
                Set-HostedContentFilterPolicy -Identity Default -EnableRegionBlockList $true -RegionBlockList CN,RU,UA,JP,IN,HK,GB,BR,DE,AO,AR,BF,BI,BJ,BO,BR,BW,CD,CF,CG,CI,CL,CM,CO,CV,DJ,DZ,EC,EG,EH,ER,ET,FK,GA,GF,GH,GM,GN,GQ,GW,GY,KE,KM,LR,LS,LY,MA,MG,ML,MR,MU,MW,MZ,NA,NE,NG,PE,PY,RE,RW,SC,SD,SH,SL,SN,SO,SR,ST,SZ,TD,TG,TN,TZ,UG,UY,VE,YT,ZA,ZM,ZW
           } '4' {
                cls
                'You chose option #4, Advanced Spam Options will be set for default mail policy'
                #Set advanced spam options
                Set-HostedContentFilterPolicy -Identity Default -MarkAsSpamSpfRecordHardFail On -IncreaseScoreWithNumericIps On -IncreaseScoreWithRedirectToOtherPort On -MarkAsSpamJavaScriptInHtml On -MarkAsSpamNdrBackscatter On
           } '5' {
                cls
                'You chose option #4, Disable IMAP and POP for all account and new accounts going forward'
                #Set advanced spam options
                $Mailboxes = Get-Mailbox -ResultSize Unlimited
                ForEach ($Mailbox in $Mailboxes) {$Mailbox | Set-CASMailbox -PopEnabled $False -ImapEnabled $False }
           } '6' {
               #Turn off autoforwarding for the Domain
               Set-RemoteDomain Default -AutoForwardEnabled $false
           } '7' {
               #Create transport rule to check if Display Name may be spoofed or update transport list with new users
               $ruleName = "External Senders with matching Display Names"
               $ruleHtml = "<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 align=left width=`"100%`" style='width:100.0%;mso-cellspacing:0cm;mso-yfti-tbllook:1184; mso-table-lspace:2.25pt;mso-table-rspace:2.25pt;mso-table-anchor-vertical:paragraph;mso-table-anchor-horizontal:column;mso-table-left:left;mso-padding-alt:0cm 0cm 0cm 0cm'>  <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;mso-yfti-lastrow:yes'><td style='background:#910A19;padding:5.25pt 1.5pt 5.25pt 1.5pt'></td><td width=`"100%`" style='width:100.0%;background:#FDF2F4;padding:5.25pt 3.75pt 5.25pt 11.25pt; word-wrap:break-word' cellpadding=`"7px 5px 7px 15px`" color=`"#212121`"><div><p class=MsoNormal style='mso-element:frame;mso-element-frame-hspace:2.25pt; mso-element-wrap:around;mso-element-anchor-vertical:paragraph;mso-element-anchor-horizontal: column;mso-height-rule:exactly'><span style='font-size:9.0pt;font-family: `"Segoe UI`",sans-serif;mso-fareast-font-family:`"Times New Roman`";color:#212121'>This message was sent from outside the company by someone with a display name matching a user in your organization. Please do not click links or open attachments unless you recognize the source of this email and know the content is safe. <o:p></o:p></span></p></div></td></tr></table>"
 
               $rule = Get-TransportRule | Where-Object {$_.Identity -contains $ruleName}
               $displayNames = (Get-Mailbox -ResultSize Unlimited).DisplayName
 
               if (!$rule) {
                    Write-Host "Rule not found, creating rule" -ForegroundColor Green
                    New-TransportRule -Name $ruleName -Priority 0 -FromScope "NotInOrganization" -SentTo "Chris Toppin" -ApplyHtmlDisclaimerLocation "Prepend" `
                         -HeaderMatchesMessageHeader From -HeaderMatchesPatterns $displayNames -ApplyHtmlDisclaimerText $ruleHtml
               }
               else {
                    Write-Host "Rule found, updating rule" -ForegroundColor Green
                    Set-TransportRule -Identity $ruleName -Priority 0 -FromScope "NotInOrganization" -ApplyHtmlDisclaimerLocation "Prepend" `
                         -HeaderMatchesMessageHeader From -HeaderMatchesPatterns $displayNames -ApplyHtmlDisclaimerText $ruleHtml
               }
           } '8' {
               $ruleName = "Block Uncommon Attachment Extensions"
               $rule = Get-TransportRule | Where-Object {$_.Identity -contains $ruleName}
 
               if (!$rule) {
                   Write-Host "Rule not found, creating rule" -ForegroundColor Green
                   New-TransportRule -Name $ruleName -Priority 0 -FromScope "NotInOrganization" -SentTo "Chris Toppin" -AttachmentExtensionMatchesWords .adp,.app,.asp,.bas,.bat,.cer,.chm,.cmd,.cnt,.com,.cpl,.crt,.csh,.der,.exe,.fxp,.gadget,.hlp,.hpj,.hta,.inf,.ins,.isp,.its,.js,.jse,.ksh,.lnk,.mad,.maf,.mag,.mam,.maq,.mar,.mas,.mat,.mau,.mav,.maw,.mda,.mdb,.mde,.mdt,.mdw,.mdz,.msc,.msh,.msh1,.msh2,.mshxml,.msh1xml,.msh2xml,.msi,.msp,.mst,.ops,.osd,.pcd,.pif,.plg,.prf,.prg,.pst,.reg,.scf,.scr,.sct,.shb,.shs,.ps1,.ps1xml,.ps2,.ps2xml,.psc1,.psc2,.tmp,.url,.vb,.vbe,.vbp,.vbs,.vsmacros,.vsw,.ws,.wsc,.wsf,.wsh,.xnk,.ade,.cla,.class,.grp,.jar,.mcf,.ocx,.pl,.xbap -DeleteMessage $true
                           }
               else {
                   Write-Host "Rule found, updating rule" -ForegroundColor Green
                   Set-TransportRule -Identity $ruleName -Priority 0 -FromScope "NotInOrganization" -AttachmentExtensionMatchesWords .adp,.app,.asp,.bas,.bat,.cer,.chm,.cmd,.cnt,.com,.cpl,.crt,.csh,.der,.exe,.fxp,.gadget,.hlp,.hpj,.hta,.inf,.ins,.isp,.its,.js,.jse,.ksh,.lnk,.mad,.maf,.mag,.mam,.maq,.mar,.mas,.mat,.mau,.mav,.maw,.mda,.mdb,.mde,.mdt,.mdw,.mdz,.msc,.msh,.msh1,.msh2,.mshxml,.msh1xml,.msh2xml,.msi,.msp,.mst,.ops,.osd,.pcd,.pif,.plg,.prf,.prg,.pst,.reg,.scf,.scr,.sct,.shb,.shs,.ps1,.ps1xml,.ps2,.ps2xml,.psc1,.psc2,.tmp,.url,.vb,.vbe,.vbp,.vbs,.vsmacros,.vsw,.ws,.wsc,.wsf,.wsh,.xnk,.ade,.cla,.class,.grp,.jar,.mcf,.ocx,.pl,.xbap -DeleteMessage $true
                    }
           } 'q' {
                cls
                Remove-PSSession $Session
                return
           }
     }
     pause
}
until ($input -eq 'q')
