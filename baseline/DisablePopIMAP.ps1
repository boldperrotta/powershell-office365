'You chose option #4, Disable IMAP and POP for all account and new accounts going forward'
#Set advanced spam options
$Mailboxes = Get-Mailbox -ResultSize Unlimited
ForEach ($Mailbox in $Mailboxes) {$Mailbox | Set-CASMailbox -PopEnabled $False -ImapEnabled $False }
