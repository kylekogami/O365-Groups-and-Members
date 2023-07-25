#Install-Module ExchangePowerShell

Import-Module PSExcel
Import-Module AzureADPreview
Import-Module ExchangePowerShell
mkdir C:\scripts

#Connect to Exchange Online
Connect-ExchangeOnline -ShowBanner:$False


Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | select Identity,PrimarySmtpAddress,User | Where-Object {($_.user -like '*@*')}|Export-Csv C:\scripts\SharedMailboxes.csv  -NoTypeInformation 