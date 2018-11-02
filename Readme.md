# Office365 Threat Hunting Scripts
## Powershell to identify compromised/vulnerable mailboxes and accounts

### Prerequisites
* **All** - [Import-Excel](https://github.com/dfinke/ImportExcel) module - Install with `Install-Module ImportExcel` - You can change to export-csv if you don't have Excel installed.

* **find-rules.ps1** - You need to connect to Exchange Online, if you can connect without multifactor authentication then you can just create a new remote powershell session
```powershell
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential (Get-Credential) -Authentication Basic -AllowRedirection
```
```powershell
Import-PSSession $Session -DisableNameChecking
```

However if you do use multifactor authentication then you'll need to [install the Exchange Online Remote PowerShell Module](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps) and then call `Connect-Exopssession" before the script.

* **audit-mfa.ps1** - Requires the [AzureAd](https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-office-365-powershell) module which can be installed with the following command `Install-Module -Name AzureAD` and then connected with the `Connect-AzureAD` cmdlet before running the script. This natively supports multifactor login.

* **audit-mailboxes.ps1** - Requires both Exchange online access and the Azure AD module so you'll need to complete both of the above steps.


More info available here: https://blog.rothe.uk/office365-detecting-compromise-with-powershell/

