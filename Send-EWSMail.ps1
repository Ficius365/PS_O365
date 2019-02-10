# 10/02/2019 EWS Managed API 2.2, which can be downloaded here:
# https://www.microsoft.com/en-gb/download/details.aspx?id=42951 

# Make sure the Import-Module command matches the Microsoft.Exchange.WebServices.dll location of EWS Managed API, chosen during the installation 
Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"; 

# Create new service 
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList Exchange2013_SP1; 

# Provide the credentials of the O365 from which you want to send the email message
$credentials = Get-Credential; 
$service.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $credentials; 

# Use autodiscover to reach ExO server
$service.AutodiscoverUrl($credentials.UserName, {$true}); 
# Or manually 
# $service.Url= new-object Uri("https://outlook.office365.com/EWS/Exchange.asmx"); 

# New instance EmailMessage class
$message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $service
$message.Subject = 'Testing'
$message.Body = 'This message is being sent through EWS with PowerShell'
$message.ToRecipients.Add("user@dominio.com")
# $message.CcRecipients.Add() and $message.BccRecipients.Add()
# Send saving a copy 
$message.SendAndSaveCopy()
# Send without save a copy 
# $message.Send()
