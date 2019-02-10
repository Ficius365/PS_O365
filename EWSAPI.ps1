# 10/02/2019 EWS Managed API 2.2, which can be downloaded here:
# https://www.microsoft.com/en-gb/download/details.aspx?id=42951 

#We will need to create a new RBAC group with ‘ApplicationImpersonation’ role assigned and add as member the service account that will connect to the mailboxes. 

# Make sure the Import-Module command matches the Microsoft.Exchange.WebServices.dll location of EWS Managed API, chosen during the installation 
Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"; 

# Create new service 
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList Exchange2013_SP1; 

# Provide the credentials of the O365 account that has impersonation rights on the mailboxes
$credentials = Get-Credential; 
$service.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $credentials; 

# Provide target mailbox SMTP 
$targetMailbox = "america.vanegas@365lab.es"; 

# Exchange Online URL
# $service.Url= new-object Uri("https://outlook.office365.com/EWS/Exchange.asmx"); 
# Or use autodiscover service
$service.AutodiscoverUrl($targetMailbox, {$True})

# Provide user that will be impersonated 
$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$targetMailbox); 
$service.ImpersonatedUserId; 
