# This script is used for search a specific SMTP address in all Distribution Groups inside your Exchange Online. 
# Before run this code you will need connect to ExO. 
# https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell?view=exchange-ps 


$SMTP = Read-Host "Who are you looking for? ";
Write-Host "..........................................";  
Write-Host "Getting all distribution groups..."; 
$AllDG = Get-DistributionGroup | Where-Object {$_.GroupType -eq "Universal"} 
ForEach ($DG in $AllDG)
{
	$Members = Get-DistributionGroupMember -Identity $DG.PrimarySmtpAddress | Where-Object{$_.RecipientType -eq "UserMailbox"}; 
	ForEach ($Member in $Members)
	{ 
		if ($Member.PrimarySmtpAddress -eq $SMTP){
			Write-Host "Distribution list: " $DG.PrimarySmtpAddress; 
			Write-Host "Member: " $Member.PrimarySmtpAddress; 
			Write-Host "................................";
		}
	}
}
