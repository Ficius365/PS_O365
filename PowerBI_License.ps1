# This script will create a report with all users with PowerBi service assigned (enabled or not). 
# Report include DisplayName, AssignedTimestamp, CapabilityStatus and Service (will be always PowerBi)
# You will be able to select export path changing variable called: $reportPath
# You will need install and connect (with GA credential) to AzureAD module: Install-Module AzureAD and Connect-AzureAD 

$usersWithPowerBI = Get-AzureADUser -All $true | Where-Object {$_.AssignedPlans.service -eq "PowerBI"}
$usersWithPowerBITable = @()
$reportPath = "C:\test\powerbiusers.csv"

$usersWithPowerBI | ForEach-Object {
	
	if ($_.AssignedPlans.service -eq "PowerBI"){
		$userWithPowerBIRow = New-Object PSObject
		Add-Member -InputObject $userWithPowerBIRow -MemberType NoteProperty -Name DisplayName -Value $_.DisplayName; 
		$service = $_.AssignedPlans | Where-Object {$_.Service -eq "PowerBI"}
		Add-Member -InputObject $userWithPowerBIRow -MemberType NoteProperty -Name AssignedTimestamp -Value $service.AssignedTimestamp; 
		Add-Member -InputObject $userWithPowerBIRow -MemberType NoteProperty -Name CapabilityStatus -Value $service.CapabilityStatus; 
		Add-Member -InputObject $userWithPowerBIRow -MemberType NoteProperty -Name Service -Value $service.Service; 
		$usersWithPowerBITable += $userWithPowerBIRow; 
	}

}

$usersWithPowerBITable | Export-Csv -Path $reportPath -NoTypeInformation
