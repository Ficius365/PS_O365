# This script list a mailbox FolderID list for all mailbox folders and subfolders 
# This FolderIds will be used for cmdlet New-ComplianceSearch in flag -ContentMatchQuery FolderID:xxxxxxxxxxxxxxx
# https://docs.microsoft.com/en-us/powershell/module/exchange/policy-and-compliance-content-search/new-compliancesearch?view=exchange-ps 

# Use Content Search in Office 365 for targeted collections: 
# https://docs.microsoft.com/en-us/office365/securitycompliance/use-content-search-for-targeted-collections 

# Before run this code you will need connect to ExO. 
# https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/connect-to-exchange-online-powershell?view=exchange-ps 

#Ask for specific SMTP address 
$SMTP = Read-Host "Write the mailbox to list the folderids "; 

$folderQueries = @()
$folderStatistics = Get-MailboxFolderStatistics $SMTP
foreach ($folderStatistic in $folderStatistics)
{
    $folderId = $folderStatistic.FolderId;
    $folderPath = $folderStatistic.FolderPath;
    $encoding= [System.Text.Encoding]::GetEncoding("us-ascii")
    $nibbler= $encoding.GetBytes("0123456789ABCDEF");
    $folderIdBytes = [Convert]::FromBase64String($folderId);
    $indexIdBytes = New-Object byte[] 48;
    $indexIdIdx=0;
    $folderIdBytes | Select-Object -skip 23 -First 24 | ForEach-Object{$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -shr 4];$indexIdBytes[$indexIdIdx++]=$nibbler[$_ -band 0xF]}
    $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))";
    $folderStat = New-Object PSObject
    Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderPath -Value $folderPath
    Add-Member -InputObject $folderStat -MemberType NoteProperty -Name FolderQuery -Value $folderQuery
    $folderQueries += $folderStat 
    #$folderQueries += $folderId 
}
Write-Host "-----Exchange Folders-----"
$folderQueries | Format-Table 