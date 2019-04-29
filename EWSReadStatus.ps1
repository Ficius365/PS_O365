#Parameters 

#We will get the ReadStatus for following message: 
#$MessageID = "<31292b408e444bd4805768a7eacc9c66@AM4PR0101MB2306.eurprd01.prod.exchangelabs.com>"; 
$MessageID = Read-Host "Please, provide us the MessageID: "; 

#Mailbox where item is stored: 
#$targetMailbox = "america.vanegas@365lab.es"; 
$targetMailbox = Read-Host "Please, provide us the targetMailbox: "; 

#-----------------------------------------------#

# Make sure the Import-Module command matches the Microsoft.Exchange.WebServices.dll location of EWS Managed API, chosen during the installation 
Import-Module -Name "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"; 

# Create new service 
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList Exchange2013_SP1; 

# Provide the credentials of the O365 account that has impersonation rights on the mailboxes
$credentials = Get-Credential; 
$service.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $credentials; 

# Find the EWS Url
$service.AutodiscoverUrl($targetMailbox, {$True}); 

# Provide user that will be impersonated 
$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$targetMailbox); 

#---------------------------------------------------#

#Get the properties of the Items
$PropSet =  new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::IsRead)
$PropSet.add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
$PropSet.add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
$PropSet.add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender)
$PropSet.add([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageId)
    
$IsEverReadProp = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0xE07,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
$PropSet.Add($IsEverReadProp)
    
#Setup the View
$itemView = new-object Microsoft.Exchange.WebServices.Data.ItemView(1000,0,[Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
$itemView.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Shallow
$itemView.PropertySet = $PropSet

#Sort objects for quick hit
$itemView.OrderBy.add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,[Microsoft.Exchange.WebServices.Data.SortDirection]::Descending)

#Search filters MessageID
$searchFilterEA1 = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageId,$MessageID) 

$oSearchFilters = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
$oSearchFilters.add($searchFilterEA1)

#Try the Inbox first
$oFindItems = $service.FindItems([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$oSearchFilters,$itemView)

If(($oFindItems.Items.Count -gt 1)){
    Write-Verbose "Duplicate Items from the mailbox, its not expected"
    Throw "Duplicate Items from Mailbox"
}elseif($oFindItems.Items.Count -lt 1){
    Write-Verbose "Item is not Present in Inbox, creating search folder to find the message from all folders"
    #Item not in inbox, search the MsgFolderRoot and all subfolders
    #Create a Search folder

    $svFldid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$targetMailbox)
    $SearchFolder = new-object Microsoft.Exchange.WebServices.Data.SearchFolder($service)
    $searchFolder.SearchParameters.RootFolderIds.Add([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot) | Out-Null
    $searchFolder.SearchParameters.Traversal = [Microsoft.Exchange.WebServices.Data.SearchFolderTraversal]::Deep
    $searchFolder.SearchParameters.SearchFilter = $oSearchFilters
    $searchFolder.DisplayName = "TEMP-MSG-ID"
    $searchFolder.Save($svFldid) | Out-Null
    Write-Verbose "Created SearchFolder TEMP-MSG-ID on the Root of mailbox"
    $oFindItems = $SearchFolder.FindItems($oSearchFilters,$itemView)
    
    If(($oFindItems.Items.Count -gt 1)){
        Write-Verbose "Duplicate Items from the mailbox, its not expected"
        Throw "Duplicate Items from Mailbox" 
    }
    elseif($oFindItems.Items.Count -lt 1){
        Write-Verbose "Item not present in Mailbox: $targetMailbox"
        Throw "Item not present in Mailbox: $targetMailbox"
    }

}else{

    Write-Verbose "Item found from Inbox Folder"
}

$Item = $oFindItems.Items[0]
#Read the extended property of the Item to find EverRead value
$IsEverReadVal = ($item.ExtendedProperties | Select-Object -Property value).value

If(($IsEverReadVal -band 0x0400) -eq 0x0400){
    
    #Write-Host "User had already READ the message"
    $IsEverRead = $true

}else{
    $IsEverRead = $false    
}

$props = @{ DateTimeReceived  = $Item.DateTimeReceived
    Sender            = $Item.Sender.Name;
    Subject           = $Item.Subject;
    Mailbox			  = $MailboxName
    IsRead            = $item.IsRead
    IsEverRead        = $IsEverRead
    Error             = $NULL
}

$eitem = New-Object -TypeName PSCustomObject -Property $props
$service.ImpersonatedUserId = $null

return $eitem

