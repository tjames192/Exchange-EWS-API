# get um enabled mailboxes
$umMailboxes  = @(Get-Mailbox -ResultSize Unlimited -Filter {UMEnabled -eq $true})

# Path to the Exchange Web Services DLL
$EWSManagedApiPath = "c:\Program Files\Microsoft\Exchange Server\V15\Bin\Microsoft.Exchange.WebServices.dll"

# Load EWS Managed API
Add-Type -Path $EWSManagedApiPath

# Create EWS Object
$Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013)

# csv object
[System.Collections.ArrayList]$csv = @()

# loop through each user
foreach ( $User in $umMailboxes ) {
	  # UM mailbox
    $Mailbox = $User

    # Get the SMTP address from the mailbox
    $SmtpAddress = $Mailbox.PrimarySmtpAddress.ToString()
	
    # Get the autodiscover URL
    $Service.AutodiscoverUrl($SmtpAddress)

    # Use EWS Impersonation to locate the root folder in the mailbox
    $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $SmtpAddress)
    $FolderId = [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root
    $Folder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$FolderId)

    # Create the search filter to find the UM custom greetingsd
    $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass,"IPM.Configuration.Um.CustomGreetings",[Microsoft.Exchange.WebServices.Data.ContainmentMode]::Substring,[Microsoft.Exchange.WebServices.Data.ComparisonMode]::IgnoreCase);

    # Define the EWS search view
    $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(100, 0, [Microsoft.Exchange.WebServices.Data.OffsetBasePoint]::Beginning)
    $view.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
    $view.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
	
	  $props = [PSCustomObject][ordered]@{
			'Name' = $Mailbox.name
			'userprincipalname' = $Mailbox.userprincipalname
			'AwayGreeting' = if((Get-UMMailboxConfiguration $Mailbox.samaccountname).greeting -eq 'Away') { $true } else { $false }
			'standardGreeting' = ''
			'oofGreeting' = ''
		}
	
    # Do the search and enumerate the results
    $results = $service.FindItems( $FolderId, $searchFilter, $view )
    if ( $results -ne $null ) {
        # Define the property set required to get the binary audio data
        $psPropset = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)

        # Add the binary audio data property to the property set
        $PidTagRoamingBinary = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x7C09,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);  
        $psPropset.Add($PidTagRoamingBinary)  

        # Load the new properties
        [Void]$service.LoadPropertiesForItems($results,$psPropset)
        
        # Create a folder to export custom greeting
        $Path = "c:\MailboxCustomGreetings\" + $Mailbox.name
        $Path = new-item $Path -itemtype directory
    }
    # Loop through the results
		foreach ( $Item in $results.Items ) {
			# If this is the main greeting, set the flag to true
			if ( $Item.ItemClass -eq "IPM.Configuration.Um.CustomGreetings.External" ) {
				$Filename = $Mailbox.samaccountname + "_Standard.wav"
            
				$props.standardGreeting = "$Path\$Filename"
			}
        
			# If this is the extended away greeting, set the flag to true
			if ( $Item.ItemClass -eq "IPM.Configuration.Um.CustomGreetings.Oof" ) {
				$Filename = $Mailbox.samaccountname + "_Oof.wav"
            
				$props.oofGreeting = "$Path\$Filename"
			}
        
			[IO.File]::WriteAllBytes("$Path\$Filename",$Item.ExtendedProperties.Value)
		}	
    $csv.Add($props) | Out-Null
}
# end

$csv | Export-Csv -NoClobber -NoTypeInformation "C:\MailboxCustomGreetings\ex.um.customgreetings.csv"
