<#
	.SYNOPSIS
		This cmdlet is available in on-premises Exchange Server 2016 and in the cloud-based service. Some parameters and settings may be exclusive to one environment or the other.

	.DESCRIPTION
		When you use the Get-Mailbox cmdlet in on-premises Exchange environments to view the quota settings for a mailbox, you first need to check the value of the UseDatabaseQuotaDefaults property. The value True means per-mailbox quota settings are ignored, and you need to use the Get-MailboxDatabase cmdlet to see the actual values. If the UseDatabaseQuotaDefaults property is False, the per-mailbox quota settings are used, so what you see with the Get-Mailbox cmdlet are the actual quota values for the mailbox.

		
		You need to be assigned permissions before you can run this cmdlet. Although all parameters for this cmdlet are listed in this topic, you may not have access to some parameters if they're not included in the permissions assigned to you. To see what permissions you need, see the "Recipient Provisioning Permissions" section in the Recipients Permissions topic.


	.PARAMETER Anr String
		The Anr parameter specifies a string on which to perform an ambiguous name resolution (ANR) search. You can specify a partial string and search for objects with an attribute that matches that string. The default attributes searched are:
		CommonName (CN)
		DisplayName
		FirstName
		LastName
		Alias

	.PARAMETER Arbitration SwitchParameter
		This parameter is available only in on-premises Exchange 2016.

		The Arbitration parameter specifies that the mailbox for which you are executing the command is an arbitration mailbox. Arbitration mailboxes are used for managing approval workflow. For example, an arbitration mailbox is used for handling moderated recipients and distribution group membership approval.

	.PARAMETER Archive SwitchParameter
		The Archive switch filters the results by archive mailboxes. When you use this switch, only archive mailboxes are included in the results. This switch is required to return archive mailboxes. You don't need to specify a value with this switch.

	.PARAMETER AuditLog SwitchParameter
		This parameter is reserved for internal Microsoft use.

	.PARAMETER AuxAuditLog SwitchParameter
		This parameter is reserved for internal Microsoft use.

	.PARAMETER Credential PSCredential
		This parameter is available only in on-premises Exchange 2016.

		The Credential parameter specifies the user name and password that's used to run this command. Typically, you use this parameter in scripts or when you need to provide different credentials that have the required permissions.

		This parameter requires the creation and passing of a credential object. This credential object is created by using the Get-Credential cmdlet. For more information, see Get-Credential (http://go.microsoft.com/fwlink/p/?linkId=142122).

	.PARAMETER Database DatabaseIdParameter
		This parameter is available only in on-premises Exchange 2016.

		The Database parameter filters the results by mailbox database. When you use this parameter, only mailboxes on the specified database are included in the results. You can any value that uniquely identifies the database. For example:

		Name
		Distinguished name (DN)
		GUID
		You can't use this parameter with the Anr, Identity, or Server parameters.

	.PARAMETER DomainController Fqdn
		This parameter is available only in on-premises Exchange 2016.

		The DomainController parameter specifies the domain controller that's used by this cmdlet to read data from or write data to Active Directory. You identify the domain controller by its fully qualified domain name (FQDN). For example, dc01.contoso.com.

	.PARAMETER Filter String
		The Filter parameter indicates the OPath filter used to filter recipients.

		For more information about the filterable properties, see Filterable properties for the -Filter parameter.

	.PARAMETER GroupMailbox SwitchParameter
		The GroupMailbox switch filters the results by Office 365 groups. When you use this switch, Office 365 groups are included in the results. This switch is required to return Office 365 groups. You don't need to specify a value with this switch.

	.PARAMETER Identity MailboxIdParameter
		The Identity parameter specifies the mailbox that you want to view. You can use any value that uniquely identifies the mailbox.

		For example:

		Name
		Display name
		Alias
		Distinguished name (DN)
		Canonical DN
		<domain name>\<account name>
		Email address
		GUID
		LegacyExchangeDN
		SamAccountName
		User ID or user principal name (UPN)
		You can't use this parameter with the Anr, Database, MailboxPlan or Server parameters.

	.PARAMETER IgnoreDefaultScope SwitchParameter
		This parameter is available only in on-premises Exchange 2016.

		The IgnoreDefaultScope switch tells the command to ignore the default recipient scope setting for the Exchange Management Shell session, and to use the entire forest as the scope. This allows the command to access Active Directory objects that aren't currently available in the default scope.

		Using the IgnoreDefaultScope switch introduces the following restrictions:
		You can't use the DomainController parameter. The command uses an appropriate global catalog server automatically.
		You can only use the DN for the Identity parameter. Other forms of identification, such as alias or GUID, aren't accepted.

	.PARAMETER InactiveMailboxOnly SwitchParameter
		This parameter is available only in the cloud-based service.

		The InactiveMailboxOnly switch filters the results by inactive mailboxes. When you use this switch, only inactive mailboxes are included in the results. You don't need to specify a value with this switch.

		An inactive mailbox is a mailbox that's placed on Litigation Hold or In-Place Hold before it's soft-deleted. The contents of an inactive mailbox are preserved until the hold is removed.

	.PARAMETER IncludeInactiveMailbox SwitchParameter
		This parameter is available only in the cloud-based service.

		The IncludeInactiveMailboxswitch specifies that the command returns both active and inactive mailboxes. You don't need to specify a value with this switch.

		An inactive mailbox is a mailbox that's placed on Litigation Hold or In-Place Hold before it's soft-deleted. The contents of an inactive mailbox are preserved until the hold is removed.

	.PARAMETER MailboxPlan MailboxPlanIdParameter
		This parameter is available only in the cloud-based service.

		The MailboxPlan parameter filters the results by mailbox plan. When you use this parameter, only mailboxes that are assigned the specified mailbox plan are returned in the results. You can use any value that uniquely identifies the mailbox plan. For example:

		Name
		Alias
		Display name
		Distinguished name (DN)
		GUID
		A mailbox plan specifies the permissions and features available to a mailbox user in cloud-based organizations. You can see the available mailbox plans by using the Get-MailboxPlan cmdlet.

		You can't use this parameter with the Anr or Identity parameters.

	.PARAMETER Monitoring SwitchParameter
		This parameter is available only in on-premises Exchange 2016.

		The Monitoringswitchfilters the results by monitoring mailboxes. When you use this switch, only monitoring mailboxes are included in the results. This switch is required to return monitoring mailboxes. You don't need to specify a value with this switch.

		Monitoring mailboxes are associated with managed availability and the Exchange Health Manager service, and have a RecipientTypeDetails property value of MonitoringMailbox.

	.PARAMETER OrganizationalUnit OrganizationalUnitIdParameter
		The OrganizationalUnit parameter filters the results based on the object's location in Active Directory. Only objects that exist in the specified location are returned. Valid input for this parameter is an organizational unit (OU) or domain that's visible using the Get-OrganizationalUnit cmdlet. You can use any value that uniquely identifies the OU or domain. For example:
		Name
		Canonical name
		Distinguished name (DN)
		GUID

	.PARAMETER PublicFolder SwitchParameter
		The PublicFolder switch filters the results by public folder mailboxes. When you use this switch, only public folder mailboxes are included in the results. This switch is required to return public folder mailboxes. You don't need to specify a value with this switch.

	.PARAMETER ReadFromDomainController SwitchParameter
		This parameter is available only in on-premises Exchange 2016.

		The ReadFromDomainController switch specifies that information should be read from a domain controller in the user's domain. If you run the command Set-AdServerSettings -ViewEntireForest $true to include all objects in the forest and you don't use the ReadFromDomainController switch, it's possible that information will be read from a global catalog that has outdated information. When you use the ReadFromDomainController switch, multiple reads might be necessary to get the information. You don't have to specify a value with this switch.

		By default, the recipient scope is set to the domain that hosts your Exchange servers.

	.PARAMETER RecipientTypeDetails RecipientTypeDetails[]
		The RecipientTypeDetails parameter specifies the type of recipients returned. Recipient types are divided into recipient types and subtypes. Each recipient type contains all common properties for all subtypes. For example, the type UserMailbox represents a user account in Active Directory that has an associated mailbox. Because there are several mailbox types, each mailbox type is identified by the RecipientTypeDetails parameter. For example, a conference room mailbox has RecipientTypeDetails set to ConferenceRoomMailbox, whereas a user mailbox has RecipientTypeDetails set to UserMailbox.

		You can select from the following values:
		ArbitrationMailbox
		ConferenceRoomMailbox
		Contact
		DiscoveryMailbox
		DynamicDistributionGroup
		EquipmentMailbox
		ExternalManagedContact
		ExternalManagedDistributionGroup
		LegacyMailbox
		LinkedMailbox
		MailboxPlan
		MailContact
		MailForestContact
		MailNonUniversalGroup
		MailUniversalDistributionGroup
		MailUniversalSecurityGroup
		MailUser
		PublicFolder
		RoleGroup
		RoomList
		RoomMailbox
		SharedMailbox
		SystemAttendantMailbox
		SystemMailbox
		User
		UserMailbox

	.PARAMETER RemoteArchive SwitchParameter
		This parameter is available only in on-premises Exchange 2016.

		The RemoteArchiveswitch filters the results by remote archive mailboxes. When you use this switch, only remote archive mailboxes are included in the results. This switch is required to return remote archive mailboxes. You don't need to specify a value with this switch.

		Remote archive mailboxes are archive mailboxes in the cloud-based service that are associated with mailbox users in on-premises Exchange organizations.

	.PARAMETER ResultSize Unlimited
		The ResultSize parameter specifies the maximum number of results to return. If you want to return all requests that match the query, use unlimited for the value of this parameter. The default value is 1000.

	.PARAMETER Server ServerIdParameter
		This parameter is available only in on-premises Exchange 2016.

		The Server parameter filters the results by Exchange server. When you use this parameter, only mailboxes on the specified Exchange server are included in the results.

		You can use any value that uniquely identifies the server. For example:

		Example: Exchange01
		Example: CN=Exchange01,CN=Servers,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups,CN=First Organization,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=contoso,DC=com
		Example: /o=First Organization/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Configuration/cn=Servers/cn=Exchange01
		Example: bc014a0d-1509-4ecc-b569-f077eec54942
		You can't use this parameter with the Anr, Database, or Identity parameters.

		The ServerName and ServerLegacyDN properties for a mailbox may not be updated immediately after a mailbox move within a database availability group (DAG). To get the most up-to-date values for these mailbox properties, run the command Get-Mailbox <Identity> | Get-MailboxStatistics | Format-List Name,ServerName,ServerLegacyDN.

	.PARAMETER SoftDeletedMailbox SwitchParameter
		This parameter is available only in the cloud-based service.

		The SoftDeletedMailbox switch filters the results by soft-deleted mailboxes. When you use this switch, only soft-deleted mailboxes are included in the results. This switch is required to return soft-deleted mailboxes. You don't need to specify a value with this switch.

		Soft-deleted mailboxes are deleted mailboxes that are still recoverable.

	.PARAMETER SortBy String
		The SortBy parameter specifies the property to sort the results by. You can sort by only one property at a time. The results are sorted in ascending order.

		If the default view doesn't include the property you're sorting by, you can append the command with | Format-Table -Auto <Property1>,<Property2>... to create a new view that contains all of the properties that you want to see. Wildcards (*) in the property names are supported.

		You can sort by the following properties:
		Name
		DisplayName
		Alias
		Office
		ServerLegacyDN

	.EXAMPLE
		This example returns a summary list of all the mailboxes in your organization.

		Get-Mailbox -ResultSize unlimited

		

	.EXAMPLE
		This example returns a list of all the mailboxes in your organization in the Users OU.

		Get-Mailbox -OrganizationalUnit Users

		

	.EXAMPLE
		This example returns all the mailboxes that resolve from the ambiguous name resolution search on the string "Chr". This example returns mailboxes for users such as Chris Ashton, Christian Hess, and Christa Geller.

		Get-Mailbox -Anr Chr

		

	.EXAMPLE
		This example returns a summary list of all archive mailboxes on the Mailbox server named Mailbox01.

		Get-Mailbox -Archive -Server Mailbox01

		

	.EXAMPLE
		This example returns information about the remote archive mailbox for the user ed@contoso.com.

		Get-Mailbox -Identity ed@contoso.com -RemoteArchive

		
#>
[CmdletBinding(
	SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets.
)]
#[System.Diagnostics.DebuggerHidden()]
Param(

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Anr = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $Arbitration = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $Archive = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $AuditLog = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $AuxAuditLog = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[System.Management.Automation.Credential()] $Credential = [System.Management.Automation.PSCredential]::Empty,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Database = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $DomainController = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Filter = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $GroupMailbox = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	Position=1)]
	[String] $Identity = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $IgnoreDefaultScope = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $InactiveMailboxOnly = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $IncludeInactiveMailbox = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $MailboxPlan = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $Monitoring = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $OrganizationalUnit = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $PublicFolder = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $ReadFromDomainController = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $RecipientTypeDetails = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $RemoteArchive = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $ResultSize = 'Unlimited',

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Server = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $SoftDeletedMailbox = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $SortBy = $NULL,

#region Script Header

	[Parameter( HelpMessage='Specifies a user name for the credential in User Principal Name (UPN) format, such as "user@domain.com".' )] 
		[String] $CredentialUserName = $NULL,
	
	[Parameter( HelpMessage='Specifies file name where the secure credential password file is located.  The default of null will prompt for the credentials.' )] 
		[String] $CredentialPasswordFilePath = $NULL,

	[Parameter( HelpMessage='Specify the script''s execution environment source.  Must be either ''ComputerName'', ''DomainName'', ''msExchOrganizationName'' or an arbitrary string. Defaults is msExchOrganizationName.' ) ]
		[String] $ExecutionSource = $NULL,
	
	[Parameter( HelpMessage='Optional string added to the end of the output file name.' ) ]
		[String] $OutFileNameTag = $NULL,
		
	[Parameter( HelpMessage='Specify where to write the output file.' ) ]
		[String] $OutFolderPath = '.\Reports',
	
	[Parameter( HelpMessage='When enabled, only unhealthy items are reported.' ) ]
		[Switch] $AlertOnly = $FALSE,
	
	[Parameter( HelpMessage='Optionally specify the address from which the mail is sent.' ) ]
		[String] $MailFrom = $NULL,
	
	[Parameter( HelpMessage='Optioanlly specify the addresses to which the mail is sent.' ) ]
		[String[]] $MailTo = $NULL,
	
	[Parameter( HelpMessage='Optionally specify the name of the SMTP server that sends the mail message.' ) ]
		[String] $MailServer = $NULL,

	[Parameter( HelpMessage='If the mail message attachment is over this size compress (zip) it.' ) ]
		[Int] $CompressAttachmentLargerThan = 5MB
)

#Requires -version 3
Set-StrictMode -Version Latest

# Detect cmdlet common parameters. 
$cmdletBoundParameters = $PSCmdlet.MyInvocation.BoundParameters
$Debug = If ( $cmdletBoundParameters.ContainsKey('Debug') ) { $cmdletBoundParameters['Debug'] } Else { $FALSE }
# Replace default -Debug preference from 'Inquire' to 'Continue'.  
If ( $DebugPreference -Eq 'Inquire' ) { $DebugPreference = 'Continue' }
$Verbose = If ( $cmdletBoundParameters.ContainsKey('Verbose') ) { $cmdletBoundParameters['Verbose'] } Else { $FALSE }
$WhatIf = If ( $cmdletBoundParameters.ContainsKey('WhatIf') ) { $cmdletBoundParameters['WhatIf'] } Else { $FALSE }
Remove-Variable -Name cmdletBoundParameters -WhatIf:$FALSE

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Collect script execution metrics.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

$scriptStartTime = Get-Date
Write-Verbose "`$scriptStartTime:,$($scriptStartTime.ToString('s'))" 

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Support PSCredential securely in batch mode.  
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

If ( $Credential -Eq [System.Management.Automation.PSCredential]::Empty ) {
	# If $CredentialUserName is a period (.) then get current user account's userPrincipalName (UPN).
	If ( $CredentialUserName -EQ '.' ) { $CredentialUserName = ( [AdsiSearcher] "(&(objectCategory=User)(sAMAccountName=$Env:USERNAME))" ).FindOne().Properties['userprincipalname'] }
	# If $Credentials is empty and credential user name and password file is supplied then use the CurrentUser and LocalMachine security tokens to read the encrypted file.  
	If ( $CredentialUserName -And $CredentialPasswordFilePath ) {
		# Read and convert encoded text into an in-memory secure string.
		$credentialPassword = Get-Content -Path $CredentialPasswordFilePath | ConvertTo-SecureString
		$Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $CredentialUserName, $credentialPassword
	# If $Credentials is empty and credential user name is supplied then securely prompt (interactive) for the user's password.
	} ElseIf ( $CredentialUserName ) {
		$Credential = Get-Credential -Credential $CredentialUserName
	}
}
# Write a secure (encrypted) string file. SecureString is encoded with default Data Protection API (DPAPI) with a DataProtectionScope of CurrentUser and LocalMachine security tokens.
# Read-Host -AsSecureString "Securely enter password" | ConvertFrom-SecureString | Out-File -FilePath '.\SecureCredentialPassword.txt'

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Add Exchange Management Shell snap-in.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	
# Close any existing WinRM Microsoft Exchange PSSessions.
Get-PSSession |
	Where-Object { $_.ConfigurationName -Eq 'Microsoft.Exchange' } |
	ForEach-Object {
		Remove-PSSession -InstanceId $_.InstanceId
	}

# Open a new Microsoft Exchange PSSession to a random specified server.  
$exchangeSession = New-PSSession -Name ExO -Credential $Credential -ConnectionUri "https://ps.outlook.com/powershell/" -ConfigurationName Microsoft.Exchange -Authentication Basic -AllowRedirection -ErrorAction Stop
#Trap { 
#	Get-PSSession |
#		Where-Object { $_.ConfigurationName -Eq 'Microsoft.Exchange' } |
#		ForEach-Object {
#			Remove-PSSession -InstanceId $_.InstanceId
#		} 
#}
Write-Debug "$exchangeSession.ComputerName:,$($exchangeSession.ComputerName)"

# Import session module.  
$moduleInfo = Import-PSSession $exchangeSession -AllowClobber -ErrorAction Stop
## $moduleInfo.ExportedFunctions.Keys | Where-Object { $_ -Like 'Get-*' }

# Try to get tenant's PowerShell throttling budget.  
Try {
	$throttlingPolicy = Get-ThrottlingPolicy
	$powerShellMaxCmdletsTimePeriodSeconds = $throttlingPolicy.PowerShellMaxCmdletsTimePeriod[0]
} Catch {
	$powerShellMaxCmdletsTimePeriodSeconds = 5 # Default tenant of 5 seconds
}
If ( $powerShellMaxCmdletsTimePeriodSeconds -Eq 'Unlimited' ) {
	$powerShellMaxCmdletsTimePeriodSeconds = 0.1
}
Write-Debug "`$powerShellMaxCmdletsTimePeriodSeconds:,$powerShellMaxCmdletsTimePeriodSeconds"

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Include external functions.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

. .\New-OutFilePathBase.ps1
. .\Format-ExpandAllProperties3.ps1
#. .\Get-DSObject2.ps1

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Define internal functions.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Build output and log file path name.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

$outFilePathBase = New-OutFilePathBase -OutFolderPath $OutFolderPath -ExecutionSource $ExecutionSource -OutFileNameTag $OutFileNameTag 

$outFilePathName = "$($outFilePathBase.Value).csv"
Write-Debug "`$outFilePathName: $outFilePathName"
$logFilePathName = "$($outFilePathBase.Value).log"
Write-Debug "`$logFilePathName: $logFilePathName"

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Optionally start or restart PowerShell transcript.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

If ( $Debug -Or $Verbose ) {
	Try {
		Start-Transcript -Path $logFilePathName -WhatIf:$FALSE
	} Catch {
		Stop-Transcript
		Start-Transcript -Path $logFilePathName -WhatIf:$FALSE
	}
}

#endregion Script Header

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Collect report information
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

# Create a hash table to splat parameters.  
$parameters = @{}
If ( $Anr ) { $parameters.Anr = $Anr }
If ( $Arbitration ) { $parameters.Arbitration = $Arbitration }
If ( $Archive ) { $parameters.Archive = $Archive }
If ( $AuditLog ) { $parameters.AuditLog = $AuditLog }
If ( $AuxAuditLog ) { $parameters.AuxAuditLog = $AuxAuditLog }
#If ( $Credential -NE [System.Management.Automation.PSCredential]::Empty ) { $parameters.Credential = $Credential }
If ( $Database ) { $parameters.Database = $Database }
If ( $DomainController ) { $parameters.DomainController = $DomainController }
If ( $Filter ) { $parameters.Filter = $Filter }
If ( $GroupMailbox ) { $parameters.GroupMailbox = $GroupMailbox }
If ( $Identity ) { $parameters.Identity = $Identity }
If ( $IgnoreDefaultScope ) { $parameters.IgnoreDefaultScope = $IgnoreDefaultScope }
If ( $InactiveMailboxOnly ) { $parameters.InactiveMailboxOnly = $InactiveMailboxOnly }
If ( $IncludeInactiveMailbox ) { $parameters.IncludeInactiveMailbox = $IncludeInactiveMailbox }
If ( $MailboxPlan ) { $parameters.MailboxPlan = $MailboxPlan }
If ( $Monitoring ) { $parameters.Monitoring = $Monitoring }
If ( $OrganizationalUnit ) { $parameters.OrganizationalUnit = $OrganizationalUnit }
If ( $PublicFolder ) { $parameters.PublicFolder = $PublicFolder }
If ( $ReadFromDomainController ) { $parameters.ReadFromDomainController = $ReadFromDomainController }
If ( $RecipientTypeDetails ) { $parameters.RecipientTypeDetails = $RecipientTypeDetails }
If ( $RemoteArchive ) { $parameters.RemoteArchive = $RemoteArchive }
If ( $ResultSize ) { $parameters.ResultSize = $ResultSize }
If ( $Server ) { $parameters.Server = $Server }
If ( $SoftDeletedMailbox ) { $parameters.SoftDeletedMailbox = $SoftDeletedMailbox }
If ( $SortBy ) { $parameters.SortBy = $SortBy }
If ( $Debug ) {
	ForEach ( $key In $parameters.Keys ) {
		Write-Debug "`$parameters[$key]`:,$($parameters[$key])"
	}
}

# Build Report
$report = @()
$report = Get-Mailbox @parameters |
	Format-ExpandAllProperties -ConvertByteQuantifiedSizeToBytes

# Optionally write report information.
If ( $report ) {
	$report | 
		Export-CSV -Path $outFilePathName -NoTypeInformation
}

#region Script Footer

Remove-PSSession $exchangeSession
# Avoid "Micro delay applied" for any subsequent commands.
Start-Sleep -Seconds $powerShellMaxCmdletsTimePeriodSeconds

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Optionally mail report.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	
If ( (Test-Path -PathType Leaf -Path $outFilePathName) -And $MailFrom -And $MailTo -And $MailServer ) {

	# Determine subject line report/alert mode.  
	If ( $AlertOnly ) {
		$reportType = 'Alert'
	} Else {
		$reportType = 'Report'
	}

	$messageSubject = "Get Mailbox $reportType for $($outFilePathBase.ExecutionSourceName) on $((Get-Date).ToString('s'))"

	# If the out file is larger then a specified limit (message size limit), then create a compressed (zipped) copy.  
	Write-Debug "$outFilePathName.Length:,$((Get-ChildItem -LiteralPath $outFilePathName).Length)"
	If ( $CompressAttachmentLargerThan -LT (Get-ChildItem -LiteralPath $outFilePathName).Length ) {
		
		$outZipFilePathName = "$outFilePathName.zip"
		Write-Debug "`$outZipFilePathName:,$outZipFilePathName"
		
		# Create a temporary empty zip file.  
		Set-Content -Path $outZipFilePathName -Value ( "PK" + [Char]5 + [Char]6 + ("$([Char]0)" * 18) ) -Force -WhatIf:$FALSE
		
		# Wait for the zip file to appear in the parent folder.  
		While ( -Not (Test-Path -PathType Leaf -Path $outZipFilePathName) ) {   
			Write-Debug "Waiting for:,$outZipFilePathName"			
			Start-Sleep -Milliseconds 20
		} 
		
		# Wait for the zip file to be written by detecting that the file size is not zero.  
		While ( -Not (Get-ChildItem -LiteralPath $outZipFilePathName).Length ) {
			Write-Debug "Waiting for ($outZipFilePathName\$($outFilePathBase.FileName).csv).Length:,$((Get-ChildItem -LiteralPath $outZipFilePathName).Length)"
			Start-Sleep -Milliseconds 20
		}
		
		# Bind to the zip file as a folder.  
		$outZipFile = (New-Object -ComObject Shell.Application).NameSpace( $outZipFilePathName )
		
		# Copy out file into Zip file.
		$outZipFile.CopyHere( $outFilePathName ) 
		
		# Wait for the compressed file to be appear in the zip file.
		While ( -Not $outZipFile.ParseName("$($outFilePathBase.FileName).csv") ) {  
			Write-Debug "Waiting for:,$outZipFilePathName\$($outFilePathBase.FileName).csv"
			Start-Sleep -Milliseconds 20
		} 
		
		# Wait for the compressed file to be written into the zip file by detecting that the file size is not zero.  
		While ( -Not ($outZipFile.ParseName("$($outFilePathBase.FileName).csv")).Size ) {
			Write-Debug "Waiting for ($outZipFilePathName\$($outFilePathBase.FileName).csv).Size:,$($($outZipFile.ParseName($($outFilePathBase.FileName).csv)).Size)"
			Start-Sleep -Milliseconds 20
		}
		
		# Send the report.  
		Send-MailMessage `
			-From $MailFrom `
			-To $MailTo `
			-SmtpServer $MailServer `
			-Subject $messageSubject `
			-Body 'See attached zipped Excel (CSV) spreadsheet.' `
			-Attachments $outZipFilePathName
			
		# Remove the temporary zip file.  
		Remove-Item -LiteralPath $outZipFilePathName
	
	} Else {
	
		# Send the report.  
		Send-MailMessage `
			-From $MailFrom `
			-To $MailTo `
			-SmtpServer $MailServer `
			-Subject $messageSubject `
			-Body 'See attached Excel (CSV) spreadsheet.' `
			-Attachments $outFilePathName
	} 
}

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Optionally write script execution metrics and stop the Powershell transcript.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

$scriptEndTime = Get-Date
Write-Verbose "`$scriptEndTime:,$($scriptEndTime.ToString('s'))" 
$scriptElapsedTime =  $scriptEndTime - $scriptStartTime
Write-Verbose "`$scriptElapsedTime:,$scriptElapsedTime"
If ( $Debug -Or $Verbose ) {
	Stop-Transcript
}
#endregion Script Footer
