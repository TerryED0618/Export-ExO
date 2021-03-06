<#
	.SYNOPSIS
		This cmdlet is available in on-premises Exchange Server 2016 and in the cloud-based service. Some parameters and settings may be exclusive to one environment or the other.

		Use the Get-ManagementRole cmdlet to view management roles that have been created in your organization.

		For information about the parameter sets in the Syntax section below, see Exchange cmdlet syntax.
		
	.DESCRIPTION
		You can view management roles in several ways, from listing all the roles in your organization to listing only the child roles of a specified parent role. You can also view the details of a specific role by piping the output of the Get-ManagementRole cmdlet to the Format-List cmdlet.

		For more information about management roles, see Understanding management roles.

		You need to be assigned permissions before you can run this cmdlet. Although all parameters for this cmdlet are listed in this topic, you may not have access to some parameters if they're not included in the permissions assigned to you. To see what permissions you need, see the "Management roles" entry in the Role management permissions topic.

	.PARAMETER GetChildren <SwitchParameter>
		The GetChildren parameter retrieves a list of all the roles that were created based on the parent role specified in the Identity parameter. Only the immediate child roles of the parent role are included. The GetChildren parameter can only be used with the Identity and RoleType parameters.

	.PARAMETER Recurse <SwitchParameter>
		The Recurse parameter retrieves a list of all the roles that were created based on the parent role specified in the Identity parameter. The role specified in the Identity parameter, its child roles, and their children are returned. The Recurse parameter can only be used with the Identity and RoleType parameters.

	.PARAMETER Cmdlet <String>
		The Cmdlet parameter returns a list of all roles that include the specified cmdlet.

	.PARAMETER CmdletParameters <String[]>
		The CmdletParameters parameter returns a list of all roles that include the specified parameter or parameters. You can specify more than one parameter by separating each parameter with a comma. If you specify multiple parameters, only the roles that include all of the specified parameters are returned.
	
	.PARAMETER DomainController <Fqdn>
		This parameter is available only in on-premises Exchange 2016.

		The DomainController parameter specifies the domain controller that's used by this cmdlet to read data from or write data to Active Directory. You identify the domain controller by its fully qualified domain name (FQDN). For example, dc01.contoso.com.
		
	.PARAMETER Identity <RoleIdParameter>
		The Identity parameter specifies the role you want to view. If the role you want to view contains spaces, enclose the name in quotation marks ("). You can use the wildcard character (*) and a partial role name to match multiple roles.
	
	.PARAMETER -RoleType <Custom | UnScoped | OrganizationManagement | RecipientManagement | ViewOnlyOrganizationManagement | DistributionGroupManagement | MyDistributionGroups | MyDistributionGroupMembership | UmManagement | RecordsManagement | MyBaseOptions | UmRecipientManagement | HelpdeskRecipientManagement | GALSynchronizationManagement | ApplicationImpersonation | UMPromptManagement | PartnerDelegatedTenantManagement | DiscoveryManagement | CentralAdminManagement | UnScopedRoleManagement | MyContactInformation | MyProfileInformation | MyVoiceMail | MyTextMessaging | MyMailSubscriptions | MyRetentionPolicies | MyOptions | MailRecipients | FederatedSharing | DatabaseAvailabilityGroups | Databases | PublicFolders | AddressLists | RecipientPolicies | DisasterRecovery | Monitoring | DatabaseCopies | UnifiedMessaging | Journaling | RemoteAndAcceptedDomains | EmailAddressPolicies | TransportRules | SendConnectors | EdgeSubscriptions | OrganizationTransportSettings | ExchangeServers | ExchangeVirtualDirectories | ExchangeServerCertificates | POP3AndIMAP4Protocols | ReceiveConnectors | UMMailboxes | UserOptions | SecurityGroupCreationAndMembership | MailRecipientCreation | MessageTracking | RoleManagement | ViewOnlyRecipients | ViewOnlyConfiguration | DistributionGroups | MailEnabledPublicFolders | MoveMailboxes | WorkloadManagement | ResetPassword | AuditLogs | RetentionManagement | SupportDiagnostics | MailboxSearch | LegalHold | MailTips | PublicFolderReplication | ActiveDirectoryPermissions | UMPrompts | Migration | DataCenterOperations | TransportHygiene | TransportQueues | Supervision | CmdletExtensionAgents | OrganizationConfiguration | OrganizationClientAccess | ExchangeConnectors | MailboxImportExport | ViewOnlyCentralAdminManagement | ViewOnlyCentralAdminSupport | ViewOnlyRoleManagement | Reporting | ViewOnlyAuditLogs | TransportAgents | DataCenterDestructiveOperations | InformationRightsManagement | LawEnforcementRequests | MyDiagnostics | MyMailboxDelegation | TeamMailboxes | MyTeamMailboxes | ActiveMonitoring | DataLossPrevention | MyFacebookEnabled | MyLinkedInEnabled | UserApplication | ArchiveApplication | LegalHoldApplication | OfficeExtensionApplication | TeamMailboxLifecycleApplication | CentralAdminCredentialManagement | PersonallyIdentifiableInformation | MailboxSearchApplication | MyMarketplaceApps | MyCustomApps | OrgMarketplaceApps | OrgCustomApps | ExchangeCrossServiceIntegration | NetworkingManagement | AccessToCustomerDataDCOnly | DatacenterOperationsDCOnly | OutlookSupportTier0 | O365SupportViewConfig | OutlookSupportTier1 | OutlookSupportTier3 | MeetingGraphApplication | ComplianceSearch | CaseManagement | Export | Hold | Preview | OutlookSupportTier9 | MyReadWriteMailboxApps | Review | SearchAndPurge | SupervisoryReviewAdmin | ServiceAssuranceView | OutlookSupportPartnersTier0 | OutlookSupportPartnersTier1 | OutlookSupportPartnersTier3 | OutlookSupportPartnersTier9 | DatacenterMailboxManagement | SendMailApplication>
		The RoleType parameter returns a list of roles that match the specified role type. For a list of valid role types, see Understanding management roles.
	
	.PARAMETER Script <String>
		The Script parameter returns a list of all roles that include the specified script.
		
	ScriptParameters <String[]>
		The ScriptParameters parameter returns a list of all roles that include the specified parameter or parameters. You can specify more than one parameter by separating each parameter with a comma. If you specify multiple parameters, only the roles that include all of the specified parameters are returned.
	
	.EXAMPLE
		This example lists all the roles that have been created in your organization.

		Get-ManagementRole

	.EXAMPLE
		This example lists all the roles that are children of the Mail Recipients management role. The command performs a recursive query of all the child roles of the specified parent role. This recursive query finds every child role from the immediate children of the parent to the last child role in the hierarchy. In a recursive list, the parent role is also returned in the list.
		
		Get-ManagementRole "Mail Recipients" -Recurse

	.EXAMPLE
		This example lists all the roles that contain both the Identity and Database parameters. Roles that contain only one parameter or the other aren't returned.

		Get-ManagementRole -CmdletParameters Identity, Database
	
	.EXAMPLE
		This example lists all the roles that have a type of UnScopedTopLevel. These roles contain custom scripts or non-Exchange cmdlets.

		Get-ManagementRole -RoleType UnScopedTopLevel
	
	.EXAMPLE
		This example retrieves only the Transport Rules role and passes the output of the Get-ManagementRole cmdlet to the Format-List cmdlet. The Format-List cmdlet then shows only the Name and RoleType properties of the Transport Rules role. For more information about pipelining and the Format-List cmdlet, see Pipelining and Working with command output.

		Get-ManagementRole "Transport Rules" | Format-List Name, RoleType
	
	.EXAMPLE
		This example lists the immediate children of the Mail Recipients role. Only the child roles that hold the Mail Recipients role as their parent role are returned. The Mail Recipients role isn't returned in the list.

		Get-ManagementRole "Mail Recipients" -GetChildren
#>
[CmdletBinding(
	SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets.
)]
#[System.Diagnostics.DebuggerHidden()]
Param(

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $GetChildren = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $Recurse = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Cmdlet = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String[]] $CmdletParameters = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $DomainController = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Identity = '*',

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $RoleType = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Script = $NULL,
	
	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String[]] $ScriptParameters = $NULL,

#region Script Header

	[System.Management.Automation.Credential()] $Credential = [System.Management.Automation.PSCredential]::Empty,
	
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
If ( $GetChildren ) { $parameters.GetChildren = $GetChildren }
If ( $Recurse ) { $parameters.Recurse = $Recurse }
If ( $Cmdlet ) { $parameters.Cmdlet = $Cmdlet }
If ( $CmdletParameters ) { $parameters.CmdletParameters = $CmdletParameters }
If ( $DomainController ) { $parameters.DomainController = $DomainController }
If ( $Identity ) { $parameters.Identity = $Identity }
If ( $RoleType ) { $parameters.RoleType = $RoleType }
If ( $Script ) { $parameters.Script = $Script }
If ( $ScriptParameters ) { $parameters.ScriptParameters = $ScriptParameters }
If ( $Debug ) {
	ForEach ( $key In $parameters.Keys ) {
		Write-Debug "`$parameters[$key]`:,$($parameters[$key])"
	}
}

# Build Report
$report = @()
$report = Get-ManagementRole @parameters |
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

	$messageSubject = "Get Management Role Entry $reportType for $($outFilePathBase.ExecutionSourceName) on $((Get-Date).ToString('s'))"

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
