<#
	.SYNOPSIS
		This cmdlet is available in on-premises Exchange Server 2016 and in the cloud-based service. Some parameters and settings may be exclusive to one environment or the other.

	.DESCRIPTION
		The search criteria for the Get-MoveRequest cmdlet is a Boolean And statement. If you use multiple parameters, it narrows your search and reduces your search results.

		You need to be assigned permissions before you can run this cmdlet. Although all parameters for this cmdlet are listed in this topic, you may not have access to some parameters if they're not included in the permissions assigned to you. To see what permissions you need, see the "Mailbox moves" entry in the Recipients Permissions topic.

	.PARAMETER BatchName String
		The BatchName parameter specifies the name that was given to a batch move request.

		You can't use this parameter with the Identity parameter.

	.PARAMETER Credential PSCredential
		This parameter is available only in on-premises Exchange 2016.

		The Credential parameter specifies the user name and password that's used to run this command. Typically, you use this parameter in scripts or when you need to provide different credentials that have the required permissions.

		This parameter requires the creation and passing of a credential object. This credential object is created by using the Get-Credential cmdlet. For more information, see Get-Credential (http://go.microsoft.com/fwlink/p/?linkId=142122).

	.PARAMETER DomainController Fqdn
		This parameter is available only in on-premises Exchange 2016.

		The DomainController parameter specifies the domain controller that's used by this cmdlet to read data from or write data to Active Directory. You identify the domain controller by its fully qualified domain name (FQDN). For example, dc01.contoso.com.

	.PARAMETER Flags None | CrossOrg | IntraOrg | Push | Pull | Offline | Protected | RemoteLegacy | HighPriority | Suspend | SuspendWhenReadyToComplete | MoveOnlyPrimaryMailbox | MoveOnlyArchiveMailbox | TargetIsAggregatedMailbox | Join | Split | MoveOnlyAuxMailbox
		The Flags parameter specifies the move type to retrieve information for. The following values may be used:
		CrossOrg
		HighPriority
		IntraOrg
		Join
		MoveOnlyArchiveMailbox
		MoveOnlyPrimaryMailbox
		None
		Offline
		Protected
		Pull
		Push
		RemoteLegacy
		Split
		Suspend
		SuspendWhenReadyToComplete
		TargetIsAggregatedMailbox

	.PARAMETER HighPriority $true | $false
		This parameter is available only in on-premises Exchange 2016.

		The HighPriority parameter specifies that the cmdlet returns requests that were created with the HighPriority flag. The HighPriority flag indicates that the request should be processed before other lower priority requests in the queue.

		You can't use this parameter with the Identity parameter.

	.PARAMETER Identity MoveRequestIdParameter
		The Identity parameter specifies the identity of the move request, which is the identity of the mailbox or mail user. You can use any value that uniquely identifies the mailbox or mail user.

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
		This parameter can't be used with the following parameters:
		BatchName
		HighPriority
		MoveStatus
		Offline
		Protect
		RemoteHostName
		SourceDatabase
		Suspend
		SuspendWhenReadyToComplete
		TargetDatabase

	.PARAMETER IncludeSoftDeletedObjects SwitchParameter
		This parameter is available only in on-premises Exchange 2016.

		The IncludeSoftDeletedObjects parameter specifies whether to return mailboxes that have been soft deleted. This parameter accepts $true or $false.

	.PARAMETER MoveStatus None | Queued | InProgress | AutoSuspended | CompletionInProgress | Synced | Completed | CompletedWithWarning | Suspended | Failed
		The MoveStatus parameter returns move requests in the specified status. You can use the following values:

		AutoSuspended
		Completed
		CompletedWithWarning
		CompletionInProgress
		Failed
		InProgress
		None
		Suspended
		You can't use this parameter with the Identity parameter.

	.PARAMETER Offline $true | $false
		The Offline parameter specifies whether to return mailboxes that are being moved in offline mode. This parameter accepts $true or $false.

		You can't use this parameter with the Identity parameter.

	.PARAMETER OrganizationalUnit OrganizationalUnitIdParameter
		The OrganizationalUnit parameter filters the results based on the object's location in Active Directory. Only objects that exist in the specified location are returned. Valid input for this parameter is an organizational unit (OU) or domain that's visible using the Get-OrganizationalUnit cmdlet. You can use any value that uniquely identifies the OU or domain. For example:
		Name
		Canonical name
		Distinguished name (DN)
		GUID

	.PARAMETER Protect $true | $false
		This parameter is available only in on-premises Exchange 2016.

		The Protect parameter returns mailboxes being moved in protected mode. This parameter accepts $true or $false.

		You can't use this parameter with the Identity parameter.

	.PARAMETER ProxyToMailbox MailboxIdParameter
		PARAMVALUE: MailboxIdParameter

	.PARAMETER RemoteHostName Fqdn
		The RemoteHostName parameter specifies the FQDN of the cross-forest organization from which you're moving the mailbox.

		You can't use this parameter with the Identity parameter.

	.PARAMETER ResultSize Unlimited
		The ResultSize parameter specifies the maximum number of results to return. If you want to return all requests that match the query, use unlimited for the value of this parameter. The default value is 1000.

	.PARAMETER SortBy String
		The SortBy parameter specifies the property to sort the results by. You can sort by only one property at a time. The results are sorted in ascending order.

		If the default view doesn't include the property you're sorting by, you can append the command with | Format-Table -Auto <Property1>,<Property2>... to create a new view that contains all of the properties that you want to see. Wildcards (*) in the property names are supported.

		You can sort by the following properties:
		Name
		DisplayName
		Alias

	.PARAMETER SourceDatabase DatabaseIdParameter
		This parameter is available only in on-premises Exchange 2016.

		The SourceDatabase parameter specifies that all mailboxes being moved from the specified source database are returned. You can use the following values:

		GUID of the database
		Database name
		You can't use this parameter with the Identity parameter.

	.PARAMETER Suspend $true | $false
		The Suspend parameter specifies whether to return mailboxes with moves that have been suspended. This parameter accepts $true or $false.

		You can't use this parameter with the Identity parameter.

	.PARAMETER SuspendWhenReadyToComplete $true | $false
		The SuspendWhenReadytoComplete parameter specifies whether to return mailboxes that have been moved with the New-MoveRequest command and its SuspendWhenReadyToComplete switch. This parameter accepts $true or $false.

		You can't use this parameter with the Identity parameter.

	.PARAMETER TargetDatabase DatabaseIdParameter
		This parameter is available only in on-premises Exchange 2016.

		The TargetDatabase parameter specifies whether to return all mailboxes that are being moved to the specified target database. You can use the following values:

		GUID of the database
		Database name
		You can't use this parameter with the Identity parameter.

	.EXAMPLE
		This example retrieves the status of the ongoing mailbox move for Tony Smith's mailbox (tony@contoso.com).

		Get-MoveRequest -Identity 'tony@contoso.com'

		

	.EXAMPLE
		This example retrieves the status of ongoing mailbox moves to the target database DB05.

		Get-MoveRequest -MoveStatus InProgress -TargetDatabase DB05

		

	.EXAMPLE
		This example retrieves the status of move requests in the FromDB01ToDB02 batch that completed, but had warnings.

		Get-MoveRequest -BatchName "FromDB01ToDB02" -MoveStatus CompletedWithWarning

		
#>
[CmdletBinding(
	SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets.
)]
#[System.Diagnostics.DebuggerHidden()]
Param(

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $BatchName = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[System.Management.Automation.Credential()] $Credential = [System.Management.Automation.PSCredential]::Empty,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $DomainController = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Flags = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $HighPriority = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	Position=1)]
	[String] $Identity = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $IncludeSoftDeletedObjects = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $MoveStatus = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Offline = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $OrganizationalUnit = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Protect = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $ProxyToMailbox = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $RemoteHostName = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $ResultSize = 'Unlimited',

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $SortBy = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $SourceDatabase = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Suspend = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $SuspendWhenReadyToComplete = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $TargetDatabase = $NULL,

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
If ( $BatchName ) { $parameters.BatchName = $BatchName }
#If ( $Credential -NE [System.Management.Automation.PSCredential]::Empty ) { $parameters.Credential = $Credential }
If ( $DomainController ) { $parameters.DomainController = $DomainController }
If ( $Flags ) { $parameters.Flags = $Flags }
If ( $HighPriority ) { $parameters.HighPriority = $HighPriority }
If ( $Identity ) { $parameters.Identity = $Identity }
If ( $IncludeSoftDeletedObjects ) { $parameters.IncludeSoftDeletedObjects = $IncludeSoftDeletedObjects }
If ( $MoveStatus ) { $parameters.MoveStatus = $MoveStatus }
If ( $Offline ) { $parameters.Offline = $Offline }
If ( $OrganizationalUnit ) { $parameters.OrganizationalUnit = $OrganizationalUnit }
If ( $Protect ) { $parameters.Protect = $Protect }
If ( $ProxyToMailbox ) { $parameters.ProxyToMailbox = $ProxyToMailbox }
If ( $RemoteHostName ) { $parameters.RemoteHostName = $RemoteHostName }
If ( $ResultSize ) { $parameters.ResultSize = $ResultSize }
If ( $SortBy ) { $parameters.SortBy = $SortBy }
If ( $SourceDatabase ) { $parameters.SourceDatabase = $SourceDatabase }
If ( $Suspend ) { $parameters.Suspend = $Suspend }
If ( $SuspendWhenReadyToComplete ) { $parameters.SuspendWhenReadyToComplete = $SuspendWhenReadyToComplete }
If ( $TargetDatabase ) { $parameters.TargetDatabase = $TargetDatabase }
If ( $Debug ) {
	ForEach ( $key In $parameters.Keys ) {
		Write-Debug "`$parameters[$key]`:,$($parameters[$key])"
	}
}

# Build Report
$report = @()
$report = Get-MoveRequest @parameters |
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

	$messageSubject = "Get Move Request $reportType for $($outFilePathBase.ExecutionSourceName) on $((Get-Date).ToString('s'))"

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
