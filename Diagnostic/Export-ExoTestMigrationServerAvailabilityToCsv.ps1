<#
	.SYNOPSIS
		This cmdlet is available in on-premises Exchange Server 2016 and in the cloud-based service. Some parameters and settings may be exclusive to one environment or the other.

	.DESCRIPTION
		The Test-MigrationServerAvailability cmdlet verifies that you can communicate with the on-premises mail server that houses the mailbox data that you want to migrate to cloud-based mailboxes. When you run this cmdlet, you must specify the migration type. You can specify whether to communicate with an IMAP server or with an Exchange server.

		For an IMAP migration, this cmdlet uses the server's fully qualified domain name (FQDN) and a port number to verify the connection. If the verification is successful, use the same connection settings when you create a migration request with the New-MigrationBatch cmdlet.

		For an Exchange migration, this cmdlet uses one of the following settings to communicate with the on-premises server:

		For Exchange 2003, it uses the server's FQDN and credentials for an administrator account that can access the server.
		For Exchange Server 2007 and later versions, you can connect using the Autodiscover service and the email address of an administrator account that can access the server.
		If the verification is successful, you can use the same settings to create a migration endpoint. For more information, see:

		New-MigrationEndpoint
		New-MigrationBatch
		You need to be assigned permissions before you can run this cmdlet. Although all parameters for this cmdlet are listed in this topic, you may not have access to some parameters if they're not included in the permissions assigned to you. To see what permissions you need, see the "Mailbox Move and Migration Permissions" section in the Recipients Permissions topic.

	.PARAMETER Autodiscover SwitchParameter
		The Autodiscover parameter specifies that the cmdlet should use the Autodiscover service to obtain the connection settings for the target server.

	.PARAMETER Compliance SwitchParameter
		PARAMVALUE: SwitchParameter

	.PARAMETER Credentials PSCredential
		The Credentials parameter specifies the logon credentials for an account that can access mailboxes on the target server. Specify the username in the domain\username format or the user principal name (UPN) (user@example.com) format.

		This parameter requires you to create a credentials object by using the Get-Credential cmdlet. For more information, see Get-Credential (http://go.microsoft.com/fwlink/p/?linkId=142122).

	.PARAMETER EmailAddress SmtpAddress
		The EmailAddress parameter specifies the email address of an administrator account that can access the remote server. This parameter is required when you use the Autodiscover parameter.

	.PARAMETER Endpoint MigrationEndpointIdParameter
		The Endpoint parameter specifies the name of the migration endpoint to connect to. A migration endpoint contains the connection settings and other migration configuration settings. If you include this parameter, the Test-MigrationServerAvailability cmdlet attempts to verify the ability to connect to the remote server using the settings in the migration endpoint.

	.PARAMETER ExchangeOutlookAnywhere SwitchParameter
		This parameter is available only in the cloud-based service.

		The ExchangeOutlookAnywhere parameter specifies a migration type for migrating on-premises mailboxes to Exchange Online. Use this parameter if you plan to migrate mailboxes to Exchange Online using a staged Exchange migration or a cutover Exchange migration.

	.PARAMETER ExchangeRemoteMove SwitchParameter
		The ExchangeRemoteMove parameter specifies a type of migration where mailboxes are moved with full fidelity between two on-premises forests or between an on-premises forest and Exchange Online. Use this parameter if you plan to perform a cross-forest move or migrate mailboxes between an on-premises Exchange organization and Exchange Online in a hybrid deployment.

	.PARAMETER ExchangeServer String
		This parameter is available only in the cloud-based service.

		The ExchangeServer parameter specifies the FQDN of the on-premises Exchange server. Use this parameter when you plan to perform a staged Exchange migration or a cutover Exchange migration. This parameter is required if you don't use the Autodiscover parameter.

	.PARAMETER Imap SwitchParameter
		This parameter is available only in the cloud-based service.

		The Imap parameter specifies an IMAP migration as the migration type. This parameter is required when you want to migrate data from an IMAP mail server to Exchange Online mailboxes.

	.PARAMETER Port Int32
		This parameter is available only in the cloud-based service.

		The Port parameter specifies the TCP port number used by the IMAP migration process to connect to the target server. This parameter is required only for IMAP migrations.

		The standard is to use port 143 for unencrypted connections, port 143 for Transport Layer Security (TLS), and port 993 for Secure Sockets Layer (SSL).

	.PARAMETER PSTImport SwitchParameter
		This parameter is reserved for internal Microsoft use.

	.PARAMETER PublicFolder SwitchParameter
		This parameter is reserved for internal Microsoft use.

	.PARAMETER PublicFolderDatabaseServerLegacyDN String
		This parameter is reserved for internal Microsoft use.

	.PARAMETER RemoteServer Fqdn
		The RemoteServer parameter specifies the FQDN of the on-premises mail server. This parameter is required when you want to perform one of the following migration types:
		Cross-forest move
		Remote move (hybrid deployments)
		IMAP migration

	.PARAMETER RPCProxyServer Fqdn
		This parameter is available only in the cloud-based service.

		The RPCProxyServer parameter specifies the FQDN of the RPC proxy server for the on-premises Exchange server. This parameter is required when you don't use the Autodiscover parameter. Use this parameter if you plan to perform a staged Exchange migration or a cutover Exchange migration to migrate mailboxes to Exchange Online.

	.PARAMETER SourceMailboxLegacyDN String
		This parameter is available only in the cloud-based service.

		The SourceMailboxLegacyDN parameter specifies a mailbox on the target server. Use the LegacyExchangeDN for the on-premises test mailbox as the value for this parameter. The cmdlet will attempt to access this mailbox using the credentials for the administrator account on the target server.

	.PARAMETER Authentication Basic | Digest | Ntlm | Fba | WindowsIntegrated | LiveIdFba | LiveIdBasic | WSSecurity | Certificate | NegoEx | OAuth | Adfs | Kerberos | Negotiate | LiveIdNegotiate | Misconfigured
		This parameter is available only in the cloud-based service.

		The Authentication parameter specifies the authentication method used by the on-premises mail server. Use Basic or NTLM. If you don't include this parameter, Basic authentication is used.

		The parameter is only used for cutover Exchange migrations and staged Exchange migrations.

	.PARAMETER Confirm SwitchParameter
		The Confirm switch specifies whether to show or hide the confirmation prompt. How this switch affects the cmdlet depends on if the cmdlet requires confirmation before proceeding.
		Destructive cmdlets (for example, Remove-* cmdlets) have a built-in pause that forces you to acknowledge the command before proceeding. For these cmdlets, you can skip the confirmation prompt by using this exact syntax: -Confirm:$false.
		Most other cmdlets (for example, New-* and Set-* cmdlets) don't have a built-in pause. For these cmdlets, specifying the Confirm switch without a value introduces a pause that forces you acknowledge the command before proceeding.

	.PARAMETER FilePath String
		The FilePath parameter specifies the path containing the PST files when testing a PST Import migration endpoint.

	.PARAMETER MailboxPermission Admin | FullAccess
		This parameter is available only in the cloud-based service.

		The MailboxPermission parameter specifies what permissions are assigned to the migration administrator account defined by the Credentials parameter. You make the permissions assignment to test the connectivity to a user mailbox on the source mail server when you're testing the connection settings in preparation for a staged or cutover Exchange migration or for creating an Exchange Outlook Anywhere migration endpoint.

		Specify one of the following values for the account defined by the Credentials parameter:

		FullAccess The account has been assigned the Full-Access permission to the mailboxes that will be migrated.
		Admin The account is a member of the Domain Admins group in the organization that hosts the mailboxes that will be migrated.
		This parameter isn't used for testing the connection to the remote server for a remote move migration or an IMAP migration.

	.PARAMETER Partition MailboxIdParameter
		This parameter is reserved for internal Microsoft use.

	.PARAMETER Security None | Ssl | Tls
		This parameter is available only in the cloud-based service.

		The Security parameter specifies the encryption method used by the remote mail server. The options are None, Tls, or Ssl. Use this parameter only when testing the connection to an IMAP server or in preparation for creating a migration endpoint for an IMAP migration.

	.PARAMETER TestMailbox MailboxIdParameter
		This parameter is available only in the cloud-based service.

		The TestMailbox parameter specifies a mailbox on the target server. Use the primary SMTP address as the value for this parameter. The cmdlet will attempt to access this mailbox using the credentials for the administrator account on the target server.

	.PARAMETER WhatIf SwitchParameter
		The WhatIf switch simulates the actions of the command. You can use this switch to view the changes that would occur without actually applying those changes. You don't need to specify a value with this switch.

	.EXAMPLE
		For IMAP migrations, this example verifies the connection to the IMAP mail server imap.contoso.com.

		Test-MigrationServerAvailability -Imap -RemoteServer imap.contoso.com -Port 143

		

	.EXAMPLE
		This example uses the Autodiscover and ExchangeOutlookAnywhere parameters to verify the connection to an on-premises Exchange server in preparation for migrating on-premises mailboxes to Exchange Online. You can use a similar example to test the connection settings for a staged Exchange migration or a cutover Exchange migration.

		$Credentials = Get-Credential
		Test-MigrationServerAvailability -ExchangeOutlookAnywhere -Autodiscover -EmailAddress administrator@contoso.com -Credentials $Credentials

		

	.EXAMPLE
		This example verifies the connection to a server running Microsoft Exchange Server 2003 named exch2k3.contoso.com and uses NTLM for the authentication method.

		$Credentials = Get-Credential
		Test-MigrationServerAvailability -ExchangeOutlookAnywhere -ExchangeServer exch2k3.contoso.com -Credentials $Credentials -RPCProxyServer mail.contoso.com -Authentication NTLM

		
#>
[CmdletBinding(
	SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets.
)]
#[System.Diagnostics.DebuggerHidden()]
Param(

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $Autodiscover = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $Compliance = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[System.Management.Automation.Credential()] $Credentials = [System.Management.Automation.PSCredential]::Empty,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $EmailAddress = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Endpoint = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $ExchangeOutlookAnywhere = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $ExchangeRemoteMove = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $ExchangeServer = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $Imap = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Int32] $Port = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $PSTImport = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $PublicFolder = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $PublicFolderDatabaseServerLegacyDN = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $RemoteServer = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $RPCProxyServer = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $SourceMailboxLegacyDN = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Authentication = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $FilePath = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $MailboxPermission = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Partition = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Security = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $TestMailbox = $NULL,

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

If ( $Credentials -Eq [System.Management.Automation.PSCredential]::Empty ) {
	# If $CredentialUserName is a period (.) then get current user account's userPrincipalName (UPN).
	If ( $CredentialUserName -EQ '.' ) { $CredentialUserName = ( [AdsiSearcher] "(&(objectCategory=User)(sAMAccountName=$Env:USERNAME))" ).FindOne().Properties['userprincipalname'] }
	# If $Credentials is empty and credential user name and password file is supplied then use the CurrentUser and LocalMachine security tokens to read the encrypted file.  
	If ( $CredentialUserName -And $CredentialPasswordFilePath ) {
		# Read and convert encoded text into an in-memory secure string.
		$credentialPassword = Get-Content -Path $CredentialPasswordFilePath | ConvertTo-SecureString
		$Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $CredentialUserName, $credentialPassword
	# If $Credentials is empty and credential user name is supplied then securely prompt (interactive) for the user's password.
	} ElseIf ( $CredentialUserName ) {
		$Credentials = Get-Credential -Credential $CredentialUserName
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
$exchangeSession = New-PSSession -Name ExO -Credential $Credentials -ConnectionUri "https://ps.outlook.com/powershell/" -ConfigurationName Microsoft.Exchange -Authentication Basic -AllowRedirection -ErrorAction Stop
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
If ( $Autodiscover ) { $parameters.Autodiscover = $Autodiscover }
If ( $Compliance ) { $parameters.Compliance = $Compliance }
#If ( $Credentials -NE [System.Management.Automation.PSCredential]::Empty ) { $parameters.Credentials = $Credentials }
If ( $EmailAddress ) { $parameters.EmailAddress = $EmailAddress }
If ( $Endpoint ) { $parameters.Endpoint = $Endpoint }
If ( $ExchangeOutlookAnywhere ) { $parameters.ExchangeOutlookAnywhere = $ExchangeOutlookAnywhere }
If ( $ExchangeRemoteMove ) { $parameters.ExchangeRemoteMove = $ExchangeRemoteMove }
If ( $ExchangeServer ) { $parameters.ExchangeServer = $ExchangeServer }
If ( $Imap ) { $parameters.Imap = $Imap }
If ( $Port ) { $parameters.Port = $Port }
If ( $PSTImport ) { $parameters.PSTImport = $PSTImport }
If ( $PublicFolder ) { $parameters.PublicFolder = $PublicFolder }
If ( $PublicFolderDatabaseServerLegacyDN ) { $parameters.PublicFolderDatabaseServerLegacyDN = $PublicFolderDatabaseServerLegacyDN }
If ( $RemoteServer ) { $parameters.RemoteServer = $RemoteServer }
If ( $RPCProxyServer ) { $parameters.RPCProxyServer = $RPCProxyServer }
If ( $SourceMailboxLegacyDN ) { $parameters.SourceMailboxLegacyDN = $SourceMailboxLegacyDN }
If ( $Authentication ) { $parameters.Authentication = $Authentication }
If ( $FilePath ) { $parameters.FilePath = $FilePath }
If ( $MailboxPermission ) { $parameters.MailboxPermission = $MailboxPermission }
If ( $Partition ) { $parameters.Partition = $Partition }
If ( $Security ) { $parameters.Security = $Security }
If ( $TestMailbox ) { $parameters.TestMailbox = $TestMailbox }
If ( $Debug ) {
	ForEach ( $key In $parameters.Keys ) {
		Write-Debug "`$parameters[$key]`:,$($parameters[$key])"
	}
}

# Build Report
$report = @()
$report = Test-MigrationServerAvailability @parameters |
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

	$messageSubject = "Test Migration Server Availability $reportType for $($outFilePathBase.ExecutionSourceName) on $((Get-Date).ToString('s'))"

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
