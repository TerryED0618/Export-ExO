<#
	.SYNOPSIS
		This cmdlet is available in on-premises Exchange Server 2016 and in the cloud-based service. Some parameters and settings may be exclusive to one environment or the other.

	.DESCRIPTION
		On Mailbox servers only, you can use the Get-MailboxStatistics cmdlet without parameters. In this case, the cmdlet returns the statistics for all mailboxes on all databases on the local server.

		The Get-MailboxStatistics cmdlet requires at least one of the following parameters to complete successfully: Server, Database, or Identity.

		You can use the Get-MailboxStatistics cmdlet to return detailed move history and a move report for completed move requests to troubleshoot a move request. To view the move history, you must pass this cmdlet as an object. Move histories are retained in the mailbox database and are numbered incrementally, and the last executed move request is always numbered 0. For more information, see "Example 7," "Example 8," and "Example 9" in this topic.

		You can only see move reports and move history for completed move requests.

		You need to be assigned permissions before you can run this cmdlet. Although all parameters for this cmdlet are listed in this topic, you may not have access to some parameters if they're not included in the permissions assigned to you. To see what permissions you need, see the "Recipient Provisioning Permissions" section in the Recipients Permissions topic.

	.PARAMETER Database DatabaseIdParameter
		This parameter is available only in on-premises Exchange 2016.

		The Database parameter specifies the name of the mailbox database. When you specify a value for the Database parameter, the Exchange Management Shell returns statistics for all the mailboxes on the database specified.

		You can use the following values:

		GUID
		Database
		This parameter accepts pipeline input from the Get-MailboxDatabase cmdlet.

	.PARAMETER Identity GeneralMailboxOrMailUserIdParameter
		The Identity parameter specifies a mailbox. When you specify a value for the Identity parameter, the command looks up the mailbox specified in the Identity parameter, connects to the server where the mailbox resides, and returns the statistics for the mailbox.

		This parameter accepts the following values:
		Example: JPhillips
		Example: Atlanta.Corp.Contoso.Com/Users/JPhillips
		Example: Jeff Phillips
		Example: CN=JPhillips,CN=Users,DC=Atlanta,DC=Corp,DC=contoso,DC=com
		Example: Atlanta\JPhillips
		Example: fb456636-fe7d-4d58-9d15-5af57d0354c2
		Example: fb456636-fe7d-4d58-9d15-5af57d0354c2@contoso.com
		Example: /o=Contoso/ou=AdministrativeGroup/cn=Recipients/cn=JPhillips
		Example: Jeff.Phillips@contoso.com
		Example: JPhillips@contoso.com

	.PARAMETER Server ServerIdParameter
		This parameter is available only in on-premises Exchange 2016.

		The Server parameter specifies the server from which you want to obtain mailbox statistics. You can use one of the following values:

		Fully qualified domain name (FQDN)
		NetBIOS name
		When you specify a value for the Server parameter, the command returns statistics for all the mailboxes on all the databases, including recovery databases, on the specified server. If you don't specify this parameter, the command returns logon statistics for the local server.

	.PARAMETER Archive SwitchParameter
		The Archive switch parameter specifies whether to return mailbox statistics for the archive mailbox associated with the specified mailbox.

		You don't have to specify a value with this parameter.

	.PARAMETER AuditLog SwitchParameter
		This parameter is reserved for internal Microsoft use.

	.PARAMETER CopyOnServer ServerIdParameter
		This parameter is available only in on-premises Exchange 2016.

		The CopyOnServer parameter is used to retrieve statistics from a specific database copy on the server specified with the Server parameter.

	.PARAMETER DomainController Fqdn
		This parameter is available only in on-premises Exchange 2016.

		The DomainController parameter specifies the domain controller that's used by this cmdlet to read data from or write data to Active Directory. You identify the domain controller by its fully qualified domain name (FQDN). For example, dc01.contoso.com.

	.PARAMETER Filter String
		This parameter is available only in on-premises Exchange 2016.

		The Filter parameter specifies a filter to filter the results of the Get-MailboxStatistics cmdlet. For example, to display all disconnected mailboxes on a specific mailbox database, use the following syntax for this parameter: -Filter 'DisconnectDate -ne $null'

	.PARAMETER IncludeMoveHistory SwitchParameter
		The IncludeMoveHistory switch specifies whether to return additional information about the mailbox that includes the history of a completed move request, such as status, flags, target database, bad items, start times, end times, duration that the move request was in various stages, and failure codes.

	.PARAMETER IncludeMoveReport SwitchParameter
		The IncludeMoveReport switch specifies whether to return a verbose detailed move report for a completed move request, such as server connections and move stages.

		Because the output of this command is verbose, you should send the output to a .CSV file for easier analysis.

	.PARAMETER IncludePassive SwitchParameter
		This parameter is available only in on-premises Exchange 2016.

		Without the IncludePassive parameter, the cmdlet retrieves statistics from active database copies only. Using the IncludePassive parameter, you can have the cmdlet return statistics from all active and passive database copies.

	.PARAMETER IncludeQuarantineDetails SwitchParameter
		This parameter is available only in on-premises Exchange 2016.

		The IncludeQuarantineDetails switch specifies whether to return additional quarantine details about the mailbox that aren't otherwise included in the results. You can use these details to determine when and why the mailbox was quarantined.

		Specifically, this switch returns the values of the QuarantineDescription, QuarantineLastCrash and QuarantineEnd properties on the mailbox. To see these values, you need use a formatting cmdlet. For example, Get-MailboxStatistics <MailboxIdentity> -IncludeQuarantineDetails | Format-List Quarantine*.

	.PARAMETER NoADLookup SwitchParameter
		This parameter is available only in on-premises Exchange 2016.

		The NoADLookup switch specifies that information is retrieved from the mailbox database, and not from Active Directory. This helps improve cmdlet performance when querying a mailbox database that contains a large number of mailboxes.

	.PARAMETER StoreMailboxIdentity StoreMailboxIdParameter
		This parameter is available only in on-premises Exchange 2016.

		The StoreMailboxIdentity parameter specifies the mailbox identity when used with the Database parameter to return statistics for a single mailbox on the specified database. You can use one of the following values:

		MailboxGuid
		LegacyDN
		Use this syntax to retrieve information about disconnected mailboxes, which don't have a corresponding Active Directory object or that has a corresponding Active Directory object that doesn't point to the disconnected mailbox in the mailbox database.

	.EXAMPLE
		This example retrieves the mailbox statistics for the mailbox of the user Ayla Kol by using its associated alias AylaKol.

		Get-MailboxStatistics -Identity AylaKol

		

	.EXAMPLE
		This example retrieves the mailbox statistics for all mailboxes on the server MailboxServer01.

		Get-MailboxStatistics -Server MailboxServer01

		

	.EXAMPLE
		This example retrieves the mailbox statistics for the specified mailbox.

		Get-MailboxStatistics -Identity contoso\chris

		

	.EXAMPLE
		This example retrieves the mailbox statistics for all mailboxes in the specified mailbox database.

		Get-MailboxStatistics -Database "Mailbox Database"

		

	.EXAMPLE
		This example retrieves the mailbox statistics for the disconnected mailboxes for all mailbox databases in the organization. The -ne operator means not equal.

		Get-MailboxDatabase | Get-MailboxStatistics -Filter 'DisconnectDate -ne $null'

		

	.EXAMPLE
		This example retrieves the mailbox statistics for a single disconnected mailbox. The value for the StoreMailboxIdentity parameter is the mailbox GUID of the disconnected mailbox. You can also use the LegacyDN.

		Get-MailboxStatistics -Database "Mailbox Database" -StoreMailboxIdentity 3b475034-303d-49b2-9403-ae022b43742d

		

	.EXAMPLE
		This example returns the summary move history for the completed move request for Ayla Kol's mailbox. If you don't pipeline the output to the Format-List cmdlet, the move history doesn't display.

		Get-MailboxStatistics -Identity AylaKol -IncludeMoveHistory | Format-List

		

	.EXAMPLE
		This example returns the detailed move history for the completed move request for Ayla Kol's mailbox. This example uses a temporary variable to store the mailbox statistics object. If the mailbox has been moved multiple times, there are multiple move reports. The last move report is always MoveReport[0].

		$temp=Get-MailboxStatistics -Identity AylaKol -IncludeMoveHistory
		$temp.MoveHistory[0]

		

	.EXAMPLE
		This example returns the detailed move history and a verbose detailed move report for Ayla Kol's mailbox. This example uses a temporary variable to store the move request statistics object and outputs the move report to a CSV file.

		$temp=Get-MailboxStatistics -Identity AylaKol -IncludeMoveReport
		$temp.MoveHistory[0] | Export-CSV C:\MoveReport_AylaKol.csv

		
#>
[CmdletBinding(
	SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets.
)]
#[System.Diagnostics.DebuggerHidden()]
Param(

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Database = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	Position=1)]
	[String] $Identity = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Server = $NULL,

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
	[String] $CopyOnServer = $NULL,

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
	[Switch] $IncludeMoveHistory = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $IncludeMoveReport = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $IncludePassive = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $IncludeQuarantineDetails = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $NoADLookup = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	Position=1)]
	[String] $StoreMailboxIdentity = $NULL,
		
	[Switch] $IncludeMailboxInfo = $TRUE,
	
	[String[]] $LdapProperties = ( 'distinguishedName','assistant','accountExpires','c','canonicalName','cn','co','comment','company','countryCode','department','description','employeeType','facsimileTelephoneNumber','givenName','homeDirectory','homePhone','info','initials','l','legacyExchangeDN','mail','mailNickname','manager','mobile','msExchAssistantName','msExchHideFromAddressLists','otherTelephone','pager','physicalDeliveryOfficeName','postalCode','proxyAddresses','pwdLastSet','sAMAccountName','sn','st','streetAddress','targetAddress','telephoneAssistant','telephoneNumber','title' ),

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
# Add Active Directory module for Windows PowerShell.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

If ( -Not ( Get-Module ActiveDirectory ) ) {
	Import-Module ActiveDirectory -ErrorAction Stop
}

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Include external functions.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

. .\New-OutFilePathBase.ps1
. .\Format-ExpandAllProperties3.ps1
. .\Get-DSObject2.ps1

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
$getMailboxParameters = @{} 
If ( $Identity ) { $getMailboxParameters.Identity = $Identity }
$getMailboxParameters.ResultSize = 'Unlimited'
If ( $Debug ) {
	ForEach ( $key In $getMailboxParameters.Keys ) {
		Write-Debug "`$getMailboxParameters[$key]`:,$($getMailboxParameters[$key])"
	}
}

$getMailboxStatisticsParameters = @{}
If ( $Database ) { $getMailboxStatisticsParameters.Database = $Database }
If ( $Identity ) { $getMailboxStatisticsParameters.Identity = $Identity }
If ( $Server ) { $getMailboxStatisticsParameters.Server = $Server }
If ( $Archive ) { $getMailboxStatisticsParameters.Archive = $Archive }
If ( $AuditLog ) { $getMailboxStatisticsParameters.AuditLog = $AuditLog }
If ( $CopyOnServer ) { $getMailboxStatisticsParameters.CopyOnServer = $CopyOnServer }
If ( $DomainController ) { $getMailboxStatisticsParameters.DomainController = $DomainController }
If ( $Filter ) { $getMailboxStatisticsParameters.Filter = $Filter }
If ( $IncludeMoveHistory ) { $getMailboxStatisticsParameters.IncludeMoveHistory = $IncludeMoveHistory }
If ( $IncludeMoveReport ) { $getMailboxStatisticsParameters.IncludeMoveReport = $IncludeMoveReport }
If ( $IncludePassive ) { $getMailboxStatisticsParameters.IncludePassive = $IncludePassive }
If ( $IncludeQuarantineDetails ) { $getMailboxStatisticsParameters.IncludeQuarantineDetails = $IncludeQuarantineDetails }
If ( $NoADLookup ) { $getMailboxStatisticsParameters.NoADLookup = $NoADLookup }
If ( $StoreMailboxIdentity ) { $getMailboxStatisticsParameters.StoreMailboxIdentity = $StoreMailboxIdentity }
If ( $Debug ) {
	ForEach ( $key In $getMailboxStatisticsParameters.Keys ) {
		Write-Debug "`$getMailboxStatisticsParameters[$key]`:,$($getMailboxStatisticsParameters[$key])"
	}
}

# Create a persistent connection to an on-premises Active Directory domain controller.
If ( $LdapProperties ) {
	$directorySearcher = Connect-DirectorySearcher
}

# Build Report
Get-Mailbox @getMailboxParameters |
	Where-Object { $PSItem.Identity -NotLike 'DiscoverySearchMailbox{*}*' } |
	ForEach-Object {
		# Avoid "Micro delay applied" for any subsequent commands.
		#Start-Sleep -Seconds $powerShellMaxCmdletsTimePeriodSeconds
		
		$mailbox = $PSItem |
			Format-ExpandAllProperties -ConvertByteQuantifiedSizeToBytes
			
		Write-Verbose "Mailbox:,$($mailbox.Identity)"
		$getMailboxStatisticsParameters.Identity = $mailbox.Identity
		$reportProperties = Get-MailboxStatistics @getMailboxStatisticsParameters | 
			Format-ExpandAllProperties -ConvertByteQuantifiedSizeToBytes
		
		# Optionally append mailbox info for this mailbox. 
		If ( $IncludeMailboxInfo ) {			
			$reportProperties = $mailbox | 
				Format-ExpandAllProperties -AppendTo $reportProperties -ConvertByteQuantifiedSizeToBytes
		}
			
		# Optionally append Directory Services info for this mailbox.  
		If ( $LdapProperties ) {
			Write-Debug "`$mailbox.UserPrincipalName:$($mailbox.UserPrincipalName)"
			#$reportProperties = Get-ADUser -LDAPFilter "(userPrincipalName=$($mailbox.UserPrincipalName))" -Properties $LdapProperties |
			$reportProperties = Get-DSObject -DirectorySearcher $directorySearcher -LDAPFilter "(userPrincipalName=$($mailbox.UserPrincipalName))" -Properties $LdapProperties -ViewEntireForest -AugmentProperties |
				#Select-Object -Property * -ExcludeProperty PropertyNames,AddedProperties,RemovedProperties,ModifiedProperties,PropertyCount |
				Format-ExpandAllProperties -AppendTo $reportProperties 
		}
		
		Write-Output $reportProperties		
		
	} |
	Export-CSV -Path $outFilePathName -Encoding UTF8 -NoTypeInformation


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

	$messageSubject = "Get Mailbox Statistics $reportType for $($outFilePathBase.ExecutionSourceName) on $((Get-Date).ToString('s'))"

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
