<#
	.SYNOPSIS
		This cmdlet is available only in on-premises Exchange Server 2016.

	.DESCRIPTION
		The Get-FailedContentIndexDocuments cmdlet returns a list of documents that couldn't be indexed. The most common reason is that there was no filter available for that document type or there was an attachment within the document. For example, the PDF filter isn't available by default. If an email message contains a PDF document, because there is no PDF filter, the document is marked as failed content indexing.

		After a new filter is installed, only new messages with attachments of the type for which the filter is installed are indexed. If you want to index older messages for the document type, the mailbox has to be moved.

		The cmdlet output provides details about items in a mailbox that couldn't be indexed, including an error code and the reason for failure.

		You need to be assigned permissions before you can run this cmdlet. Although all parameters for this cmdlet are listed in this topic, you may not have access to some parameters if they're not included in the permissions assigned to you. To see what permissions you need, see the "Get unsearchable items" entry in the Recipients Permissions topic.

	.PARAMETER Identity MailboxIdParameter
		The Identity parameter specifies the mailbox. You can use one of the following values:
		GUID
		Distinguished name (DN)
		Domain\Account
		User principal name (UPN)
		LegacyExchangeDN
		SmtpAddress
		Alias

	.PARAMETER MailboxDatabase DatabaseIdParameter
		The MailboxDatabase parameter specifies the database from which to get the mailbox. You can use the following values:
		GUID of the database
		Database name
		DN

	.PARAMETER Server ServerIdParameter
		The Server parameter specifies a Mailbox server. You can use the following values:
		Name
		GUID
		DN

	.PARAMETER Archive SwitchParameter
		The Archive switch restricts the scope of the cmdlet to the user's archive. When using the Archive switch, you must also specify the Identity parameter.

	.PARAMETER Confirm SwitchParameter
		The Confirm switch specifies whether to show or hide the confirmation prompt. How this switch affects the cmdlet depends on if the cmdlet requires confirmation before proceeding.
		Destructive cmdlets (for example, Remove-* cmdlets) have a built-in pause that forces you to acknowledge the command before proceeding. For these cmdlets, you can skip the confirmation prompt by using this exact syntax: -Confirm:$false.
		Most other cmdlets (for example, New-* and Set-* cmdlets) don't have a built-in pause. For these cmdlets, specifying the Confirm switch without a value introduces a pause that forces you acknowledge the command before proceeding.

	.PARAMETER DomainController Fqdn
		The DomainController parameter specifies the domain controller that's used by this cmdlet to read data from or write data to Active Directory. You identify the domain controller by its fully qualified domain name (FQDN). For example, dc01.contoso.com.

	.PARAMETER EndDate DateTime
		The EndDate parameter specifies the end date of the date range.

		Use the short date format that's defined in the Regional Options settings on the computer where you're running the command. For example, if the computer is configured to use the short date format mm/dd/yyyy, enter 09/01/2015 to specify September 1, 2015. You can enter the date only, or you can enter the date and time of day. If you enter the date and time of day, enclose the value in quotation marks ("), for example, "09/01/2015 5:00 PM".

	.PARAMETER ErrorCode Int32
		The ErrorCode parameter allows you to retrieve documents that failed indexing with a specific error code. You can use the cmdlet without this parameter to list all failed documents for a mailbox, a mailbox database or a Mailbox server. The output includes the error codes and reason for failure. If required, you can then restrict the output to a specific error code from the results.

	.PARAMETER FailureMode Transient | Permanent | All
		The FailureMode parameter specifies the type of error. Use the following values.
		Transient Returns items that couldn't be indexed due to transient errors. Exchange Search attempts to index these items again.
		Permanent Returns items that couldn't be indexed due to a permanent error. Exchange Search does not attempt to index these items again.
		All Returns items that couldn't be indexed regardless of nature of failure.

	.PARAMETER ResultSize Unlimited
		This parameter is reserved for internal Microsoft use.

	.PARAMETER StartDate DateTime
		The StartDate parameter specifies the start date of the date range.

		Use the short date format that's defined in the Regional Options settings on the computer where you're running the command. For example, if the computer is configured to use the short date format mm/dd/yyyy, enter 09/01/2015 to specify September 1, 2015. You can enter the date only, or you can enter the date and time of day. If you enter the date and time of day, enclose the value in quotation marks ("), for example, "09/01/2015 5:00 PM".

	.PARAMETER WhatIf SwitchParameter
		The WhatIf switch simulates the actions of the command. You can use this switch to view the changes that would occur without actually applying those changes. You don't need to specify a value with this switch.

	.EXAMPLE
		This example retrieves a list of items that couldn't be indexed by Exchange Search from the mailbox of user Terry Adams.

		Get-FailedContentIndexDocuments -Identity "Terry Adams"

		

	.EXAMPLE
		This example retrieves a list of items that couldn't be indexed by Exchange Search from the mailbox database Mailbox Database MDB2.

		Get-FailedContentIndexDocuments -MailboxDatabase "Mailbox Database MDB2"

		
#>
[CmdletBinding(
	SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets.
)]
#[System.Diagnostics.DebuggerHidden()]
Param(

	[Parameter(
	ValueFromPipeline=$TRUE,
	Position=1)]
	[String] $Identity = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $MailboxDatabase = $NULL,

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
	[String] $DomainController = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[ValidateScript( { If ( -Not $PSItem -Or [System.DateTime]::TryParse( $PSItem, [Ref](New-Object System.DateTime) ) ) { $TRUE } Else { Throw "'$PSItem' is an invalid date." } } )]
	[String] $EndDate,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Int32] $ErrorCode = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $FailureMode = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $ResultSize = 'Unlimited',

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[ValidateScript( { If ( -Not $PSItem -Or [System.DateTime]::TryParse( $PSItem, [Ref](New-Object System.DateTime) ) ) { $TRUE } Else { Throw "'$PSItem' is an invalid date." } } )]
	[String] $StartDate,

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

# If no start/end dates specified create out files with today's date stamp, otherwise creating out files with yesterday's date.  
If ( -Not ( $StartDate -Or $EndDate ) ) {
	$outFilePathBase = New-OutFilePathBase -DateOffsetDays '-1' -OutFolderPath $OutFolderPath -ExecutionSource $ExecutionSource -OutFileNameTag $OutFileNameTag 
} Else {
	$outFilePathBase = New-OutFilePathBase -OutFolderPath $OutFolderPath -ExecutionSource $ExecutionSource -OutFileNameTag $OutFileNameTag 
}

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

# If no start/end dates specified, default to yesterday.  Yesterday's logs should be closed and represent a full 24 hours.  
If ( -Not ( $StartDate -Or $EndDate ) ) {
	$StartDate = $(Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(-1) # Yesterday at midnight, local time.
	$EndDate = $(Get-Date -Hour 0 -Minute 0 -Second 0).AddSeconds(-1) # Yesterday at 23:59:59, local time. 
}
Write-Debug "`$StartDate:,$StartDate"
Write-Debug "`$EndDate:,$EndDate"

# Create a hash table to splat parameters.  
$getMailboxParameters = @{} 
If ( $Identity ) { $getMailboxParameters.Identity = $Identity }
If ( $ResultSize ) { $getMailboxParameters.ResultSize = $ResultSize }
If ( $Debug ) {
	ForEach ( $key In $getMailboxParameters.Keys ) {
		Write-Debug "`$getMailboxParameters[$key]`:,$($getMailboxParameters[$key])"
	}
}

$parameters = @{}
#If ( $Identity ) { $getMailboxParameters.Identity = $Identity }
If ( $MailboxDatabase ) { $getMailboxParameters.MailboxDatabase = $MailboxDatabase }
If ( $Server ) { $getMailboxParameters.Server = $Server }
If ( $Archive ) { $getMailboxParameters.Archive = $Identity }
If ( $DomainController ) { $getMailboxParameters.DomainController = $DomainController }
#If ( $EndDate ) { $getMailboxParameters.EndDate = $EndDate }
If ( $ErrorCode ) { $getMailboxParameters.ErrorCode = $ErrorCode }
If ( $FailureMode ) { $getMailboxParameters.FailureMode = $FailureMode }
If ( $ResultSize ) { $getMailboxParameters.ResultSize = $ResultSize }
#If ( $StartDate ) { $getMailboxParameters.StartDate = $StartDate }
If ( $Debug ) {
	ForEach ( $key In $parameters.Keys ) {
		Write-Debug "`$parameters[$key]`:,$($parameters[$key])"
	}
}

# Build Report
Get-Mailbox @getMailboxParameters |
	Where-Object { $PSItem.Identity -NotLike 'DiscoverySearchMailbox{*}*' } |
	ForEach-Object {
		$parameters.Identity = $PSItem.UserPrincipalName
		Get-FailedContentIndexDocuments @parameters 
		## Avoid "Micro delay applied" for any subsequent commands.
		#Start-Sleep -Seconds $powerShellMaxCmdletsTimePeriodSeconds
	} |
	Format-ExpandAllProperties -ConvertByteQuantifiedSizeToBytes |
	Export-CSV -Path $outFilePathName -NoTypeInformation

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

	$messageSubject = "Get Failed Content Index Documents $reportType for $($outFilePathBase.ExecutionSourceName) on $((Get-Date).ToString('s'))"

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
