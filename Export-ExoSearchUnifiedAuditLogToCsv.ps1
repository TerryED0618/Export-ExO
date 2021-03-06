<#
	.SYNOPSIS
		This cmdlet is available only in the cloud-based service.

	.DESCRIPTION
		The Search-UnifiedAuditLog cmdlet presents pages of data based on repeated iterations of the same command. Use SessionId and SessionCommand to repeatedly call the cmdlet until you get zero returns, or hit the maximum number of results based on the session command. To gauge progress, look at the ResultIndex (hits in the current iteration) and ResultCount (hits for all iterations) properties of the data returned by the cmdlet.

		The Search-UnifiedAuditLog cmdlet is available in Exchange Online PowerShell. You can also view events from the unified auditing log by using the Office 365 Security & Compliance Center. For more information, see Search the audit log in the Office 365 Protection Center (http://go.microsoft.com/fwlink/p/?LinkId=708432).

		You need to be assigned permissions before you can run this cmdlet. Although this topic lists all parameters for the cmdlet, you may not have access to some parameters if they're not included in the permissions assigned to you. To find the permissions required to run any cmdlet or parameter in your organization, see Find the permissions required to run any Exchange cmdlet.

	.PARAMETER EndDate ExDateTime
		The StartDate parameter specifies the start date of the date range.

		Use the short date format that's defined in the Regional Options settings on the computer where you're running the command. For example, if the computer is configured to use the short date format mm/dd/yyyy, enter 09/01/2015 to specify September 1, 2015. You can enter the date only, or you can enter the date and time of day. If you enter the date and time of day, enclose the value in quotation marks ("), for example, "09/01/2015 5:00 PM".

	.PARAMETER StartDate ExDateTime
		The EndDate parameter specifies the end date of the date range.

		Use the short date format that's defined in the Regional Options settings on the computer where you're running the command. For example, if the computer is configured to use the short date format mm/dd/yyyy, enter 09/01/2015 to specify September 1, 2015. You can enter the date only, or you can enter the date and time of day. If you enter the date and time of day, enclose the value in quotation marks ("), for example, "09/01/2015 5:00 PM".

	.PARAMETER Formatted SwitchParameter
		If present, the Formatted switch causes attributes (such as RecordType and Operation), which are normally returned as encoded integers to be formatted into descriptive strings.

	.PARAMETER FreeText String
		The FreeText parameter is no longer supported.

	.PARAMETER Identity UnifiedAuditLogEventIdParameter
		The Identity parameter filters the log entries by event ID. You identify the event ID by its GUID value. For example, 12e489f7-251f-4a3b-b022-08d2686be64d.

	.PARAMETER IPAddresses String[]
		Specifies the Internet Protocol (IP) address whose audit records will be returned. Enter multiple IP addresses separated by commas.

	.PARAMETER ObjectIds String[]
		The ObjectIds parameter filters the log entries by object ID. The object ID is the target object that was acted upon, and depends on the RecordType and Operations values of the event. For example, for SharePoint operations, the object ID is the URL path to a file, folder, or site. For Azure Active Directory operations, the object ID is the account name or GUID value of the account.

		The ObjectId value appears in the AuditData (also known as Details) property of the event.

		To enter multiple values, use the following syntax: <value1>,<value2>,...<valueX>. If the values contain spaces or otherwise require quotation marks, use the following syntax: "<value1>","<value2>",..."<valueX>".

	.PARAMETER Operations String[]
		The Operations parameter filters the log entries by operation. The available values for this parameter depend on the RecordType value. For a list of the available values for this parameter, see Search the audit log in the Office 365 Protection Center (http://go.microsoft.com/fwlink/p/?LinkId=708432).

		To enter multiple values, use the following syntax: <value1>,<value2>,...<valueX>. If the values contain spaces or otherwise require quotation marks, use the following syntax: "<value1>","<value2>",..."<valueX>".

	.PARAMETER RecordType ExchangeAdmin | ExchangeItem | ExchangeItemGroup | SharePoint | SyntheticProbe | SharePointFileOperation | OneDrive | AzureActiveDirectory | AzureActiveDirectoryAccountLogon | DataCenterSecurityCmdlet | ComplianceDLPSharePoint | Sway | ComplianceDLPExchange | SharePointSharingOperation | AzureActiveDirectoryStsLogon | SkypeForBusinessPSTNUsage | SkypeForBusinessUsersBlocked
		The RecordType parameter filters the log entries by record type. Valid values are:
		AzureActiveDirectory
		AzureActiveDirectoryAccountLogon
		AzureActiveDirectoryStsLogon
		ComplianceDLPSharePoint
		DataCenterSecurityCmdlet
		ExchangeAdmin
		ExchangeItem
		ExchangeItemGroup
		SharePoint
		SharePointFileOperation
		SharePointSharingOperation

	.PARAMETER ResultSize Int32
		The ResultSize parameter specifies the maximum number of results to return. The default value is 100, maximum is 5,000.

	.PARAMETER SessionCommand Initialize | ReturnLargeSet | ReturnNextPreviewPage
		The SessionCommand parameter specifies how much information to be returned and how it is organized. Valid values are:
		ReturnNextPreviewPage This value causes the cmdlet to return data sorted on date with duplicate records removed. The maximum number of records returned through use of either paging or the ResultSize parameter is 5,000 records.
		ReturnLargeSet This value causes the cmdlet to return unsorted data which may contain duplicates. By using paging, you can access a maximum of 50,000 results.
		Initialize This value is for Microsoft Internal use only.

	.PARAMETER SessionId String
		The SessionId parameter specifies an ID you provide in the form of a string to identify a command (the cmdlet and its parameters) that will be run multiple times to return paged data. The SessionId can be any string value you choose.

		When the cmdlet is called sequentially with the same session ID, the cmdlet will return the data in sequential blocks of the size specified by ResultSize.

	.PARAMETER UserIds String[]
		The UserIds parameter filters the log entries by the ID of the user who performed the action.

		To enter multiple values, use the following syntax: <value1>,<value2>,...<valueX>. If the values contain spaces or otherwise require quotation marks, use the following syntax: "<value1>","<value2>",..."<valueX>".

	.EXAMPLE
		This example searches the unified audit log for all events from May 1, 2015 to May 8, 2015. The data is returned in pages as the command is rerun sequentially while using the same SessionId value.

		Search-UnifiedAuditLog -StartDate 5/1/2015 -EndDate 5/8/2015 -SessionId "UnifiedAuditLogSearch 05/08/15" -SessionCommand ReturnNextPreviewPage

		

	.EXAMPLE
		This example searches the unified audit log for any files accessed in SharePoint Online from May 1, 2015 to May 8, 2015. The data is returned in pages as the command is rerun sequentially while using the same SessionId value.

		Search-UnifiedAuditLog -StartDate 5/1/2015 -EndDate 5/8/2015 -RecordType SharePointFileOperation -Operations FileAccessed -SessionId "WordDocs_SharepointViews" -SessionCommand ReturnNextPreviewPage

		

	.EXAMPLE
		This example searches the unified audit log from May 1, 2015 to May 8, 2015 for all events relating to a specific Word document identified by its ObjectIDs value.

		Search-UnifiedAuditLog -StartDate 5/1/2015 -EndDate 5/8/2015 -ObjectIDs "https://alpinehouse.sharepoint.com/sites/contoso/Departments/SM/International/Shared Documents/Sales Invoice - International.docx"

		
#>
[CmdletBinding(
	SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets.
)]
#[System.Diagnostics.DebuggerHidden()]
Param(

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[ValidateScript( { If ( -Not $PSItem -Or [System.DateTime]::TryParse( $PSItem, [Ref](New-Object System.DateTime) ) ) { $TRUE } Else { Throw "'$PSItem' is an invalid date." } } )]
	[String] $EndDate = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[ValidateScript( { If ( -Not $PSItem -Or [System.DateTime]::TryParse( $PSItem, [Ref](New-Object System.DateTime) ) ) { $TRUE } Else { Throw "'$PSItem' is an invalid date." } } )]
	[String] $StartDate = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $Formatted = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $FreeText = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Identity = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String[]] $IPAddresses = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String[]] $ObjectIds = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String[]] $Operations = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $RecordType = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Int32] $ResultSize = 5000,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $SessionCommand = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $SessionId = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String[]] $UserIds = $NULL,
	
	[Switch] $IncludeLogon = $NULL,

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
$parameters = @{}
If ( $EndDate ) { $parameters.EndDate = $EndDate }
If ( $StartDate ) { $parameters.StartDate = $StartDate }
If ( $Formatted ) { $parameters.Formatted = $Formatted }
If ( $FreeText ) { $parameters.FreeText = $FreeText }
If ( $Identity ) { $parameters.Identity = $Identity }
If ( $IPAddresses ) { $parameters.IPAddresses = $IPAddresses }
If ( $ObjectIds ) { $parameters.ObjectIds = $ObjectIds }
If ( $Operations ) { $parameters.Operations = $Operations }
If ( $RecordType ) { $parameters.RecordType = $RecordType }
If ( $ResultSize ) { $parameters.ResultSize = $ResultSize }
If ( $SessionCommand ) { $parameters.SessionCommand = $SessionCommand }
If ( $SessionId ) { $parameters.SessionId = $SessionId }
If ( $UserIds ) { $parameters.UserIds = $UserIds }
If ( $Debug ) {
	ForEach ( $key In $parameters.Keys ) {
		Write-Debug "`$parameters[$key]`:,$($parameters[$key])"
	}
}

# Build a dynamic Where-Object FilterScript. 
If ( $IncludeLogon ) { 
	$whereFilterScriptString = '$TRUE'
} Else { 
	$whereFilterScriptString = '$PSItem.RecordType -NE ''AzureActiveDirectoryAccountLogon'' '
}
Write-Debug "`$whereFilterScriptString:,$whereFilterScriptString"
$whereFilterScript = [ScriptBlock]::Create( $whereFilterScriptString )

# Build Report
$report = @()

# Continuously query a page at a time until no results.  
$report = Search-UnifiedAuditLog @parameters | 
	Where-Object -FilterScript $whereFilterScript |
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

	$messageSubject = "Search Unified Audit Log $reportType for $($outFilePathBase.ExecutionSourceName) on $((Get-Date).ToString('s'))"

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
