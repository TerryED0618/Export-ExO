<#
	.SYNOPSIS
		This cmdlet is available in on-premises Exchange Server 2016 and in the cloud-based service. Some parameters and settings may be exclusive to one environment or the other.

	.DESCRIPTION
		If you run the Search-AdminAuditLog cmdlet without any parameters, up to 1,000 log entries are returned by default.

		You need to be assigned permissions before you can run this cmdlet. Although all parameters for this cmdlet are listed in this topic, you may not have access to some parameters if they're not included in the permissions assigned to you. To see what permissions you need, see the "View-only administrator audit logging" entry in the Exchange infrastructure and PowerShell permissions topic.

	.PARAMETER Cmdlets MultiValuedProperty
		The Cmdlets parameter specifies the cmdlets you want to search for in the administrator audit log. Only the log entries that contain the cmdlets you specify are returned.

		If you want to specify more than one cmdlet, separate each cmdlet with a comma.

	.PARAMETER DomainController Fqdn
		This parameter is available only in on-premises Exchange 2016.

		The DomainController parameter specifies the domain controller that's used by this cmdlet to read data from or write data to Active Directory. You identify the domain controller by its fully qualified domain name (FQDN). For example, dc01.contoso.com.

	.PARAMETER EndDate ExDateTime
		The EndDate parameter specifies the end date of the date range.

		Use the short date format that's defined in the Regional Options settings on the computer where you're running the command. For example, if the computer is configured to use the short date format mm/dd/yyyy, enter 09/01/2015 to specify September 1, 2015. You can enter the date only, or you can enter the date and time of day. If you enter the date and time of day, enclose the value in quotation marks ("), for example, "09/01/2015 5:00 PM".

	.PARAMETER ExternalAccess $true | $false
		The ExternalAccess parameter returns only audit log entries for cmdlets that were run by a user outside of your organization. In Exchange Online, use this parameter to return audit log entries for cmdlets run by Microsoft datacenter administrators.

	.PARAMETER IsSuccess $true | $false
		The IsSuccess parameter specifies whether only administrator audit log entries that indicated a success or failure should be returned. Valid values are $true and $false.

	.PARAMETER ObjectIds MultiValuedProperty
		The ObjectIds parameter specifies that only administrator audit log entries that contain the specified changed objects should be returned. This parameter accepts a variety of objects, such as mailbox aliases, Send connector names, and so on.

		If you want to specify more than one object ID, separate each ID with a comma.

	.PARAMETER Parameters MultiValuedProperty
		The Parameters parameter specifies the parameters you want to search for in the administrator audit log. Only the log entries that contain the parameters you specify are returned. You can only use this parameter if you use the Cmdlets parameter.

		If you want to specify more than one parameter, separate each parameter with a comma.

	.PARAMETER ResultSize Int32
		The ResultSize parameter specifies the maximum number of results to return. If you want to return all requests that match the query, use unlimited for the value of this parameter. The default value is 1000.

	.PARAMETER StartDate ExDateTime
		The StartDate parameter specifies the start date of the date range.

		Use the short date format that's defined in the Regional Options settings on the computer where you're running the command. For example, if the computer is configured to use the short date format mm/dd/yyyy, enter 09/01/2015 to specify September 1, 2015. You can enter the date only, or you can enter the date and time of day. If you enter the date and time of day, enclose the value in quotation marks ("), for example, "09/01/2015 5:00 PM".

	.PARAMETER StartIndex Int32
		The StartIndex parameter specifies the position in the result set where the displayed results start.

	.PARAMETER UserIds MultiValuedProperty
		The UserIds parameter specifies that only the administrator audit log entries that contain the specified ID of the user who ran the cmdlet should be returned.

		If you want to specify more than one user ID, separate each ID with a comma.

	.EXAMPLE
		This example finds all the administrator audit log entries that contain either the New-RoleGroup or the New-ManagementRoleAssignment cmdlet.

		Search-AdminAuditLog -Cmdlets New-RoleGroup, New-ManagementRoleAssignment

		

	.EXAMPLE
		This example finds all the administrator audit log entries that match the following criteria:

		CmdletsSet-Mailbox

		ParametersUseDatabaseQuotaDefaults, ProhibitSendReceiveQuota, ProhibitSendQuota

		StartDate 01/24/2015

		EndDate 02/12/2015

		The command completed successfully
		Search-AdminAuditLog -Cmdlets Set-Mailbox -Parameters UseDatabaseQuotaDefaults, ProhibitSendReceiveQuota, ProhibitSendQuota -StartDate 01/24/2015 -EndDate 02/12/2015 -IsSuccess $true

		

	.EXAMPLE
		This example displays all the comments written to the administrator audit log by the Write-AdminAuditLog cmdlet.

		First, store the audit log entries in a temporary variable using the following command.

		$LogEntries = Search-AdminAuditLog -Cmdlets Write-AdminAuditLog
		Then, iterate through all the audit log entries returned and display the Parameters property using the following command.

		$LogEntries | ForEach { $_.CmdletParameters }

		

	.EXAMPLE
		This example returns entries in the administrator audit log of an Exchange Online organization for cmdlets run by Microsoft datacenter administrators between September 17, 2015 and October 2, 2015.

		Search-AdminAuditLog -ExternalAccess $true -StartDate 09/17/2015 -EndDate 10/02/2015

		
#>
[CmdletBinding(
	SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets.
)]
#[System.Diagnostics.DebuggerHidden()]
Param(

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Cmdlets = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $DomainController = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[ValidateScript( { If ( -Not $PSItem -Or [System.DateTime]::TryParse( $PSItem, [Ref](New-Object System.DateTime) ) ) { $TRUE } Else { Throw "'$PSItem' is an invalid date." } } )]
	[String] $EndDate = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $ExternalAccess = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $IsSuccess = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $ObjectIds = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Parameters = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Int32] $ResultSize = 250000,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[ValidateScript( { If ( -Not $PSItem -Or [System.DateTime]::TryParse( $PSItem, [Ref](New-Object System.DateTime) ) ) { $TRUE } Else { Throw "'$PSItem' is an invalid date." } } )]
	[String] $StartDate = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Int32] $StartIndex = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $UserIds = $NULL,

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
$searchAdminAuditLogParameters = @{}
If ( $Cmdlets ) { $searchAdminAuditLogParameters.Cmdlets = $Cmdlets }
If ( $DomainController ) { $searchAdminAuditLogParameters.DomainController = $DomainController }
If ( $EndDate ) { $searchAdminAuditLogParameters.EndDate = $EndDate }
If ( $ExternalAccess ) { $searchAdminAuditLogParameters.ExternalAccess = $ExternalAccess }
If ( $IsSuccess ) { $searchAdminAuditLogParameters.IsSuccess = $IsSuccess }
If ( $ObjectIds ) { $searchAdminAuditLogParameters.ObjectIds = $ObjectIds }
If ( $Parameters ) { $searchAdminAuditLogParameters.Parameters = $Parameters }
If ( $ResultSize ) { $searchAdminAuditLogParameters.ResultSize = $ResultSize }
If ( $StartDate ) { $searchAdminAuditLogParameters.StartDate = $StartDate }
If ( $StartIndex ) { $searchAdminAuditLogParameters.StartIndex = $StartIndex }
If ( $UserIds ) { $searchAdminAuditLogParameters.UserIds = $UserIds }
If ( $Debug ) {
	ForEach ( $key In $searchAdminAuditLogParameters.Keys ) {
		Write-Debug "`$searchAdminAuditLogParameters[$key]`:,$($searchAdminAuditLogParameters[$key])"
	}
}

# Build Report
$report = @()
$report = Search-AdminAuditLog @searchAdminAuditLogParameters |
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

	$messageSubject = "Search Admin Audit Log $reportType for $($outFilePathBase.ExecutionSourceName) on $((Get-Date).ToString('s'))"

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
