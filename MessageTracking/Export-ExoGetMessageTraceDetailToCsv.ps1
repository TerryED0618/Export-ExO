<#
	.SYNOPSIS
		This cmdlet is available only in the cloud-based service.

	.DESCRIPTION
		You need to be assigned permissions before you can run this cmdlet. Although all parameters for this cmdlet are listed in this topic, you may not have access to some parameters if they're not included in the permissions assigned to you. To see what permissions you need, see the "View reports" entry in the Feature permissions in Exchange Online topic.

	.PARAMETER Action MultiValuedProperty
		The Action parameter filters the report by the action taken by DLP policies, transport rules, malware filtering, or spam filtering. To view the complete list of valid values for this parameter, run the command Get-MailFilterListReport -SelectionTarget Actions. The action you specify must correspond to the report type. For example, you can only specify malware filter actions for malware reports.

		You can specify multiple values separated by commas.

	.PARAMETER EndDate DateTime
		The EndDate parameter specifies the end date of the date range.

		Use the short date format that's defined in the Regional Options settings on the computer where you're running the command. For example, if the computer is configured to use the short date format mm/dd/yyyy, enter 09/01/2015 to specify September 1, 2015. You can enter the date only, or you can enter the date and time of day. If you enter the date and time of day, enclose the value in quotation marks ("), for example, "09/01/2015 5:00 PM".

		When you run a message trace for messages that are greater than 7 days old, results may take up to a few hours. 
		
	.PARAMETER Event MultiValuedProperty
		The Event parameter filters the report by the message event. The following are examples of common events:

		RECEIVE The message was received by the service.
		SEND The message was sent by the service.
		FAIL The message failed to be delivered.
		DELIVER The message was delivered to a mailbox.
		EXPAND The message was sent to a distribution group that was expanded.
		TRANSFER Recipients were moved to a bifurcated message because of content conversion, message recipient limits, or agents.
		DEFER The message delivery was postponed and may be re-attempted later.
		You can specify multiple values separated by commas.

	.PARAMETER Expression Expression
		This parameter is reserved for internal Microsoft use.

	.PARAMETER MessageId String
		The MessageId parameter filters the results by the Message-ID header field of the message. This value is also known as the Client ID. The format of the Message-ID depends on the messaging server that sent the message. The value should be unique for each message. However, not all messaging servers create values for the Message-ID in the same way. Be sure to include the full Message ID string. This may include angle brackets.

	.PARAMETER MessageTraceId Guid
		The MessageTraceId parameter can be used with the recipient address to uniquely identify a message trace and obtain more details. A message trace ID is generated for every message that's processed by the system.

	.PARAMETER Page Int32
		The Page parameter specifies the page number of the results you want to view. Valid input for this parameter is an integer between 1 and 1000. The default value is 1.

	.PARAMETER PageSize Int32
		The PageSize parameter specifies the maximum number of entries per page. Valid input for this parameter is an integer between 1 and 5000. The default value is 1000.

	.PARAMETER ProbeTag String
		This parameter is reserved for internal Microsoft use.

	.PARAMETER RecipientAddress String
		The RecipientAddress parameter filters the results by the recipient's email address. You can specify multiple values separated by commas.

	.PARAMETER SenderAddress String
		The SenderAddress parameter filters the results by the sender's email address. You can specify multiple values separated by commas.

	.PARAMETER StartDate DateTime
		The StartDate parameter specifies the start date of the date range.

		Use the short date format that's defined in the Regional Options settings on the computer where you're running the command. For example, if the computer is configured to use the short date format mm/dd/yyyy, enter 09/01/2015 to specify September 1, 2015. You can enter the date only, or you can enter the date and time of day. If you enter the date and time of day, enclose the value in quotation marks ("), for example, "09/01/2015 5:00 PM".

		When you run a message trace for messages that are greater than 7 days old, results may take up to a few hours. 
		
	.EXAMPLE
		This example uses the Get-MessageTrace cmdlet to retrieve message trace information for messages with the Exchange Network Message ID value 2bbad36aa4674c7ba82f4b307fff549f send by john@contoso.com between June 13, 2012 and June 15, 2012, and pipelines the results to the Get-MessageTraceDetail cmdlet.

		Get-MessageTrace -MessageTraceId 2bbad36aa4674c7ba82f4b307fff549f -SenderAddress john@contoso.com -StartDate 06/13/2012 -EndDate 06/15/2012 | Get-MessageTraceDetail

		
#>
[CmdletBinding(
	SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets.
)]
#[System.Diagnostics.DebuggerHidden()]
Param(

[Parameter(

	# Get-MessageTrace parameters
	
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[ValidateScript( { If ( -Not $PSItem -Or [System.DateTime]::TryParse( $PSItem, [Ref](New-Object System.DateTime) ) ) { $TRUE } Else { Throw "'$PSITem' is an invalid date." } } )]
	[String] $EndDate = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Expression = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $FromIP = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $MessageId = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $MessageTraceId = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[Int32] $Page = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Int32] $PageSize = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $ProbeTag = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $RecipientAddress = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $SenderAddress = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[ValidateScript( { If ( -Not $PSItem -Or [System.DateTime]::TryParse( $PSItem, [Ref](New-Object System.DateTime) ) ) { $TRUE } Else { Throw "'$PSITem' is an invalid date." } } )]
	[String] $StartDate = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Status = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $ToIP = $NULL,
		
	# Get-MessageTraceDetail parameters
	
	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Action = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[ValidateScript( { If ( -Not $PSItem -Or [System.DateTime]::TryParse( $PSItem, [Ref](New-Object System.DateTime) ) ) { $TRUE } Else { Throw "'$PSITem' is an invalid date." } } )]
	#[String] $EndDate,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Event = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[String] $Expression = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[String] $MessageId = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[String] $MessageTraceId = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[Int32] $Page = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[Int32] $PageSize = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[String] $ProbeTag = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[String] $RecipientAddress = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[String] $SenderAddress = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[ValidateScript( { If ( -Not $PSItem -Or [System.DateTime]::TryParse( $PSItem, [Ref](New-Object System.DateTime) ) ) { $TRUE } Else { Throw "'$PSITem' is an invalid date." } } )]
	#[String] $StartDate,

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
#If ( -Not ( $StartDate -Or $EndDate ) ) {
#	$outFilePathBase = New-OutFilePathBase -DateOffsetDays '-1' -OutFolderPath $OutFolderPath -ExecutionSource $ExecutionSource -OutFileNameTag $OutFileNameTag 
#} Else {
	$outFilePathBase = New-OutFilePathBase -OutFolderPath $OutFolderPath -ExecutionSource $ExecutionSource -OutFileNameTag $OutFileNameTag 
#}

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
	$StartDate = $(Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(-7) 
	$EndDate = $(Get-Date -Hour 0 -Minute 0 -Second 0).AddDays(1).AddTicks(-1) # End of today at 23:59:59.9999999, local time.
}
Write-Debug "`$StartDate:,$StartDate"
Write-Debug "`$EndDate:,$EndDate"

# Create a hash table to splat parameters.  
$getMessageTraceParameters = @{} 
If ( $EndDate ) { $getMessageTraceParameters.EndDate = $EndDate }
If ( $Expression ) { $getMessageTraceParameters.Expression = $Expression }
If ( $FromIP ) { $getMessageTraceParameters.FromIP = $FromIP }
If ( $MessageId ) { $getMessageTraceParameters.MessageId = $MessageId }
If ( $MessageTraceId ) { $getMessageTraceParameters.MessageTraceId = $MessageTraceId }
#If ( $Page ) { $getMessageTraceParameters.Page = $Page }
If ( $PageSize ) { $getMessageTraceParameters.PageSize = $PageSize }
If ( $ProbeTag ) { $getMessageTraceParameters.ProbeTag = $ProbeTag }
If ( $RecipientAddress ) { $getMessageTraceParameters.RecipientAddress = $RecipientAddress }
If ( $SenderAddress ) { $getMessageTraceParameters.SenderAddress = $SenderAddress }
If ( $StartDate ) { $getMessageTraceParameters.StartDate = $StartDate }
If ( $Status ) { $getMessageTraceParameters.Status = $Status }
If ( $ToIP ) { $getMessageTraceParameters.ToIP = $ToIP }
If ( $Debug ) {
	ForEach ( $key In $getMessageTraceParameters.Keys ) {
		Write-Debug "`$getMessageTraceParameters[$key]`:,$($getMessageTraceParameters[$key])"
	}
}

$getMessageTraceDetailParameters = @{}
If ( $Action ) { $getMessageTraceDetailParameters.Action = $Action }
#If ( $EndDate ) { $getMessageTraceDetailParameters.EndDate = $EndDate }
If ( $Event ) { $getMessageTraceDetailParameters.Event = $Event }
#If ( $Expression ) { $getMessageTraceDetailParameters.Expression = $Expression }
#If ( $MessageId ) { $getMessageTraceDetailParameters.MessageId = $MessageId }
If ( $MessageTraceId ) { $getMessageTraceDetailParameters.MessageTraceId = $MessageTraceId }
#If ( $Page ) { $getMessageTraceDetailParameters.Page = $Page }
#If ( $PageSize ) { $getMessageTraceDetailParameters.PageSize = $PageSize }
#If ( $ProbeTag ) { $getMessageTraceDetailParameters.ProbeTag = $ProbeTag }
#If ( $RecipientAddress ) { $getMessageTraceDetailParameters.RecipientAddress = $RecipientAddress }
#If ( $SenderAddress ) { $getMessageTraceDetailParameters.SenderAddress = $SenderAddress }
#If ( $StartDate ) { $getMessageTraceDetailParameters.StartDate = $StartDate }
If ( $Debug ) {
	ForEach ( $key In $getMessageTraceDetailParameters.Keys ) {
		Write-Debug "`$getMessageTraceDetailParameters[$key]`:,$($getMessageTraceDetailParameters[$key])"
	}
}

# Build Report
$messageTraces = @()

# Continuously query a page at a time until no results.  
$getMessageTraceParameters.Page = 1
While ( $messageTrace = Get-MessageTrace @getMessageTraceParameters ) {
	Write-Verbose "Collecting message trace page: $($getMessageTraceParameters.Page)"
	Write-Verbose "Collecting message trace count: $($messageTrace.Count)"
	
	# Avoid "Micro delay applied": http://support.microsoft.com/kb/2881759
	Start-Sleep -Seconds $PowerShellMaxCmdletsTimePeriodSeconds
	
	# Collect report properties.
	$messageTraces += $messageTrace
	
	# Increment to request next page.  
	$getMessageTraceParameters.Page++
} 
Write-Verbose "Collected message trace count: $($messageTraces.Count)"

$report = @()
$messageTraces |
		ForEach-Object {
			$messageTrace = $PSItem 
			
			# Pass manditory parameters.
			$getMessageTraceDetailParameters.MessageTraceId = $messageTrace.MessageTraceId
			$getMessageTraceDetailParameters.RecipientAddress = $messageTrace.RecipientAddress
			
			# Continuously query a page at a time until no results. 
			$getMessageTraceDetailParameters.Page = 1
			While ( $reportProperties = Get-MessageTraceDetail @getMessageTraceDetailParameters | Format-ExpandAllProperties -ConvertByteQuantifiedSizeToBytes -AppendTo $messageTrace ) {
				Write-Verbose "Collecting message trace detail page: $($getMessageTraceDetailParameters.Page)"
				Write-Verbose "Collecting message trace detail count: $($reportProperties.Count)"
				
				# Avoid "Micro delay applied": http://support.microsoft.com/kb/2881759
				Start-Sleep -Seconds $PowerShellMaxCmdletsTimePeriodSeconds	

				# Collect report properties.
				$report += $reportProperties

				# Increment to request next page.  
				$getMessageTraceDetailParameters.Page++
			}
		}
Write-Verbose "Collected message trace detail count: $($report.Count)"

# Optionally write report information.
If ( $report ) {
	$report | 
		Export-CSV -Path $outFilePathName -Encoding UTF8 -NoTypeInformation
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

	$messageSubject = "Get Message Trace Detail $reportType for $($outFilePathBase.ExecutionSourceName) on $((Get-Date).ToString('s'))"

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
