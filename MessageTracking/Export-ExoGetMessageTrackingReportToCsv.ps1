<#
	.SYNOPSIS
		This cmdlet is available in on-premises Exchange Server 2016 and in the cloud-based service. Some parameters and settings may be exclusive to one environment or the other.

	.DESCRIPTION
		The Get-MessageTrackingReport cmdlet requires you to specify the ID for the message tracking report you want to view. Therefore, first you need to use the Search-MessageTrackingReport cmdlet to find the message tracking report ID for a specific message. You then pass the message tracking report ID from the output of the Search-MessageTrackingReport cmdlet to the Get-MessageTrackingReport cmdlet. For more information, see Search-MessageTrackingReport.

		You need to be assigned permissions before you can run this cmdlet. Although all parameters for this cmdlet are listed in this topic, you may not have access to some parameters if they're not included in the permissions assigned to you. To see what permissions you need, see the "Message tracking" entry in the Mail flow permissions topic.

	.PARAMETER Identity MessageTrackingReportId
		The Identity parameter specifies the ID of the message tracking report ID to retrieve.

		You should run the Search-MessageTrackingReport cmdlet to find the message tracking report ID for the specific message you're tracking, and then pass the value of the MessageTrackingReportID field to this parameter.

	.PARAMETER BypassDelegateChecking SwitchParameter
		The BypassDelegateChecking switch allows Help desk staff and administrators to retrieve message tracking reports for any user. You don't have to specify a value with this switch.

		By default, each user can only see the message tracking reports for messages sent or received by the user. When you use this switch, Exchange allows you to view the message tracking reports for message exchanges among other users.

	.PARAMETER DetailLevel Basic | Verbose
		This parameter is available only in on-premises Exchange 2016.

		The DetailLevel parameter specifies the amount of detail to be displayed for the message tracking report. You can use one of the following values:

		Basic
		Verbose
		If you specify Basic, simple delivery report information is displayed, which is more appropriate for information workers. If you specify Verbose, full report information is displayed, including server names and physical topology information.

	.PARAMETER DomainController Fqdn
		This parameter is available only in on-premises Exchange 2016.

		The DomainController parameter specifies the domain controller that's used by this cmdlet to read data from or write data to Active Directory. You identify the domain controller by its fully qualified domain name (FQDN). For example, dc01.contoso.com.

	.PARAMETER DoNotResolve SwitchParameter
		The DoNotResolve switch prevents the resolution of email addresses to display names. This improves performance, but the end result may not be as easy to interpret because it's missing the display names. You don't have to specify a value with this switch.

	.PARAMETER RecipientPathFilter SmtpAddress
		The RecipientPathFilter parameter specifies the recipient for which the command returns the detailed tracking report.

		Use this parameter when you're using the RecipientPath report template.

	.PARAMETER Recipients String[]
		The Recipients parameter specifies the recipients for whom you want to retrieve the message tracking data.

		You can use this parameter to specify the recipients in the report details if you're using the Summary report template.

	.PARAMETER ReportTemplate Summary | RecipientPath
		The ReportTemplate parameter specifies a predefined format for the output. You can either return a summary for all recipients or a detailed tracking report for one recipient. You can specify one of the following values:
		RecipientPath
		Summary

	.PARAMETER ResultSize Unlimited
		The ResultSize parameter specifies the maximum number of results to return. If you want to return all requests that match the query, use unlimited for the value of this parameter. The default value is 1000.

	.PARAMETER Status Unsuccessful | Pending | Delivered | Transferred | Read
		The Status parameter specifies the delivery status codes you're interested in. You can specify one of the following values:
		Delivered
		Read
		Pending
		Transferred
		Unsuccessful

	.PARAMETER TraceLevel Low | Medium | High
		This parameter is available only in on-premises Exchange 2016.

		The TraceLevel parameter specifies whether additional trace details are included in the output of the message tracking report. This parameter is intended to be used when troubleshooting message tracking issues.

		The acceptable values for the TraceLevel parameter are:
		Low Minimal additional data is returned, including servers that were accessed, timing, message tracking search result counts, and any error information.
		Medium In addition to the data returned for the Low setting, the actual message tracking search results are also returned.
		High Full diagnostic data is returned.

	.EXAMPLE
		This example gets the message tracking report for messages sent from one user to another. This example returns the summary of the message tracking report for a message that David Jones sent to Wendy Richardson.

		$Temp = Search-MessageTrackingReport -Identity "David Jones" -Recipients "wendy@contoso.com"
		Get-MessageTrackingReport -Identity $Temp.MessageTrackingReportID -ReportTemplate Summary

		

	.EXAMPLE
		This example gets the message tracking report for the following scenario: The user Cigdem Akin was expecting an email message from joe@contoso.com that never arrived. She contacted the Help desk, which needs to generate the message tracking report on behalf of Cigdem and doesn't need to see the display names.

		This example searches the message tracking data for the specific message tracking reports, and then returns detailed troubleshooting information for the specific recipient path.
		Search-MessageTrackingReport -Identity "Cigdem Akin" -Sender "joe@contoso.com" -ByPassDelegateChecking -DoNotResolve | ForEach-Object { Get-MessageTrackingReport -Identity $_.MessageTrackingReportID -DetailLevel Verbose -BypassDelegateChecking -DoNotResolve -RecipientPathFilter "cigdem@fabrikam.com" -ReportTemplate RecipientPath }

		
#>
[CmdletBinding(
	SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets.
)]
#[System.Diagnostics.DebuggerHidden()]
Param(

	# Search-MessageTrackingReport parameters

	[Parameter(
	ValueFromPipeline=$TRUE,
	Position=1)]
	[String] $Identity = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Sender = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $BypassDelegateChecking = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $DomainController = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[Switch] $DoNotResolve = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $MessageEntryId = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $MessageId = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Recipients = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $ResultSize = 'Unlimited',

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Subject = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $TraceLevel = $NULL,
	
	# Get-MessageTrackingReport parameters

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#Position=1)]
	#[String] $Identity = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[Switch] $BypassDelegateChecking = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $DetailLevel = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[String] $DomainController = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[Switch] $DoNotResolve = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $RecipientPathFilter = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[String[]] $Recipients = $NULL,

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $ReportTemplate = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[String] $ResultSize = 'Unlimited',

	[Parameter(
	ValueFromPipeline=$TRUE,
	ValueFromPipelineByPropertyName=$TRUE )]
	[String] $Status = $NULL,

	#[Parameter(
	#ValueFromPipeline=$TRUE,
	#ValueFromPipelineByPropertyName=$TRUE )]
	#[String] $TraceLevel = $NULL,

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
$searchMessageTrackingReportParameters = @{}
If ( $Identity ) { $searchMessageTrackingReportParameters.Identity = $Identity }
If ( $Sender ) { $searchMessageTrackingReportParameters.Sender = $Sender }
If ( $BypassDelegateChecking ) { $searchMessageTrackingReportParameters.BypassDelegateChecking = $BypassDelegateChecking }
If ( $DomainController ) { $searchMessageTrackingReportParameters.DomainController = $DomainController }
If ( $DoNotResolve ) { $searchMessageTrackingReportParameters.DoNotResolve = $DoNotResolve }
If ( $MessageEntryId ) { $searchMessageTrackingReportParameters.MessageEntryId = $MessageEntryId }
If ( $MessageId ) { $searchMessageTrackingReportParameters.MessageId = $MessageId }
If ( $Recipients ) { $searchMessageTrackingReportParameters.Recipients = $Recipients }
If ( $ResultSize ) { $searchMessageTrackingReportParameters.ResultSize = $ResultSize }
If ( $Subject ) { $searchMessageTrackingReportParameters.Subject = $Subject }
If ( $TraceLevel ) { $searchMessageTrackingReportParameters.TraceLevel = $TraceLevel }
If ( $Debug ) {
	ForEach ( $key In $searchMessageTrackingReportParameters.Keys ) {
		Write-Debug "`$searchMessageTrackingReportParameters[$key]`:,$($searchMessageTrackingReportParameters[$key])"
	}
}

$getMessageTrackingReportParameters = @{}
#If ( $Identity ) { $getMessageTrackingReportParameters.Identity = $Identity }
If ( $BypassDelegateChecking ) { $getMessageTrackingReportParameters.BypassDelegateChecking = $BypassDelegateChecking }
If ( $DetailLevel ) { $getMessageTrackingReportParameters.DetailLevel = $DetailLevel }
If ( $DomainController ) { $getMessageTrackingReportParameters.DomainController = $DomainController }
If ( $DoNotResolve ) { $getMessageTrackingReportParameters.DoNotResolve = $DoNotResolve }
If ( $RecipientPathFilter ) { $getMessageTrackingReportParameters.RecipientPathFilter = $RecipientPathFilter }
If ( $Recipients ) { $getMessageTrackingReportParameters.Recipients = $Recipients }
If ( $ReportTemplate ) { $getMessageTrackingReportParameters.ReportTemplate = $ReportTemplate }
If ( $ResultSize ) { $getMessageTrackingReportParameters.ResultSize = $ResultSize }
If ( $Status ) { $getMessageTrackingReportParameters.Status = $Status }
If ( $TraceLevel ) { $getMessageTrackingReportParameters.TraceLevel = $TraceLevel }
If ( $Debug ) {
	ForEach ( $key In $getMessageTrackingReportParameters.Keys ) {
		Write-Debug "`$getMessageTrackingReportParameters[$key]`:,$($getMessageTrackingReportParameters[$key])"
	}
}

# Build Report
$report = @()
$report = Search-MessageTrackingReport @searchMessageTrackingReportParameters |
	ForEach-Object {
		$getMessageTrackingReportParameters.Identity = $PSItem.MessageTrackingReportId
		Get-MessageTrackingReport @getMessageTrackingReportParameters | 
			Format-ExpandAllProperties -ConvertByteQuantifiedSizeToBytes -AppendTo $PSItem
	} 
	
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

	$messageSubject = "Get Message Tracking Report $reportType for $($outFilePathBase.ExecutionSourceName) on $((Get-Date).ToString('s'))"

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
