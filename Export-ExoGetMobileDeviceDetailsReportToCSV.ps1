<#
    .SYNOPSIS
	Export Exchange Online Protection Mail Traffic Report to a CSV file.  
	
    .DESCRIPTION
	
	.PARAMETER Identity
	The Identity parameter specifies a mailbox. When you specify a value for the Identity parameter, the command looks up the mailbox that is specified in the Identity parameter, connects to the server where the mailbox resides, and returns the statistics for the mailbox. You can use one of the following values: 
	* GUID 
	* Distinguished name (DN)
	* Domain\Account
	* User principal name (UPN)
	* Legacy Exchange DN
	* SmtpAddress
	* Alias
	This parameter accepts pipeline input and wildcards.  
			
			
	.PARAMETER MsExchOrganizationName
	Specify the script's execution environment source.  Must be either a string, 'ComputerName', 'DomainName', 'msExchOrganizationName' or an arbitrary string.  If msExchOrganizationName is requested, but there is no Exchange organization the domain name will be used; if the domain name is requested, but the computer is not a domain member, the computer name is used.  Defaults is msExchOrganizationName.  Used in the case where the Microsoft Exchange Organization name or domain name is too generic (e.g. 'EMAIL', 'CORP' or 'ROOT').  
		
	.PARAMETER OutFileNameTag
	Optional comment string added to the end of the output file name.
	
	.PARAMETER OutFolderPath
	Specify where to write the output file(s).  Supports UNC and relative reference to the current script folder.  The default is .\Reports subfolder.  Except for UNC paths this function will attempt to create and compress the output folder if it doesn’t exist.
	
	.PARAMETER AlertOnly
	When enabled, only unhealthy items are reported and the optional mail subject will contain 'alert' instead of 'report', and if there are no unhealthy items there is no output.  
	
	.PARAMETER MailFrom
	Optionally specify the address from which the mail is sent. Enter a name (optional) and e-mail address, such as 'Name <LocalPart@domain.com>'. 
	
	.PARAMETER MailTo
	Optioanlly specify the addresses to which the mail is sent. Enter names (optional) and the e-mail address, such as 'Name1 <LocalPart1@domain.com>','Name2 <LocalPart2@domain.com>'. 
	
	.PARAMETER MailServer
	Optionally specify the name of the SMTP server that sends the mail message.
	
	.PARAMETER CompressAttachmentLargerThan
	Optionally specify that when a file attachment size is over this limit that it should be compressed when e-mailed.  The default is 5MB.  There is no guarantee the compressed attachment will be below the sender or recipeint's message size limit.  
	
	.EXAMPLE
	...
	
	.EXAMPLE 
	To change the location where the output files are written to a UNC path.
	... -OutFolderPath '\\RemoteServer\C$\Reports'
	
	.EXAMPLE 
	To create output and output file names with a comment appended to the end:
	... -OutFileNameTag 'FirstRun'
	
	.EXAMPLE 
	To send the output or report file via email:
	... -MailFrom 'My Operations<My.Operations@Corp.com>' -MailTo 'MyOpsTeam@corp.com','MyMgmtTeam@corp.com' -MailServer MailHost.Corp.com

	.EXAMPLE 
	Create the output files, and report title and message subject with a custom organization label:
	... -ExecutionSource 'MyCorp'
	
	.EXAMPLE
	To send an alert of unhealty or erroring items instead of a full report:
	... -AlertOnly
	
 	.NOTES
	Author: Terry E Dow
	Last Modified: 2014-08-01
	
#>
[CmdletBinding( 
	SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets. 
)] 
#[System.Diagnostics.DebuggerHidden()]
Param(

	[Parameter( 
		ValueFromPipeline=$TRUE, 
		ValueFromPipelineByPropertyName=$TRUE )]
		[String] $Organization = $NULL,
		
	[Parameter( 
		ValueFromPipeline=$TRUE, 
		ValueFromPipelineByPropertyName=$TRUE )]
		#[DateTime] $StartDate = $NULL,	
		[String] $StartDate = $NULL,	
		
	[Parameter( 
		ValueFromPipeline=$TRUE, 
		ValueFromPipelineByPropertyName=$TRUE )]
		#[DateTime] $EndDate = $NULL,
		[String] $EndDate = $NULL,
	
	[Parameter( 
		ValueFromPipeline=$TRUE, 
		ValueFromPipelineByPropertyName=$TRUE )]
		[String] $Expression = $NULL,

#region Script Exchange Online Header
	
	[Parameter( HelpMessage='Specifies a user name for the credential in User Principal Name (UPN) format, such as "user@domain.com".' )] 
		[String] $CredentialUserName = $NULL,
	
	[Parameter( HelpMessage='Specifies file name where the secure credential password file is located.  The default of null will prompt for the credentials.' )] 
		[String] $CredentialPasswordFileName = $NULL,
		
	[Parameter( HelpMessage='Specify the tenant service as either ''Exchange Online'' or ''Exchange Online Protection''.  The default is Exchange Online.')]
		[ValidateSet( 'Exchange Online', 'Exchange Online Protection' )] 
		[String] $Service = 'Exchange Online',
		
	[Parameter( HelpMessage='The number of seconds to wait after each Exchange Online PowerShell command so not to exceed budgeted throttling policy.')]
		[Int] $PowerShellMaxCmdletsTimePeriodSeconds = 5,
	
#region Script Header
	
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

#Requires -version 2
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
# Include external functions.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

. .\New-OutFilePathBase.ps1
. .\Format-ExpandAllProperties2.ps1
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

# Create a secure string file using current user's credentials.  
# Read-Host -AsSecureString "Securely enter password" | ConvertFrom-SecureString | Out-File -FilePath '.\SecureString.txt'

# read the password securely from a file, or prompt the user.
If ( $CredentialPasswordFileName ) {
	$credentialPassword = Get-Content -Path $CredentialPasswordFileName | 
		ConvertTo-SecureString
	$credential = New-Object System.Management.Automation.PSCredential -ArgumentList $CredentialUserName, $credentialPassword
} Else {
	$credential = Get-Credential -Credential $CredentialUserName
}

Switch ( $Service ) {
	'Exchange Online' {
		# $session = New-PSSession -Credential (Get-Credential -Credential '.onmicrosoft.com' ) -Name ExO -ConnectionUri https://ps.outlook.com/powershell/ -ConfigurationName Microsoft.Exchange -AllowRedirection -Authentication Basic
		$session = New-PSSession -Credential $credential -Name ExO -ConnectionUri https://ps.outlook.com/powershell/ -ConfigurationName Microsoft.Exchange -AllowRedirection -Authentication Basic -ErrorAction Stop
	}

	'Exchange Online Protection' {
		# $session = New-PSSession -Credential (Get-Credential -Credential '.onmicrosoft.com' ) -Name EOP -ConnectionUri https://ps.protection.outlook.com/powershell-liveid -ConfigurationName Microsoft.Exchange -AllowRedirection -Authentication Basic
		$session = New-PSSession -Credential $credential -Name EOP -ConnectionUri https://ps.protection.outlook.com/powershell-liveid -ConfigurationName Microsoft.Exchange -AllowRedirection -Authentication Basic -ErrorAction Stop
	}
}
Trap { Remove-PSSession $session }

$moduleInfo = Import-PSSession $session -AllowClobber
# $moduleInfo.ExportedFunctions.Keys | Where-Object { $_ -Like 'get-*' }

# Try to get tenant's PowerShell throttling budget.  
Try {
	$throttlingPolicy = Get-ThrottlingPolicy
	$PowerShellMaxCmdletsTimePeriodSeconds = $throttlingPolicy.PowerShellMaxCmdletsTimePeriod
} Catch {
	$PowerShellMaxCmdletsTimePeriodSeconds = 5 # Default tenant of 5 seconds
}

#endregion Script Exchange Online Header 

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Collect report information
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

# Create a hash table to splat parameters
$parameters = @{}
If ( $EndDate ) { $parameters.Add( 'EndDate', $EndDate ) } 
If ( $Expression ) { $parameters.Add( 'Expression', $Expression ) }
If ( $Organization ) { $parameters.Add( 'Organization', $Organization ) }
If ( $StartDate ) { $parameters.Add( 'StartDate', $StartDate ) } 

$report = @()

# Continuously query a page at a time until no results.  
$parameters.Page = 1
While ( $properties = Get-MobileDeviceDetailsReport @parameters | Format-ExpandAllProperties ) {
	Write-Verbose "Collecting page: $($parameters.Page)"
	Write-Verbose "Collecting count: $($properties.Count)"
	$report += $properties
	
	# Increment to request next page.  
	$parameters.Page++
	
	# Avoid "Micro delay applied": http://support.microsoft.com/kb/2881759
	Start-Sleep -Seconds $PowerShellMaxCmdletsTimePeriodSeconds
} 

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Optionally write report information.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

If ( $report ) {
	$report | 
		Export-CSV -Path $outFilePathName -NoTypeInformation -WhatIf:$FALSE
}

#region Script Footer

Remove-PSSession $session
# Avoid "Micro delay applied" for any subsequent commands.
Start-Sleep -Seconds $PowerShellMaxCmdletsTimePeriodSeconds

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
	$messageSubject = "Exchange Online Mail Traffic $reportType for $msExchOrganizationName on $((Get-Date).ToString('s'))"
	
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