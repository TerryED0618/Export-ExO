<#
    .SYNOPSIS
	Export the ActiveSync statistics of all mailboxes that have successfully synchronizedin the last X days to a CSV file.
	
    .DESCRIPTION
    Export the ActiveSync statistics of all devices of all CAS mailboxes, which have an ActiveSync device partnership, that have successfully synchronized in the last specified days to a CSV file.  
	
	.PARAMETER Identity
	CAS mailbox identity <MailboxIdParameter>.  The default is all ('*').  
	
	.PARAMETER propertiesToLoad
	A comma seperated list of Active Directory properties to be retrieved for each mailbox.  The default is common phone number properties.
	
	.PARAMETER DaysSinceLastSuccessSync
	Only export on ActiveSync devices that have successfully synchronized since this number of days.  The default is 30 days ago.  
	
	.PARAMETER Unique
	Optionally screen out duplicate wireless devices for the same user mailbox.
	NOTE: When the unique option is specified the LastSuccessSync is not guaranteed to be the most recent.  
	
	.PARAMETER MsExchOrganizationName
	Specifiy the script's execution environment source.  Must be either a string, 'ComputerName', 'DomainName', 'msExchOrganizationName' or an arbitrary string.  If msExchOrganizationName is requested, but there is no Exchange organization the domain name will be used; if the domain name is requested, but the computer is not a domain member, the computer name is used.  Defaults is msExchOrganizationName.  Used in the case where the Microsoft Exchange Organization name or domain name is too generic (e.g. 'EMAIL', 'CORP' or 'ROOT').  
		
	.PARAMETER OutFileNameTag
	Optional comment string added to the end of the output file name.
	
	.PARAMETER OutFolderPath
	Specify where to write the output file.  Supports UNC and relative reference to the current script folder.  The default is .\Reports subfolder.  Except for UNC paths this function will attempt to create and compress the log folder if it doesn’t exist.
	
	.PARAMETER AlertOnly
	When enabled, only unhealthy items are reported and the optional mail subject will contain 'alert' instead of 'report', and if there are no unhealthy items no report file is created or mailed.  
	
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
	To change the location where the log files are written to a UNC path.
	... -OutFolderPath '\\RemoteServer\C$\Reports'
	
	.EXAMPLE 
	To create output and log file names with a comment appended to the end:
	... -OutFileNameTag 'FirstRun'
	
	.EXAMPLE 
	To send the output or report file via email:
	... -MailFrom 'My Operations<My.Operations@Corp.com>' -MailTo 'MyOpsTeam@corp.com','MyMgmtTeam@corp.com' -MailServer MailHost.Corp.com

	.EXAMPLE 
	Create the output and log file, and report title and message subject with a custom organization label:
	... -MsExchOrganizationName 'MyCorp'
	
	.EXAMPLE
	To send an alert of unhealty or erroring items only with a mail subject containing 'alert':
	... -AlertOnly
	
 	.NOTES
	Author: Terry E Dow
	Last Modified: 2013-02-25
	
#>
[ CmdletBinding( 
	SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets. 
) ] 
Param(
		
	[Parameter( HelpMessage='CAS mailbox identity <MailboxIdParameter>, wildcarding is supported.' ) ]
		[String] $Identity = '*',
	
	[Parameter( HelpMessage='A comma seperated list of Active Directory properties to be retrieved for each CAS mailbox.' ) ]
		[String[]] $propertiesToLoad = ( 'distinguishedName', 'facsimileTelephoneNumber', 'homePhone', 'ipPhone', 'mobile', 'pager', 'telephoneNumber', 'msExchHomeServerName', 'homeMDB' ),
	
	[Parameter( HelpMessage='Only export on ActiveSync devices that have successfully synchronized since this number of days ago.' ) ]
		[Int] $DaysSinceLastSuccessSync = 30,
	
	[Parameter( HelpMessage='Optionally screen out duplicate wireless devices for the same user mailbox.' ) ]
		[Switch] $Unique,
	
	[Parameter( HelpMessage='Optional organization name used in the output file name, message subject, and any embeded report title.' ) ]
		[String] $MsExchOrganizationName = '',
	
	[Parameter( HelpMessage='Optional string added to the end of the output file name.' ) ]
		[String] $OutFileNameTag = '',
		
	[Parameter( HelpMessage='Specify where to write the output file.' ) ]
		[String] $OutFolderPath = '.\Reports',
	
	[Parameter( HelpMessage='When enabled, only unhealthy items are reported.' ) ]
		[Switch] $AlertOnly = $FALSE,
	
	[Parameter( HelpMessage='Optionally specify the address from which the mail is sent.' ) ]
		[String] $MailFrom = '',
	
	[Parameter( HelpMessage='Optioanlly specify the addresses to which the mail is sent.' ) ]
		[String[]] $MailTo = '',
	
	[Parameter( HelpMessage='Optionally specify the name of the SMTP server that sends the mail message.' ) ]
		[String] $MailServer = '',

	[Parameter( HelpMessage='If the mail message attachment is over this size compress (zip) it.' ) ]
		[Int] $CompressAttachmentLargerThan = 5MB
)

#Requires -version 2
Set-StrictMode -Version Latest

# Detect cmdlet common parameters.  
$Debug = $NULL
If ( -Not $PSBoundParameters.TryGetValue( 'Debug', [Ref] $Debug ) ) {
	$Debug = $FALSE
}
# Replace default -Debug preference from 'Inquire' to 'Continue'.  
If ( $DebugPreference -Eq 'Inquire' ) {
	$DebugPreference = 'Continue' 
}
#$Verbose = $PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Verbose')
$Verbose = $NULL
If ( -Not $PSBoundParameters.TryGetValue( 'Verbose', [Ref] $Verbose ) ) {
	$Verbose = $FALSE
}
#$WhatIf = $PSCmdlet.MyInvocation.BoundParameters.ContainsKey('WhatIf')
$WhatIf = $NULL
If ( -Not $PSBoundParameters.TryGetValue( 'WhatIf', [Ref] $WhatIf ) ) {
	$WhatIf = $FALSE
}

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Collect script execution metrics.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

$scriptStartTime = Get-Date
Write-Verbose "`$scriptStartTime:,$($scriptStartTime.ToString('s'))" 

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Add Exchange Mangement Shell snap-in.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	
##Requires -PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn # Exchange 2013
If ((Get-PSSnapin -Registered) -Match 'Microsoft.Exchange.Management.PowerShell.SnapIn') {
	Write-Debug "Registered:Microsoft.Exchange.Management.PowerShell.SnapIn"
	$registeredExchangeVersion = '2013'
	# If not already added...
	If (-Not ((Get-PSSnapin) -Match 'Microsoft.Exchange.Management.PowerShell.SnapIn')) {
		# ...then add Exchange Management Shell 2013 snap-in to the current console.  
		Write-Debug "Add-PSSnapin:Microsoft.Exchange.Management.PowerShell.SnapIn"
		Add-PSSnapin 'Microsoft.Exchange.Management.PowerShell.SnapIn' -ErrorAction 'Stop'
		. $ENV:ExchangeInstallPath\bin\RemoteExchange.ps1
		Connect-ExchangeServer -Auto
	} 
##Requires -PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 # Exchange 2010
} ElseIf ((Get-PSSnapin -Registered) -Match 'Microsoft.Exchange.Management.PowerShell.E2010') {
	Write-Debug "Registered:Microsoft.Exchange.Management.PowerShell.E2010"
	$registeredExchangeVersion = '2010'
	# If not already added...
	If (-Not ((Get-PSSnapin) -Match 'Microsoft.Exchange.Management.PowerShell.E2010')) {
		# ...then add Exchange Management Shell 2010 snap-in to the current console.  
		Write-Debug "Add-PSSnapin:Microsoft.Exchange.Management.PowerShell.E2010"
		Add-PSSnapin 'Microsoft.Exchange.Management.PowerShell.E2010' -ErrorAction 'Stop'
		. $ENV:ExchangeInstallPath\bin\RemoteExchange.ps1
		Connect-ExchangeServer -Auto
	} 
##Requires -PSSnapin Microsoft.Exchange.Management.PowerShell.Admin # Exchange 2007
} ElseIf ((Get-PSSnapin -Registered) -Match 'Microsoft.Exchange.Management.PowerShell.Admin') {
	Write-Debug "Registered:Microsoft.Exchange.Management.PowerShell.Admin"
	$registeredExchangeVersion = '2007'
	# If not already added...
	If (-Not ((Get-PSSnapin) -Match 'Microsoft.Exchange.Management.PowerShell.Admin')) {
		# ...then add Exchange Management Shell 2007 snap-in to the current console.  
		Write-Debug "Add-PSSnapin:Microsoft.Exchange.Management.PowerShell.Admin"
		Add-PSSnapin 'Microsoft.Exchange.Management.PowerShell.Admin' -ErrorAction 'Stop'
	}
} Else {
	Write-Error 'Either Exchange Management Tools for 2013, 2010 or 2007 is required but not installed.  Exiting.' 
	Exit	
}
Write-Debug "`$registeredExchangeVersion:,$registeredExchangeVersion"

# Change the Exchange Management Shell recipient scope to the entire forest.  Increase chance of stale results due to global catalog replication delays.  
Try {
	# Exchange 2010
 	If ( -Not (Get-ADServerSettings).ViewEntireForest ) {
		Set-ADServerSettings -ViewEntireForest $TRUE
		Write-Debug "`(Get-ADServerSettings).ViewEntireForest:,$((Get-ADServerSettings).ViewEntireForest)"
	}
} Catch {
	Write-Debug "`Get-ADServerSettings:,not available."
	Try {
		# Exchange 2007
		If ( -Not $AdminSessionADSettings.ViewEntireForest ) {
			$AdminSessionADSettings.ViewEntireForest = $TRUE
			Write-Debug "`$AdminSessionADSettings.ViewEntireForest:,$($AdminSessionADSettings.ViewEntireForest)"
		}
	} Catch {
		Write-Debug "`$AdminSessionADSettings.ViewEntireForest:,not available."
	}
	Write-Debug "`Get-ADServerSettings:,not available."
}

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Include external functions.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

. .\New-OutFilePathBase.ps1
. .\Format-ExpandPropertiesToString2.ps1
. .\Get-DSObject2.ps1

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Define internal functions.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Build output and log file path name.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

$outFilePathBase = New-OutFilePathBase  -OutFolderPath $OutFolderPath -ExecutionSource $MsExchOrganizationName -OutFileNameTag $OutFileNameTag

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

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Collect report information
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

# Build a list of domains within this forest.
$domains = @( ([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()).Domains | ForEach-Object { $_.Name } )
Write-Debug "`$domains:,$domains"

[DateTime] $DateSinceLastSuccessSync = $(Get-Date).AddDays(-$DaysSinceLastSuccessSync)
Write-Debug "`$DateSinceLastSuccessSync: $($DateSinceLastSuccessSync.ToString('s'))"

$report = @()

# Exchange 2007 Management Shell's Get-CASMailbox command does not have a parameter set that includes both -Filter and -Identity, they are mutually exclusive.
# Using the -Filter argument is more efficient than piping through Where-Object.  
Switch ( $registeredExchangeVersion ) {
	'2013' {
		$command = { Get-CASMailbox -Identity $Identity -Filter { HasActiveSyncDevicePartnership -Eq $TRUE -And ( Name -NotLike "CAS_{*") -And ( Name -NotLike "DiscoverySearchMailbox {*") } -ResultSize Unlimited }
	}
	'2010' {
		$command = { Get-CASMailbox -Identity $Identity -Filter { HasActiveSyncDevicePartnership -Eq $TRUE -And ( Name -NotLike "CAS_{*") -And ( Name -NotLike "DiscoverySearchMailbox {*") } -ResultSize Unlimited }
	}
	'2007' {
		If ( $Identity -Eq '*' ) {
			$command = { Get-CASMailbox -Filter { HasActiveSyncDevicePartnership -Eq $TRUE -And ( Name -NotLike "CAS_{*") -And ( Name -NotLike "DiscoverySearchMailbox {*") } -ResultSize Unlimited }
		} Else {
			$command = { Get-CASMailbox -Identity $Identity -ResultSize Unlimited | 
							Where-Object { $_.HasActiveSyncDevicePartnership -Eq $TRUE -And ( $_.Name -NotLike "CAS_{*") -And ( $_.Name -NotLike "DiscoverySearchMailbox {*") } }
		}
	}
}

Invoke-Command -ScriptBlock $command |
	ForEach {
		$cASMailbox = $_ |
			Format-ExpandPropertiesToString
		Write-Debug "`$cASMailbox:,$cASMailbox"
		Write-Verbose "`$cASMailbox.Identity:,$($cASMailbox.Identity)"
		
		# Extract the CAS mailbox's distinguished name domain component.
		If ( $cASMailbox.DistinguishedName -Match ',?(?<DomainComponent>DC=.*)$' ) {
			$cASMailboxDomain = $Matches['DomainComponent']
		} Else {
			$cASMailboxDomain = ''
		}
		Write-Debug "`$cASMailboxDomain:,$cASMailboxDomain"
				
		$user = Get-DSObject -Server $cASMailboxDomain -LDAPFilter "(&(objectCategory=person)(objectClass=user)(proxyAddresses=SMTP:$($cASMailbox.PrimarySmtpAddress)))" -Properties $propertiesToLoad |
			Format-ExpandPropertiesToString -AppendTo $cASMailbox
		Write-Debug "`$user:,$user"
	
		Write-Debug "`$cASMailbox.Identity:,$($cASMailbox.Identity)"
		Get-ActiveSyncDeviceStatistics -Mailbox $cASMailbox.Identity |
			Where-Object { $DateSinceLastSuccessSync -LE $_.LastSuccessSync } |
			ForEach {
				
				$activeSyncDeviceStatistics = $_ |
					Format-ExpandPropertiesToString -AppendTo $user
				Write-Debug "`$activeSyncDeviceStatistics:,$activeSyncDeviceStatistics"
				
				$report += $activeSyncDeviceStatistics
				
			} 
	} 
		
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Optionally write report information.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

If ( $report ) {

	If ( $Unique ) {
			$report |
				Sort-Object -Property PrimarySmtpAddress -Unique | 
				Export-Csv -noTypeInformation -Path $outFilePathName -WhatIf:$FALSE
	} Else {	
		$report |
			Sort-Object -Property PrimarySmtpAddress,LastSuccessSync  | 
			Export-Csv -noTypeInformation -Path $outFilePathName -WhatIf:$FALSE
	} 
	
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	# Optionally mail report.
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	
	If ( $MailFrom -And $MailTo -And $MailServer ) {

		# Determine subject line report/alert mode.  
		If ( $AlertOnly ) {
			$reportType = 'Alert'
		} Else {
			$reportType = 'Report'
		}
		$messageSubject = "Exchange ActiveSync devices that have successfully synchronized in past $DaysSinceLastSuccessSync days $reportType for $($outFilePathBase.ExecutionSourceName) on $($outFilePathBase.DateTimeStamp)" 
		
		# If the out file is large then a specified limit (message size limit), then create a compressed (zipped) copy.  
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
}

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
# Write script execution metrics.
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

$scriptEndTime = Get-Date
Write-Verbose "`$scriptEndTime:,$($scriptEndTime.ToString('s'))" 
$scriptElapsedTime =  $scriptEndTime - $scriptStartTime
Write-Verbose "`$scriptElapsedTime:,$scriptElapsedTime"
If ( $Verbose -Or $Debug ) {
	Stop-Transcript
}