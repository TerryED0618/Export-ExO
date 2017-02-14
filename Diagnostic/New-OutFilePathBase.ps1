<#
	.SYNOPSIS
	Build a output folder path and file name without extension.  
	  
	.DESCRIPTION
	Build a output folder path and file name without extension.  The output file name is in the form of "YYYYMMDDTHHMMSSZZZ-<ExecutionSourceName>-<CallingScriptName>[-<OutFileNameTag>]".  
	The date format is sortable date/time stamp in ISO-8601:2004 basic format with no invalid file name characters (such as colon ':').  
	The ExecutionSourceName is either the forest, Exchange orgainzation, domain or computer name of the current server.  
	The CallingScriptName is the name of the script file that is calling this function.  
	OutFileNameTag is an optional comment added to the output file name.  

	.COMPONENT
	System.DirectoryServices.ActiveDirectory
	System.IO.Path
	WMI Win32_ComputerSystem
	WMI Win32_Directory

	.PARAMETER DateOffsetDays
	Optionally specify the number of days added or subtracted from the current date.  Default is 0 days.  
	
	.PARAMETER ExecutionSource
	Specifiy the script's execution environment source.  Must be either; 'ComputerName', 'DomainName', 'ForestName', 'msExchOrganizationName' or an arbitrary string.  If msExchOrganizationName is requested, but there is no Exchange organization the domain name will be used; If ForestName is requested, but there is no forest the domain name will be used; if the domain name is requested, but the computer is not a domain member, the computer name is used.  Defaults is msExchOrganizationName.  An arbitrary string can be used in the case where the Microsoft Exchange Organization name, forest name or domain name is too generic (e.g. 'EMAIL', 'CORP' or 'ROOT').  

	.PARAMETER FileNameComponentDelimiter
	Optional file name component delimiter.  Default is hyphen '-'.  
	
	.PARAMETER InvalidFilePathCharsSubstitute  
	Optionally specify which character to use to replace invalid folder and file name characters.  Default is underscore '_'.  The substitute character cannot itself be an folder or file name invalid character.  

	.PARAMETER OutFileNameTag 
	Optional comment string added to the end of the output file name.  

	.PARAMETER OutFolderPath 
	Specify where to write the output file.  Supports UNC and relative reference to the current script folder.  The default is .\Reports subfolder.  Except for UNC paths this function will attempt to create and compress the output folder if it doesn’t exist.
	
	.OUTPUTS
	An string with six custom properties:
		A string containing the ouput file path name which contains a full folder path and file name without extension.  If the folder path does exist and is not a UNC path an attempt is made to create the folder and mark it as compressed.  
		FolderPath: Full output folder path name.
		DateTime: The date/time used to create the DateTimeStamp string.
		DateTimeStamp: The date/time stamp used in the output file name.  The sortable ISO-8601:2004 basic format includes the time zone offset from the executing server.  
		ExecutionSourceName: The execution environmental source provided or retrieved. 
		ScriptFileName: Calling script file name.
		FileName: Output file name without extension.
	
	
	.EXAMPLE
	To change the location where the output files are written to an relative path use the -OutFolderPath argument.  
	To add a comment to the file name use the -OutFileNameTag argument.  
	
	$outFilePathBase = New-OutFilePathBase -OutFolderPath '.\Logs' -OutFileNameTag 'TestRun#7'
	$outFilePathName = "$outFilePathBase.csv"
	$logFilePathName = "$outFilePathBase.log"
	
	$outFilePathName
	<CurrentLocation>\Logs\19991231T235959-0600-<MyExchangeOrgName>-<CallingScriptName>-TestRun#7.csv
	
	$logFilePathName
	<CurrentLocation>\Logs\19991231T235959-0600-<MyExchangeOrgName>-<CallingScriptName>-TestRun#7.log
	
	$outFilePathBase.FolderPath
	<CurrentLocation>\Logs\
	
	$outFilePathBase.DateTimeStamp
	19991231T235959-0600
	
	$outFilePathBase.ExecutionSourceName
	<MyExchangeOrgName>
	
	$outFilePathBase.ScriptFileName
	<CallingScriptName>
	
	$outFilePathBase.FileName
	19991231T235959-0600-<MyExchangeOrgName>-<CallingScriptName>-TestRun#7
	
	
	.EXAMPLE
	To change the location where the output files are written to an absolute path use the -OutFolderPath argument.  
	To change the exection environment source to the domain name use the -ExecutionSource argument.  
	
	$outFilePathBase = New-OutFilePathBase -ExecutionSource ForestName -OutFolderPath C:\Reports\
	
	$outFilePathBase
	C:\Reports\19991231T235959-0600-<MyForestName>-<CallingScriptName>
	
	$outFilePathBase.FolderPath
	C:\Reports\
	
	$outFilePathBase.ExecutionSourceName
	<MyForestName>
	
	$outFilePathBase.FileName
	19991231T235959-0600-<MyForestName>-<CallingScriptName>
	
	
	.EXAMPLE
	To change the location where the output files are written to an absolute path use the -OutFolderPath argument.  
	To change the exection environment source to the domain name use the -ExecutionSource argument.  
	
	$outFilePathBase = New-OutFilePathBase -ExecutionSource DomainName -OutFolderPath C:\Reports\
	
	$outFilePathBase
	C:\Reports\19991231T235959-0600-<MyDomainName>-<CallingScriptName>
	
	$outFilePathBase.FolderPath
	C:\Reports\
	
	$outFilePathBase.ExecutionSourceName
	<MyDomainName>
	
	$outFilePathBase.FileName
	19991231T235959-0600-<MyDomainName>-<CallingScriptName>
	
	
	.EXAMPLE
	To change the location where the output files are written to a UNC path use the -OutFolderPath argument.  
	To change the exection environment source to the computer name use the -ExecutionSource argument.  
	
	$outFilePathBase = New-OutFilePathBase -ExecutionSource ComputerName -OutFolderPath \\Server1\C$\Reports\
	
	$outFilePathBase
	\\Server1\C$\Reports\19991231T235959-0600-<MyComputerName>-<CallingScriptName>
	
	$outFilePathBase.FolderPath
	\\Server1\C$\Reports\
	
	$outFilePathBase.ExecutionSourceName
	<MyComputerName>
	
	$outFilePathBase.FileName
	19991231T235959-0600-<MyComputerName>-<CallingScriptName>
	
	
	.EXAMPLE
	To change the exection environment source to an arbitrary string use the -ExecutionSource argument.  
	
	$outFilePathBase = New-OutFilePathBase -ExecutionSource 'MyOrganization' 
	
	$outFilePathBase
	<CurrentLocation>\Reports\19991231T235959-0600-MyOrganization-<CallingScriptName>
		
	$outFilePathBase.ExecutionSourceName
	MyOrganization
	
	$outFilePathBase.FileName
	19991231T235959-0600-MyOrganization-<CallingScriptName>
	
	
	.EXAMPLE
	To change the date/time stamp to the yeterday's date, as when collecting information from yesterday's data use the -DateOffsetDays argument.  
	
	$outFilePathBase = New-OutFilePathBase -DateOffsetDays -1 
	
	$outFilePathBase
	<CurrentLocation>\Reports\<yesterday's date>T235959-0600-<MyExchangeOrgName>-<CallingScriptName>
	
	$outFilePathBase.DateTimeStamp
	<yesterday's date>T235959-0600
	
	$outFilePathBase.FileName
	<yesterday's date>T235959-0600-<MyExchangeOrgName>-<CallingScriptName>
	
	
	.EXAMPLE
	To change which charater is used to join the file name components together use the -FileNameComponentDelimiter argument.  Note the date/time stamp time zone offset component is prefixed with a plus '+' or minus '-' and is not affected by the argument.  
	
	$outFilePathBase = New-LogFilePathBase -FileNameComponentDelimiter '_' 
	
	$outFilePathBase
	<CurrentLocation>\Reports\19991231T235959T235959-0600_<MyExchangeOrgName>_<CallingScriptName>
	
	$outFilePathBase.FileName
	19991231T235959-0600_<MyExchangeOrgName>_<CallingScriptName>
	
	
	.EXAMPLE
	To change the character used to replace invalid folder and file name characters use the -InvalidFilePathCharsSubstitute argument.  
	
	$outFilePathBase = New-LogFilePathBase -InvalidFilePathCharsSubstitute '#' -LogFileNameTag 'From:LocalPart@domain.com'
	
	$outFilePathBase
	<CurrentLocation>\Reports\19991231T235959-0600-<MyExchangeOrgName>-<CallingScriptName>-From#LocalPart@domain.com

	$outFilePathBase.FileName
	19991231T235959-0600-<MyExchangeOrgName>-<CallingScriptName>-From#LocalPart@domain.com
	
	
	.NOTES
	Author: Terry E Dow
	2013-09-12 Terry E Dow - Added support for ExecutionSource of ForestName.  
	2013-09-21 Terry E Dow - Peer reviewed with the North Texas PC User Group PowerShell SIG and specific suggestion by Josh Miller.
	2013-09-21 Terry E Dow - Changed output from PSObject to String.  No longer require referencing returned object's ".Value" property.
	Last Modified: 2013-09-21

	.LINK
#>
Function New-OutFilePathBase {
	[ CmdletBinding() ]
	Param(
		[ Parameter( HelpMessage='Specify a folder path or UNC where the output file is written.' ) ]
			[String] $OutFolderPath = '.\Reports',
		
		[ Parameter( HelpMessage='Optional name representing the name of the organization this script is running under, used in the output file name, message subject, and embeded report title.  Supported values: ForestName, msExchOrganizationName, DomainName, ComputerName, or any other arbitrary string.' ) ]
			[String] $ExecutionSource = 'msExchOrganizationName',
		
		[ Parameter( HelpMessage='Optional string added to the end of the output file name.' ) ]
			[String] $OutFileNameTag = '',
		
		[ Parameter( HelpMessage='Optionally specify the number of days added or subtracted from the current date.' ) ]
			[Int] $DateOffsetDays = 0,
		
		[ Parameter( HelpMessage='Optional file name component delimiter.  The specified string cannot be an invalid file name character.' ) ]
		[ ValidateScript( { [System.IO.Path]::GetInvalidFileNameChars() -NotContains $_ } ) ] 
			[String] $FileNameComponentDelimiter = '-',
		
		[ Parameter( HelpMessage='Optionally specify which character to use to replace invalid folder and file name characters.  The specified string cannot be an invalid folder or file name character.' ) ]
		[ ValidateScript( { [System.IO.Path]::GetInvalidPathChars() -NotContains $_ -And [System.IO.Path]::GetInvalidFileNameChars() -NotContains $_ } ) ] 
			[String] $InvalidFilePathCharsSubstitute = '_'
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
	# Declare internal functions.
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

	Function Get-ComputerName {
		Write-Output (Get-WmiObject -Class Win32_ComputerSystem).Name
	}
	
	Function Get-DomainName {
		Write-Output ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().GetDirectoryEntry()).Name
	}

	Function Get-ForestName {
		Write-Output ([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()).Name
	}
	
	Function Get-MsExchOrganizationName {
		$currentForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
		$rootDomainDN = $currentForest.RootDomain.GetDirectoryEntry().DistinguishedName
		$msExchConfigurationContainerSearcher = New-Object DirectoryServices.DirectorySearcher 
		$msExchConfigurationContainerSearcher.SearchRoot = "LDAP://CN=Microsoft Exchange,CN=Services,CN=Configuration,$rootDomainDN"
		$msExchConfigurationContainerSearcher.Filter = '(objectCategory=msExchOrganizationContainer)'
		$msExchConfigurationContainerResult = $msExchConfigurationContainerSearcher.FindOne()
		Write-Output $msExchConfigurationContainerResult.Properties.Item('Name')
	}
		
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	# Build output folder path.
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	
	# Replace invalid folder characters: "<>| and others.  
	$OutFolderPath = [RegEx]::Replace( $OutFolderPath, "[$([System.IO.Path]::GetInvalidPathChars())]", $InvalidFilePathCharsSubstitute )
	Write-Debug "`$OutFolderPath:,$OutFolderPath"
	
	# Get the current path.  If invoked from a script...
	Write-Debug "`$script:MyInvocation.InvocationName:,$($script:MyInvocation.InvocationName)"
	If ( $script:MyInvocation.InvocationName ) {
		# ...get the parent script's command path.
		$currentPath = Split-Path $script:MyInvocation.MyCommand.Path -Parent
	} Else {
		# ...else get the current location.
		$currentPath = (Get-Location).Path
	}
	Write-Debug "`$currentPath:,$currentPath"
	
	# Get the full path of the combined folders of the current path and the specified output folder, which may be a relative path.
	$OutFolderPath = [System.IO.Path]::GetFullPath( [System.IO.Path]::Combine( $currentPath, $OutFolderPath ) )

	# Verify Output folder path name has trailing directory separator character.
	If ( -Not $OutFolderPath.EndsWith( [System.IO.Path]::DirectorySeparatorChar ) ) { 
		$OutFolderPath += [System.IO.Path]::DirectorySeparatorChar
	}
	Write-Debug "`$OutFolderPath:,$OutFolderPath"

	# If the output folder does not exist and not a UNC path, try to create and set it to compressed. 
	If ( -Not ((Test-Path $OutFolderPath -PathType Container) -Or ($OutFolderPath -Match '^\\\\[^\\]+\\')) ) { 
		[VOID] (New-Item -Path $OutFolderPath -ItemType Directory -WhatIf:$FALSE) #<<<< ToDo: support recursive directory creation.
		[VOID] (Get-WmiObject -Class 'Win32_Directory' -Filter "Name='$($OutFolderPath.Replace('\','\\').TrimEnd('\'))'" -ErrorAction SilentlyContinue).Compress() 
	} 

	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	# Build file name components.  
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

	# Get sortable date/time stamp in ISO-8601:2004 basic format "YYYYMMDDTHHMMSSZZZ" with no invalid file name characters.
	$dateTime = (Get-Date).AddDays($DateOffsetDays)
	$dateTimeStamp = [RegEx]::Replace( $dateTime.ToString('yyyyMMdd\THHmmsszzz'), "[$([System.IO.Path]::GetInvalidFileNameChars())]", '' )
	Write-Debug "`$dateTimeStamp:,$dateTimeStamp"
	
	# Get execution environment source name.
	Switch ( $ExecutionSource ) {
		
		'ComputerName' {
			# Get current computer name.
			$executionSourceName =  Get-ComputerName
			Break
		}
		
		'DomainName' {
			# Try to get current domain name, else get computer name.
			Try {
				$executionSourceName = Get-DomainName
			} Catch {
				$executionSourceName =  Get-ComputerName
			}
			Break
		}
		
		'ForestName' { 
			# Try to get current forest name, else get domain or computer name.  
			Try {
				$executionSourceName = Get-ForestName
			} Catch {
			
				# Try to get current domain name, else get computer name.
				Try {
					$executionSourceName = Get-DomainName
				} Catch {
					$executionSourceName =  Get-ComputerName
				}
				
			}
			Break
		}

		{ -Not $_ -Or $_ -Eq 'msExchOrganizationName' } { # If null or 'msExchOrganizationName'
			# Try to get current forest's Exchange organization name, else get domain or computer name.  
			Try {
				$executionSourceName = Get-MsExchOrganizationName
			} Catch {
			
				# Try to get current domain name, else get computer name.
				Try {
					$executionSourceName = Get-DomainName
				} Catch {
					$executionSourceName =  Get-ComputerName
				}
				
			}
			Break
		}
		
		Default {
			$executionSourceName = $ExecutionSource
		}
	}
	Write-Debug "`$executionSourceName:,$executionSourceName"
		
	# Get current script name.
	$myScriptFileName = [System.IO.Path]::GetFileNameWithoutExtension( $MyInvocation.ScriptName )
	Write-Debug "`$myScriptFileName:,$myScriptFileName"
	
	Write-Debug "`$OutFileNameTag:,$OutFileNameTag"

	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	# Build file path name without extension.  
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

	# Join non-null file name components with delimiter.
	$outFileName =  $( ( $dateTimeStamp, $executionSourceName, $myScriptFileName, $OutFileNameTag ) | Where-Object { $_ } ) -Join $FileNameComponentDelimiter
	# Replace invalid file name characters: "*/:<>?[\]|
	$outFileName = [RegEx]::Replace( $outFileName, "[$([System.IO.Path]::GetInvalidFileNameChars())]", $InvalidFilePathCharsSubstitute )
	Write-Debug "`$outFileName:,$outFileName"
	
	# Join folder path and file name and other information derived from this function.    
	$outFilePathBase = New-Object PSObject
	#$outFilePathBase = "$OutFolderPath$outFileName"
	Add-Member -InputObject $outFilePathBase -MemberType NoteProperty -Name 'Value' -Value "$OutFolderPath$outFileName"
	Add-Member -InputObject $outFilePathBase -MemberType NoteProperty -Name 'FolderPath' -Value $OutFolderPath
	Add-Member -InputObject $outFilePathBase -MemberType NoteProperty -Name 'FileName' -Value $outFileName
	Add-Member -InputObject $outFilePathBase -MemberType NoteProperty -Name 'DateTimeStamp' -Value $dateTimeStamp
	Add-Member -InputObject $outFilePathBase -MemberType NoteProperty -Name 'DateTime' -Value $dateTime
	Add-Member -InputObject $outFilePathBase -MemberType NoteProperty -Name 'ExecutionSourceName' -Value $ExecutionSourceName
	Add-Member -InputObject $outFilePathBase -MemberType NoteProperty -Name 'ScriptFileName' -Value $myScriptFileName
	
	Write-Output $outFilePathBase
}