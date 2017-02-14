<#
	.Synopsis
	Connect to an Lightweight Directory Access Protocol (LDAP) server and return a directory searcher object.  
	  
	.Description
	Connects to an LDAP server and return a directory searcher object. 

	.Component
	System.DirectoryServices

	.Parameter Server
	Specifies the Active Directory Domain Services instance to connect to, by providing one of the following values for a corresponding domain name or directory server. The service may be any of the following:  Active Directory Lightweight Domain Services, Active Directory Domain Services or Active Directory Snapshot instance.
	Domain name values:
	  Fully qualified domain name
		Examples: corp.contoso.com
	  NetBIOS name
		Example: CORP
		
	.Parameter SearchBase
	Specifies an Active Directory path to search under.

	The default value of this parameter is the default naming context of the target domain.

	The following example shows how to set this parameter to search under an OU.
	  -SearchBase "ou=mfg,dc=noam,dc=corp,dc=contoso,dc=com"

	When the value of the SearchBase parameter is set to an empty string and you are connected to a GC port, all partitions will be searched. If the value of the SearchBase parameter is set to an empty string and you are not connected to a GC port, an error will be thrown.
	The following example shows how to set this parameter to an empty string.   -SearchBase ""
	
	.Outputs
	Returns an object of type [System.DirectoryServices.DirectorySearcher].

	.Notes
	Author: Terry E Dow
	Last Modified: 2014-02-22

	.Link
#>
Function Connect-DirectorySearcher{
	[ CmdletBinding( 
		SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets. 
	) ] 
	Param( 
		[ Parameter( HelpMessage='Optional server name to bind to.' ) ] 
			[String] $Server = '',
		
		[ Parameter( HelpMessage='Optional root of the LDAP search in distinguished name form.' ) ] 		
			[String] $SearchBase = '',
			
		[ Parameter( HelpMessage='Optional search the entire forest instead of a single domain.' ) ] 		
			[Switch] $ViewEntireForest = $FALSE
	)

#region Script Header
	
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

#endregion Script Header
	
	Write-Debug "`$Server:,$Server"
	
	# Validate search root distinguished name.  If not specified, get domain controller's Root Directory Server agent Entry default naming context in DN form.  
	If ( -Not $SearchBase ) { 
		If ( $ViewEntireForest) {
			$Server = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().FindGlobalCatalog().Name				
			$SearchBase =  "dc=$([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().RootDomain.Name.Split('.') -Join ',dc=')"
		} Else {
			$SearchBase = ([System.DirectoryServices.DirectoryEntry] "LDAP://$Server`RootDSE").DefaultNamingContext
		}
	}
	Write-Debug "`$SearchBase:,$SearchBase"

	# Combine the non-null server and search base with a slash.  
	$serverSearchBase = ( ( $Server, $SearchBase ) | Where-Object { $_ } ) -Join '/'
	Write-Debug "`$serverSearchBase:,$serverSearchBase"

	# Bind to a LDAP server. "LDAP://HostName[:PortNumber][/DistinguishedName]" LDAP ADsPath (Windows) http://msdn.microsoft.com/en-us/library/windows/desktop/aa746384(v=vs.85).aspx  http://www.ietf.org/rfc/rfc2255.txt
	If ( $ViewEntireForest) {
		Write-Debug "`$directorySearcher:,GC://$serverSearchBase"
		$directorySearcher = New-Object DirectoryServices.DirectorySearcher( [System.DirectoryServices.DirectoryEntry] "GC://$serverSearchBase" )
	} Else {
		Write-Debug "`$directorySearcher:,LDAP://$serverSearchBase"
		$directorySearcher = New-Object DirectoryServices.DirectorySearcher( [System.DirectoryServices.DirectoryEntry] "LDAP://$serverSearchBase" )
	}
	Write-Debug "`$directorySearcher.SearchRoot.DistinguishedName:,$($directorySearcher.SearchRoot.DistinguishedName)"
	
	# Set CacheResults to not cache the search results.  
	# http://msdn.microsoft.com/en-us/library/system.directoryservices.directorysearcher.cacheresults.aspx
	$directorySearcher.CacheResults = $FALSE
	      
	# Set PageSize to specified domain's MaxPageSize.
	$maxPageSize = $NULL
	$configurationNamingContext = ([System.DirectoryServices.DirectoryEntry] "LDAP://$Server`RootDSE").ConfigurationNamingContext
	$lDAPAdminLimits = ([System.DirectoryServices.DirectoryEntry] "LDAP://CN=Default Query Policy,CN=Query-Policies,cn=Directory Service,cn=Windows NT,CN=Services,$configurationNamingContext").lDAPAdminLimits
	If ( $lDAPAdminLimits ) {
		$maxPageSize = [Int] ($lDAPAdminLimits -Match '^MaxPageSize=(?<MaxPageSize>\d*)$')[0].Split('=')[1]
		$directorySearcher.PageSize = $maxPageSize  
	} Else {
		$directorySearcher.PageSize = 1000
	}
	Write-Debug "`$directorySearcher.PageSize:,$($directorySearcher.PageSize)"
		
	Write-Output $directorySearcher	
}

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

<#
	.Synopsis
	Gets one or more Directory Services objects.  
	  
	.Description
	The Get-DSObject cmdlet gets an Directory Service object or performs a search to retrieve multiple objects.

	This function mimics in part the Remote Active Directory module's Get-ADObject command and a subset of its command arguments, and can be used in its place on downlevel systems prior to Windows 2008 R2 and Windows 7 (less than NT 6.1).  Relevant portions of the Get-ADObject help has been added to this help section.  
	
	To search for and retrieve more than one object, use the LDAPFilter parameter. If you have existing LDAP query strings, you can use the LDAPFilter parameter.

	This cmdlet gets a default set of Directory Service object properties. To get additional properties use the Properties parameter. For more information about the how to determine the properties for computer objects, see the Properties parameter description.

	.Component
	System.DirectoryServices

	.Parameter AugmentProperties
	For supported Active Directory properties, add one or more additional augmented properties.  For example if the following is used -Properties userAccountControl -AugmentProperties then an additional property named userAccountControlDescription is added; userAccountControl will still show a number, userAccountControlDescription will show a text description.  
	
		accountExpires
		+ accountExpiresDate
		+ accountExpiresDaysUntil	
		 
		badPasswordTime
		+ badPasswordTimeDate
		+ badPasswordTimeDaysSince
		
		distinguishedName
		+ OrganizationalUnit
		
		groupType
		+ groupTypeDescription
		
		lastLogoff
		+ lastLogoffDate
		+ lastLogoffDaysSince
		
		lastLogon
		+ lastLogonDate
		+ lastLogonDaysSince

		lastLogonTimestamp
		+ lastLogonTimestampDate
		+ lastLogonTimestampDaysSince
		
		lockoutTime
		+ lockoutTimeDate
		+ lockoutTimeDaysSince
		
		pwdLastSet
		+ pwdLastSetDate
		+ pwdLastSetDaysSince
		
		userAccountControl
		+ userAccountControlDescription
		
	.Parameter directorySearcher
	Optional connected [System.DirectoryServices.DirectorySearcher] object.  If not supplied, one will be estabished.  Creating one out of this function can reduce redundent connections for multiple queries to the same domain and search root.  Consider using Connect-DirectorySearcher for this purpose.  

	.Parameter LDAPFilter
	Specifies an LDAP query string that is used to filter Active Directory objects. You can use this parameter to run your existing LDAP queries. The Filter parameter syntax supports the same functionality as the LDAP syntax. For more information, see the Filter parameter description and the about_ActiveDirectory_Filter. The following example shows how to set this parameter to search for all objects in the organizational unit specified by the SearchBase parameter with a name beginning with "sara". 

	  -LDAPFilter "(name=sara*)"  -SearchScope Subtree -SearchBase "DC=NA,DC=fabrikam,DC=com"

	.Parameter Properties
	Specifies the properties of the output object to retrieve from the server. Use this parameter to retrieve properties that are not included in the default set.

	Specify properties for this parameter as a comma-separated list of names. To display all of the attributes that are set on the object, specify * (asterisk).
	
	.Parameter SearchBase
	Specifies an Active Directory path to search under.

	The default value of this parameter is the default naming context of the target domain.

	The following example shows how to set this parameter to search under an OU.
	  -SearchBase "ou=mfg,dc=noam,dc=corp,dc=contoso,dc=com"

	When the value of the SearchBase parameter is set to an empty string and you are connected to a GC port, all partitions will be searched. If the value of the SearchBase parameter is set to an empty string and you are not connected to a GC port, an error will be thrown.
	The following example shows how to set this parameter to an empty string.   -SearchBase ""

	.Parameter SearchScope
	Specifies the scope of an Active Directory search. Possible values for this parameter are:
	Base or 0
	OneLevel or 1
	Subtree or 2

	A Base query searches only the current path or object. A OneLevel query searches the immediate children of that path or object. A Subtree query searches the current path or object and all children of that path or object.

	The default is Subtree.
	
	The following example shows how to set this parameter to a subtree search.
	  -SearchScope Subtree
	The following lists the acceptable values for this parameter:
	  Base
	  OneLevel
	  Subtree
	
	.Parameter Server
	Specifies the Active Directory Domain Services instance to connect to, by providing one of the following values for a corresponding domain name or directory server. The service may be any of the following:  Active Directory Lightweight Domain Services, Active Directory Domain Services or Active Directory Snapshot instance.
	Domain name values:
	  Fully qualified domain name
		Examples: corp.contoso.com
	  NetBIOS name
		Example: CORP
	
	.Outputs
	Returns an object of type [System.DirectoryServices.DirectoryEntry].   

	.Example
	$results = Get-DSObject -LDAPFilter "(&(objectCategory=person)(objectClass=user)(mailNickname=*)(|(homeMDB=*)(msExchHomeServerName=*)))" -Properties 'mail','proxyAddresses','targetAddress'

	.Example
	$directorySearcher = Initialize-DSObject

	$mailboxes = Get-DSObject -directorySearcher $directorySearcher -LDAPFilter "(&(objectCategory=person)(objectClass=user)(mailNickname=*)(|(homeMDB=*)(msExchHomeServerName=*)))" -Properties 'mail','proxyAddresses','targetAddress'
	$groups = Get-DSObject -directorySearcher $directorySearcher -LDAPFilter "(&(objectCategory=group)(mailNickname=*))" -Properties 'proxyAddresses'

	.Notes
	Author: Terry E Dow
	Last Modified: 2014-02-22

	.Link
#>
Function Get-DSObject{       
	[ CmdletBinding( 
		SupportsShouldProcess = $TRUE # Enable support for -WhatIf by invoked destructive cmdlets. 
	) ] 
	Param( 
		[ Parameter( HelpMessage='Optional connected [System.DirectoryServices.DirectorySearcher] object.' ) ] 
			[DirectoryServices.DirectorySearcher] $DirectorySearcher = $NULL,
		
		[ Parameter( HelpMessage='Optional server name to bind to.' ) ] 
			[String] $Server = '',
		
		[ Parameter( HelpMessage='Optional root of the LDAP search in distinguished name form.' ) ] 
			[String] $SearchBase = '',
		
		[ Parameter( HelpMessage='Search Scope (Base/OneLevel/Subtree).  Default is Subtree.' ) ] 
		[ValidateSet( 'Base', 'OneLevel', 'Subtree' )]
			[String] $SearchScope = 'Subtree',
		
		[ Parameter( HelpMessage='LDAP search filter.  Default is (objectCategory=*).' ) ] 
			[String] $LDAPFilter = '(objectCategory=*)', 
		
		[ Parameter( HelpMessage='List of attributes to be returned with the directory services object(s) that matched the query.  Default is distinguishedName.' ) ] 
			[String[]] $Properties = 'distinguishedName',
						
		[ Parameter( HelpMessage='Optional add additional augmented properties.  Default is $FALSE.' ) ] 		
			[Switch] $AugmentProperties = $FALSE,
			
		[ Parameter( HelpMessage='Optional search the entire foresst instead of a single domain.  Default is $FALSE.' ) ] 		
			[Switch] $ViewEntireForest = $FALSE,
		
		[ Parameter( HelpMessage='The string used to join array properties values together. Default is semicolon (;).' ) ] 
			[String] $IntraValueDelimiter = ';'
	)

#region Script Header
	
	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
	
	Begin {
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
		
		# Define enumerations.
		
		# GROUP_TYPE_* Group-Type attribute http://msdn.microsoft.com/en-us/library/windows/desktop/ms675935(v=vs.85).aspx
		$GroupTypeDecriptions = @{ 
			0x00000001 = @( 'BUILTIN_LOCAL_GROUP', 'Specifies a group that is created by the system.' );
			0x00000002 = @( 'ACCOUNT_GROUP', 'Specifies a group with global scope.' );
			0x00000004 = @( 'RESOURCE_GROUP', 'Specifies a group with domain local scope.' );
			0x00000008 = @( 'UNIVERSAL_GROUP', 'Specifies a group with universal scope.' );
			0x00000010 = @( 'APP_BASIC_GROUP', 'Specifies an APP_BASIC group for Windows Server Authorization Manager.' );
			0x00000020 = @( 'APP_QUERY_GROUP', 'Specifies an APP_QUERY group for Windows Server Authorization Manager.' );
			0x80000000 = @( 'SECURITY_ENABLED', 'Specifies a security group. If this flag is not set, then the group is a distribution group.' );
		}
		
		# ADS_UF_* User-Account-Control attribute http://msdn.microsoft.com/en-us/library/windows/desktop/ms680832(v=vs.85).aspx
		$UserAccountControlDescriptions = @{ 
			0x00000001 = @( 'SCRIPT', 'The logon script is executed.' );
			0x00000002 = @( 'ACCOUNTDISABLE', 'The user account is disabled.' );
			0x00000008 = @( 'HOMEDIR_REQUIRED', 'The home directory is required.' );
			0x00000010 = @( 'LOCKOUT', 'The account is currently locked out.' );
			0x00000020 = @( 'PASSWD_NOTREQD', 'No password is required.' );
			0x00000040 = @( 'PASSWD_CANT_CHANGE', 'The user cannot change the password.' );
			0x00000080 = @( 'ENCRYPTED_TEXT_PASSWORD_ALLOWED', 'The user can send an encrypted password.' );
			0x00000100 = @( 'TEMP_DUPLICATE_ACCOUNT', 'This is an account for users whose primary account is in another domain. This account provides user access to this domain, but not to any domain that trusts this domain. Also known as a local user account.' );
			0x00000200 = @( 'NORMAL_ACCOUNT', 'This is a default account type that represents a typical user.' );
			0x00000800 = @( 'INTERDOMAIN_TRUST_ACCOUNT', 'This is a permit to trust account for a system domain that trusts other domains.' );
			0x00001000 = @( 'WORKSTATION_TRUST_ACCOUNT', 'This is a computer account for a computer that is a member of this domain.' );
			0x00002000 = @( 'SERVER_TRUST_ACCOUNT', 'This is a computer account for a system backup domain controller that is a member of this domain.' );
			0x00010000 = @( 'DONT_EXPIRE_PASSWD', 'The password for this account will never expire.' );
			0x00020000 = @( 'MNS_LOGON_ACCOUNT', 'This is an MNS logon account.' );
			0x00040000 = @( 'SMARTCARD_REQUIRED', 'The user must log on using a smart card.' );
			0x00080000 = @( 'TRUSTED_FOR_DELEGATION', 'The service account (user or computer account), under which a service runs, is trusted for Kerberos delegation. Any such service can impersonate a client requesting the service.' );
			0x00100000 = @( 'NOT_DELEGATED', 'The security context of the user will not be delegated to a service even if the service account is set as trusted for Kerberos delegation.' );
			0x00200000 = @( 'USE_DES_KEY_ONLY', 'Restrict this principal to use only Data Encryption Standard (DES) encryption types for keys.' );
			0x00400000 = @( 'DONT_REQUIRE_PREAUTH', 'This account does not require Kerberos pre-authentication for logon.' );
			0x00800000 = @( 'PASSWORD_EXPIRED', 'The user password has expired. This flag is created by the system using data from the Pwd-Last-Set attribute and the domain policy.' );
			0x01000000 = @( 'TRUSTED_TO_AUTHENTICATE_FOR_DELEGATION', 'The account is enabled for delegation. This is a security-sensitive setting; accounts with this option enabled should be strictly controlled. This setting enables a service running under the account to assume a client identity and authenticate as that user to other remote servers on the network.' );
			0x04000000 = @( 'PARTIAL_SECRETS_ACCOUNT', 'The account is a read-only domain controller (RODC). This is a security-sensitive setting. Removing this setting from an RODC compromises security on that server.' )
		}
				
		# Bind to current domain's schema.  
		$schema =[DirectoryServices.ActiveDirectory.ActiveDirectorySchema]::GetCurrentSchema()
		
		$resultPropertyExpressions = @()
		
	} 

	#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

#endregion Script Header
	
	Process {
		
		# If no directory searcher provided, bind to LDAP server.   
		If ( -Not $DirectorySearcher ) {
			Write-Debug "Calling:,Connect-DirectorySearcher"
			$directorySearcher = Connect-DirectorySearcher -Server $Server -SearchBase $SearchBase -ViewEntireForest:$ViewEntireForest
			Write-Debug "Returning from:,Connect-DirectorySearcher"
		}
		Write-Debug "`$directorySearcher.SearchRoot.DistinguishedName:,$($directorySearcher.SearchRoot.DistinguishedName)"
		
		# Set search scope.
		$directorySearcher.SearchScope = $SearchScope
		Write-Debug "`$directorySearcher.SearchScope:,$($directorySearcher.SearchScope)"
			
		# Validate and set search filter. 
		If ( $LDAPFilter -And
			-Not $LDAPFilter.StartsWith( '(' ) -And $LDAPFilter.EndsWith( ')' ) ) { 
			$LDAPFilter = "($LDAPFilter)" 
		}
		If ($LDAPFilter) { 
			$directorySearcher.Filter = $LDAPFilter 
		}
		Write-Debug "`$directorySearcher.Filter:,$($directorySearcher.Filter)"
			
		# Add properties to load.  
		$directorySearcher.PropertiesToLoad.Clear()
		$Properties | 
			ForEach-Object {
				[VOID] $directorySearcher.PropertiesToLoad.Add($_) 
			}
		Write-Debug "`$directorySearcher.PropertiesToLoad:,$($directorySearcher.PropertiesToLoad)"
		
		#Write-Progress -Activity "$($MyInvocation.MyCommand.Name)" -Status "Searching for directory services entries..." 
		
		# Place LDAP query.  Wait for results.
		$directorySearcher.FindAll() | 
			ForEach-Object {
			
				# Copy this $searchResults.Properties to $resultProperties.  
				$resultProperties = $_.Properties # $_.PSObject.Properties
				Write-Debug '--'
				Write-Debug "`$resultProperties.PropertyNames:,$($resultProperties.PropertyNames)"
				
				# Only build name/expression hashtable upon first object.  
				#Write-Debug "`$resultPropertyExpressions.Count:,$($resultPropertyExpressions.Count)"
				If ( -Not $resultPropertyExpressions.Count ) {
				
					Write-Debug "`$Properties:,$Properties"
					If ( $Properties ) { # -Or $Properties -Eq '*' ) {
						# If specified PropertiesToLoad then return specified list of properties.
						$propertiesToShow = $Properties 
					} Else {
						# If Null PropertiesToLoad then return dynamic list of properties.
						$propertiesToShow = $resultProperties.PropertiesLoaded 
					}
					Write-Debug "`$propertiesToShow:,$propertiesToShow"
				
					$propertiesToShow  | 
						ForEach-Object {
							$resultPropertyName = $_.ToString()
							Write-Debug "--"
							Write-Debug "`$resultPropertyName:,$resultPropertyName" 
							Write-Debug	"`$resultProperties.Item($resultPropertyName).Count:,$($resultProperties.Item($resultPropertyName).Count)"
							#$resultProperty = $resultProperties.Item($resultPropertyName)
							
							# Get Active Directory schema metadata for this property.  
							$propertyIsSingleValued = $schema.FindProperty($resultPropertyName).IsSingleValued
							Write-Debug "`$propertyIsSingleValued:,$propertyIsSingleValued"
							
							# Try to get property's type, else get Active Directory schema syntax. # 
							#If ( $resultProperties.Item($resultPropertyName)[0] -NE $NULL ) {
							Try {
								$propertyType = $resultProperties.Item($resultPropertyName)[0].GetType()
							#} Else {
							} Catch {
								# At this point AD may have returned a NULL value for this property in the first result object, 
								#  but SHOULD have preserved the data type.
								# If data type is not available then the default expression will be used for the remainder of the result set.  
								$propertyType = $schema.FindProperty($resultPropertyName).Syntax
							}
							Write-Debug "`$propertyType:,$propertyType"
							
							# Handle single-value or multi-value properties.  
							If ( $schema.FindProperty($resultPropertyName).IsSingleValued ) {
								
								Switch ( $propertyType ) {
								
									# Write byte arrays as string of escaped hexadecimal.
									'byte[]' {
										$resultPropertyExpressions += @{Name=$resultPropertyName;Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0]; 
`$string = ''; 
ForEach ( `$char In `$value ) { 
	`$string += `"\`$(`$char.ToString('X2'))`" 
}; 
`$string
"@ )}
										Break
									}
									
									# Handle default single-value properties.
									Default {
										$resultPropertyExpressions += @{Name=$resultPropertyName;Expression=[ScriptBlock]::Create( @"
`$resultProperties.Item('$resultPropertyName')[0]
"@ )}
									}
									
								}
								
#region Augment single-value properties
								If ( $AugmentProperties ) {
								
									Switch ( $resultPropertyName ) {
									
										'accountExpires' {
											$resultPropertyExpressions += @{Name='accountExpiresDate';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
If ( `$value -NE [Long]::MaxValue ) { 
	[DateTime]::FromFileTime( `$value ) 
} Else { 
	$NULL 
}
"@ )}
											$resultPropertyExpressions += @{Name='accountExpiresDaysUntil';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
If ( `$value -NE [Long]::MaxValue ) { 
	-( New-TimeSpan ([DateTime]::FromFileTime( `$value )) (Get-Date) ).Days
} Else {
	$NULL
}
"@ )}
											Break
										} 
										
										'badPasswordTime' {
											$resultPropertyExpressions += @{Name='badPasswordTimeDate';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
If ( `$value -NE [Long]::MaxValue ) { 
	[DateTime]::FromFileTime( `$value ) 
} Else { 
	$NULL 
}
"@ )}
											$resultPropertyExpressions += @{Name='badPasswordTimeDaysSince';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
If ( `$value -NE [Long]::MaxValue ) { 
	( New-TimeSpan ([DateTime]::FromFileTime( `$value )) (Get-Date) ).Days
} Else {
	$NULL
}
"@ )}
											Break
										} 
										
										'distinguishedName' {
											$resultPropertyExpressions += @{Name='OrganizationalUnit';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0].Replace( '\,', '%2E' ) -Match '^(?<RDN>CN|OU=[^,]*),(?<PDN>.*?),(?<DC>DC=.*)$'; 
( (`$Matches.DC).Replace( 'DC=', '' ).Replace( ',', '.' ) + '/' + ( (`$Matches.PDN -Split ',')[((`$Matches.PDN -Split ',').Length-1)..0] -Join '/' ).Replace( 'CN=', '' ).Replace( 'OU=', '' ) ).Replace( '%2E', '\,' )
"@ )}
											Break
										} 
										
										'groupType' {
											$resultPropertyExpressions += @{Name='groupTypeDescription';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
`$description = @(); 
`$GroupTypeDecriptions.Keys | 
	Sort-Object | 
	ForEach-Object { 
		If ( ( `$value -BAnd `$_ ) -Eq `$_ ) { 
			`$description += `$GroupTypeDecriptions[ `$_ ][0] 
		} 
	} 
`$description -Join '$IntraValueDelimiter'
"@ )}
											Break
										} 
										
										'lastLogoff' {
											$resultPropertyExpressions += @{Name='lastLogoffDate';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
If ( `$value -NE [Long]::MaxValue ) { 
	[DateTime]::FromFileTime( `$value ) 
} Else { 
	$NULL 
}
"@ )}
											$resultPropertyExpressions += @{Name='lastLogoffDaysSince';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
If ( `$value -NE [Long]::MaxValue ) { 
	( New-TimeSpan ([DateTime]::FromFileTime( `$value )) (Get-Date) ).Days
} Else {
	$NULL
}
"@ )}
											Break
										} 
										
										'lastLogon' {
											$resultPropertyExpressions += @{Name='lastLogonDate';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
If ( `$value -NE [Long]::MaxValue ) { 
	[DateTime]::FromFileTime( `$value ) 
} Else { 
	$NULL 
}
"@ )}
											$resultPropertyExpressions += @{Name='lastLogonDaysSince';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
If ( `$value -NE [Long]::MaxValue ) { 
	( New-TimeSpan ([DateTime]::FromFileTime( `$value )) (Get-Date) ).Days
} Else {
	$NULL
}
"@ )}
											Break
										} 
										
										'lastLogonTimestamp' {
											$resultPropertyExpressions += @{Name='lastLogonTimestampDate';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
If ( `$value -NE [Long]::MaxValue ) { 
	[DateTime]::FromFileTime( `$value ) 
} Else { 
	$NULL 
}
"@ )}
											$resultPropertyExpressions += @{Name='lastLogonTimestampDaysSince';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
If ( `$value -NE [Long]::MaxValue ) { 
	( New-TimeSpan ([DateTime]::FromFileTime( `$value )) (Get-Date) ).Days
} Else {
	$NULL
}
"@ )}
											Break
										} 
										
										'lockoutTime' {
											$resultPropertyExpressions += @{Name='lockoutTimeDate';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
If ( `$value -NE [Long]::MaxValue ) { 
	[DateTime]::FromFileTime( `$value ) 
} Else { 
	$NULL 
}
"@ )}
											$resultPropertyExpressions += @{Name='lockoutTimeDaysSince';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
If ( `$value -NE [Long]::MaxValue ) { 
	( New-TimeSpan ([DateTime]::FromFileTime( `$value )) (Get-Date) ).Days
} Else {
	$NULL
}
"@ )}
											Break
										} 
										
										'pwdLastSet' {
											$resultPropertyExpressions += @{Name='pwdLastSetDate';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
If ( `$value -NE [Long]::MaxValue ) { 
	[DateTime]::FromFileTime( `$value ) 
} Else { 
	$NULL 
}
"@ )}
											$resultPropertyExpressions += @{Name='pwdLastSetDaysSince';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
If ( `$value -NE [Long]::MaxValue ) { 
	( New-TimeSpan ([DateTime]::FromFileTime( `$value )) (Get-Date) ).Days
} Else {
	$NULL
}
"@ )}
											Break
										} 
										
										'userAccountControl' {
											$resultPropertyExpressions += @{Name='userAccountControlDescription';Expression=[ScriptBlock]::Create( @"
`$value = `$resultProperties.Item('$resultPropertyName')[0];
`$description = @(); 
`$UserAccountControlDescriptions.Keys | 
	Sort-Object | 
	ForEach-Object { 
		If ( ( `$value -BAnd `$_ ) -Eq `$_ ) { 
			`$description += `$userAccountControlDescriptions[ `$_ ][0] 
		} 
	} 
`$description -Join '$IntraValueDelimiter'
"@ )}
											Break
										} 
										
									}
									
								} 
#endregion Augment single-value properties
								
							} Else {
							
								# Handle default multi-value properties.  
								$resultPropertyExpressions += @{Name=$resultPropertyName;Expression=[ScriptBlock]::Create( @"
`$resultProperties.Item('$resultPropertyName')
"@ )} 
#`$resultProperties.Item('$resultPropertyName') -Join '$IntraValueDelimiter'

#region Augment multi-value properties
								If ( $AugmentProperties ) {
								
									Switch ( $resultPropertyName ) {
										
										'member' {
											$resultPropertyExpressions += @{Name='memberCount';Expression=[ScriptBlock]::Create( @"
(`$resultProperties.Item('$resultPropertyName')).Count
"@ )}
											Break
										} 
										
										'memberOf' {
											$resultPropertyExpressions += @{Name='memberOfCount';Expression=[ScriptBlock]::Create( @"
(`$resultProperties.Item('$resultPropertyName')).Count
"@ )}
											Break
										} 
										
									}
								
								} 
#endregion Augment multi-value properties
																
							}
						} 
						
				} 
				#Write-Debug "`$resultPropertyExpressions.Count:,$($resultPropertyExpressions.Count)"
				#If ( $Debug ) {
				#	ForEach ( $propertyExpression In $resultPropertyExpressions ) {
				#		ForEach ( $key In $propertyExpression.Keys ) {
				#			Write-Debug "`$propertyExpression[$key];,$($propertyExpression[$key])"
				#		}
				#	}
				#}
				
				Select-Object -InputObject $resultProperties -Property $resultPropertyExpressions # Write-Output
				
			} 
	
	}		
}

#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
#---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

<#
	.Synopsis
	Convert a LDAP filter value into an escaped LDAP filter value.
	  
	.Description
	When building an LDAP filter certain characters are special characters: Null '\00', left-parenthesis '(', right-parenthesis ')', asterisk '*' and reverse-solidus '\'.  When an LDAP filter value comes from an unknown source it is best to escape these special characters to a hexidecimal format '\FF'.  Optionally allow the wildcard special character asterisk '*' to not be escaped.

	.Parameter AttributeValue
	The LDAP filter string value.

	.Parameter AllowWildcard 
	Optional switch to suppress escaping the wildcard special character asterisk '*' and allow it to pass through as is.  

	.Outputs
	Returns an LDAP filter string value.

	.Notes
	Author: Terry E Dow
	Last Modified: 2014-02-22

	.Link
#>
Filter ConvertTo-EscapedLDAPSearchFilterAttributeValue {
	[CmdletBinding()]
	Param( 
		[Parameter( HelpMessage='The LDAP filter string value.', 
			ValueFromPipeline=$TRUE, 
			ValueFromPipelineByPropertyName=$TRUE ) ]
		[Alias('Identity')]
			[String] $AttributeValue = '',
		
		[ Parameter( HelpMessage='Optional switch to suppress escaping the wildcard special character asterisk "*" and allow it to pass through as is.' ) ] 
			[Switch] $AllowWildcard = $FALSE
	)

#region Script Header
	
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

	# Define escape character substitution variables.  
	$invalidLDAPCharactersWithWildcard = @{ 
		[String][Char]0 = '\00'; 
		'(' = '\28'; 
		')' = '\29'; 
		'*' = '\2A'; 
		#'\' = '\5C' 
	}		
	
	$invalidLDAPCharactersWithoutWildcard = @{ 
		[String][Char]0 = '\00'; 
		'(' = '\28'; 
		')' = '\29'; 
		#'\' = '\5C' 
	}		

#endregion Script Header

	# Escape RFC2254 LDAP search filter attribute value special characters. http://www.ietf.org/rfc/rfc2254.txt
	$escapedAttributeValue = $attributeValue.Replace( '\', '\5C' )
	If ( $AllowWildcard ) {
		$invalidLDAPCharactersWithoutWildcard.Keys |
			ForEach-Object {
				$escapedAttributeValue = $escapedAttributeValue.Replace( $_, $invalidLDAPCharactersWithoutWildcard[$_] )
			}
	} Else {
		$invalidLDAPCharactersWithWildcard.Keys |
			ForEach-Object {
				$escapedAttributeValue = $escapedAttributeValue.Replace( $_, $invalidLDAPCharactersWithWildcard[$_] )
			}
	}
	# <<<< TODO: Expand to replace high ASCII (128-255) characters to their escaped values.  
	Write-Debug "`$escapedAttributeValue:,$escapedAttributeValue"
	
	Write-Output $escapedAttributeValue
}
