<#
	.SYNOPSIS
	Expand all property values by joining arrays into singleton strings.  

	.DESCRIPTION
	Export-CSV does not expanding properties that are arrays or collections but instead renders that property's data type.  
	This function expands collection properties that are arrays to a singleton string, suitible for Export-CSV.  It uses an intra-value separator, by default a semicolon (;), when joining the array/collection elements together.  

	.INPUTS
	InputObject
	An object to be formatted.  
	
	.PARAMETER AppendTo PSObject
	An object that the InputObject's properties will be appended to.  If no object is provided a new one is created.  Note: subsequent redundent property names will be skipped.  
		
	.PARAMETER IntraValueDelimiter String
	The string used to joing array properties values together.  Default is a semicolon (;).

	.PARAMETER RetainNewline Switch
	By default carraige return and linefeed characters are replaced with "\r" and "\n".  Include this switch to not replace them.
	
	.PARAMETER ByteArrayFull Switch
	By default to export byte arrays digest - first and last 8 bytes only.  Include this switch to export all bytes.
	
	.PARAMETER ConvertByteQuantifiedSizeToBytes Switch
	Convert Exchange ByteQuantifiedSize type to bytes.  Include this switch to not replace them.
		
	.OUTPUTS
	PSObject of properties, arrays contents converted to strings that is suitable for Export-CSV.  

	.EXAMPLE
	Create a sorted CSV file export of the Get-Process command named 'processes.csv'.  The Get-Process column 'Modules' would normally show the data type (System.Diagnostics.ProcessModuleCollection) not the array's values.  
	
	Get-Process | 
		Format-ExpandAllProperties | 
		Export-CSV processes.csv
	
	.EXAMPLE
	Create a sorted CSV file export of the Get-Process command named 'processes.csv', using the forward-slash "/" to seperate array items.
	
	Get-Process | 
		Sort-Object -Property CPU -Descending |
		Format-ExpandAllProperties -IntraValueDelimiter '/' | 
		Export-CSV processes.csv
		
	.EXAMPLE
	Combine more than one datasource of related objects into one exportable collection.  
	Export the current folder properties:
	
	$folder = Get-Item .\ | 
		Format-ExpandAllProperties
		
	Then export each file property including the parent folder properties:  
	
	Get-ChildItem *.* | 
		Format-ExpandAllProperties -AppendTo $folder |
		Export-CSV FolderFiles.csv

	.NOTES
	Author: Terry E Dow
	Last Modified: 2016-01-18

	.LINK
	
#>
Function Format-ExpandAllProperties {
    [ CmdletBinding( 
		SupportsShouldProcess = $TRUE # Enable support for -WhatIf. 
	) ] 
    Param(
		[ Parameter( ValueFromPipeline=$TRUE, 
			#ValueFromPipelineByPropertyName=$TRUE,
			HelpMessage='Specifies objects to send to the cmdlet through the pipeline. This parameter enables you to pipe objects to Select-Object.' ) ] 
			[PSObject[]] $InputObject,
			
		[ Parameter( HelpMessage='Optionally specify existing object data that the new InputObject propteries are being appended to.' ) ] 
			[PSObject] $AppendTo = $NULL,
		
		[ Parameter( HelpMessage='The delimiter string used to joining array properties values together.  The default is a semicolon (;).' ) ] 
			[String] $IntraValueDelimiter = ';',
			
		[ Parameter( HelpMessage='By default carraige return and linefeed characters are replaced with "\r" and "\n".  Include this switch to not replace them.' ) ] 
			[Switch] $RetainNewline,
	
		[ Parameter( HelpMessage='By default to export byte arrays digest - first and last 8 bytes only.  Include this switch to export all bytes.' ) ] 
			[Switch] $ByteArrayFull,
	
		[ Parameter( HelpMessage='Convert Exchange ByteQuantifiedSize type to bytes.  Include this switch to not replace them.' ) ] 
		[Switch] $ConvertByteQuantifiedSizeToBytes
	)

#region Script Header

#Requires -version 3

	Begin { 

		Set-StrictMode -Version Latest

		# Detect cmdlet common parameters.  
		$cmdletBoundParameters = $PSCmdlet.MyInvocation.BoundParameters
		$Debug = If ( $cmdletBoundParameters.ContainsKey('Debug') ) { $cmdletBoundParameters['Debug'] } Else { $FALSE }
		# Replace default -Debug preference from 'Inquire' to 'Continue'.  
		If ( $DebugPreference -Eq 'Inquire' ) { $DebugPreference = 'Continue' }
		$Verbose = If ( $cmdletBoundParameters.ContainsKey('Verbose') ) { $cmdletBoundParameters['Verbose'] } Else { $FALSE }
		$WhatIf = If ( $cmdletBoundParameters.ContainsKey('WhatIf') ) { $cmdletBoundParameters['WhatIf'] } Else { $FALSE }
		Remove-Variable -Name cmdletBoundParameters -WhatIf:$FALSE

		$matchByteQuantifiedSize = New-Object RegEx '[\d,.]*\s(?:K|M|G|T)?B\s\((?<ToBytes>[\d,.]*)\sbytes\)', @( 'Compiled', 'IgnoreCase' )
		
		$propertyExpressions = @()
	}

#endregion Script Header

	Process { 
	
		ForEach ( $object In ($InputObject) ) {
			# If not already done, dynamically build a collection of name/expression hash table pairs.
			# The expression script block will attempt to join (expand) all array values with an intra-value delimiter.  
			If ( -Not $propertyExpressions.Count ) {
			
				$propertyNames = @{}
			
				# If specified, first build a collection of name/expression hash table pairs for each of the $AppendTo's properties.
				If ( $AppendTo ) {
					#Write-Debug "`$AppendTo:$AppendTo"
					ForEach ( $property In $AppendTo.PSObject.Properties ) {
							
						$propertyName = $property.Name
						#Write-Debug ''
						#Write-Debug "`$propertyName:$propertyName"
						#Write-Debug "`$property.TypeNameOfValue:$($property.TypeNameOfValue)"
						#Write-Debug "`$property.Value:$($property.Value)"
						
						If ( -Not $propertyNames.ContainsKey( $propertyName ) ) {
						
							$propertyExpressions += @{Name=$propertyName;Expression=[ScriptBlock]::Create("`$AppendTo.'$propertyName' -Join '$IntraValueDelimiter'")}
							
							$propertyNames.Add( $propertyName, $TRUE )
						}
							
					}
				}
				
				# Append the name/expression hash table pairs for each of the $InputObject's properties.
				ForEach ( $property In $object.PSObject.Properties ) {

					$propertyName = $property.Name
					#Write-Debug ''
					#Write-Debug "`$propertyName:$propertyName"
					#Write-Debug "`$property.TypeNameOfValue:$($property.TypeNameOfValue)"
					#Write-Debug "`$property.Value:$($property.Value)"
					
					If ( -Not $propertyNames.ContainsKey( $propertyName ) ) {
					
						Switch ( $property.TypeNameOfValue ) {
						
							'System.Byte[]' { 
								If ( $ByteArrayFull ) {
									#@{ Name='$propertyName'; Expression={ $string = ''; ForEach ( $byte In $PSItem.'$propertyName' ) { $string += "\$($byte.ToString('X2'))" }; $string } }
									$propertyExpression = "`$string = ''; ForEach ( `$byte In `$PSItem.'$propertyName' ) { `$string += ""\`$(`$byte.ToString('X2'))"" }; `$string"
								} Else {
									#@{ Name='$propertyName'; Expression={ ( $PSItem.'$propertyName'[0..3] -Join ';' ) + '...' + ( $PSItem.'$propertyName'[-4..-1] -Join ';' ) } }
									$propertyExpression = "If ( `$PSItem.'$propertyName'.Length -LE 16 ) { `$string = ''; ForEach ( `$byte In `$PSItem.'$propertyName' ) { `$string += ""\`$(`$byte.ToString('X2'))"" }; `$string } Else { `$string = ''; ForEach ( `$byte In `$PSItem.'$propertyName'[0..7] ) { `$string += ""\`$(`$byte.ToString('X2'))"" }; `$string += '...'; ForEach ( `$byte In `$PSItem.'$propertyName'[-8..-1] ) { `$string += ""\`$(`$byte.ToString('X2'))"" }; `$string }"
								}
								Break
							}
							
							Default {
								#@{ Name='$propertyName'; Expression={ $PSItem.'$propertyName' -Join ';' } }
								$propertyExpression = "`$PSItem.'$propertyName' -Join '$IntraValueDelimiter'"
								#Write-Debug "`$propertyExpression:,$propertyExpression"
							
								If ( -Not $RetainNewline ) {
									#@{ Name='$propertyName'; Expression={ ($PSItem.'$propertyName' -Join ';').Replace('`r','\r').Replace('`n','\n') } }
									$propertyExpression = "($propertyExpression).Replace('`r','\r').Replace('`n','\n')"
									#Write-Debug "`$propertyExpression:,$propertyExpression"
								}
								
								If ( $ConvertByteQuantifiedSizeToBytes ) {
									# Undo Exchange's default pretty printing to something more useful.  Must you -Replace operator (instead of System.String.Replace method) inorder to get RegEx support.  
									#@{ Name='$propertyName'; Expression={ (($PSItem.'$propertyName' -Join ';').Replace('`r','\r').Replace('`n','\n')) -Replace $matchByteQuantifiedSize, '${ToBytes}' } }
									$propertyExpression = "($propertyExpression) -Replace `$matchByteQuantifiedSize, '`${ToBytes}'"
									#Write-Debug "`$propertyExpression:,$propertyExpression"
								}
							}
						}
						
						$propertyExpressions += @{ Name=$propertyName; Expression=[ScriptBlock]::Create($propertyExpression) }
						
						$propertyNames.Add( $propertyName, $TRUE )
					}
				}
				#Write-Debug "`$propertyExpressions.Count:,$($propertyExpressions.Count)"
				#If ( $Debug ) {
				#	ForEach ( $property In $propertyExpressions ) {
				#		ForEach ( $key In $property.Keys ) {
				#			Write-Debug "`$property[$key];,$($property[$key])"
				#		}
				#	}
				#}
			#} Else {
			#	Write-Debug "`$properties already exists."
			}
			
			Select-Object -InputObject $object -Property $propertyExpressions # Write-Output
		}
	}	
}