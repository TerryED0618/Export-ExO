Se Export-ExO.pptx for description and usage.

Remove ".TXT" file name extension from files with a nested extensions.  For example rename BATCH.CMD.TXT to BATCH.CMD or SCRIPT.PS1.TXT to SCRIPT.PS1.

# Remove *.TXT extension for any file name with a nested extension ending with '.TXT'
Get-ChildItem -Path *.*.TXT -File -Recurse | ForEach-Object { Rename-Item -Path $PSItem.VersionInfo.FileName -NewName ( $PSItem.VersionInfo.FileName -Replace '\.TXT$', '' ) -PassThru }

# Add *.TXT extension for any file name ending with .CMD or .PS1 extensions.
Get-ChildItem -Path *.CMD,*.PS1 -File -Recurse | ForEach-Object { Rename-Item -Path $PSItem.VersionInfo.FileName -NewName "$($PSItem.VersionInfo.FileName).TXT" -PassThru }
