#Requires -Version 5.0
<# 
.SYNOPSIS
  Pull data from Employee Database (Escape Online) and 
  update Active Direcrtory user objects using employeeID as the foreign key. 
.DESCRIPTION
.EXAMPLE
  .\Update-ADFromEscape.ps1 -DC $dc -ADCredential $adCred -DBHash $dbHash
.EXAMPLE
  .\Update-ADFromEscape.ps1 -DC $dc -ADCredential $adCred -DBHash $dbHash -WhatIf -Verbose -Debug
.INPUTS
  Common parameters are used as inputs.
.OUTPUTS
  A log file is generated when using -Log [switch]
.NOTES
#>

[cmdletbinding()]
param ( 
 [Parameter(Mandatory = $True)]
 [Alias('DC')]
 [ValidateScript( { Test-Connection -ComputerName $_ -Quiet -Count 1 })]
 [string]$DomainController,
 # PSSession to Domain Controller and Use Active Directory CMDLETS  
 [Parameter(Mandatory = $True)]
 [Alias('ADCred')]
 [System.Management.Automation.PSCredential]$ADCredential,
 # String formatted as 'server\database'
 [Parameter(Mandatory = $True)]
 [Alias('SearchBase')]
 [string]$ActiveDirectorySearchBase,
 [Parameter(Mandatory = $True)]
 [string]$SQLServer,
 [Parameter(Mandatory = $True)]
 [string]$SQLDatabase,
 # Credential object with database select permission.
 [Parameter(Mandatory = $True)]
 [Alias('SQLCred')]
 [System.Management.Automation.PSCredential]$SQLCredential,
 [SWITCH]$WhatIf
)

Get-PSSession | Remove-PSSession -WhatIf:$false

# AD Domain Controller Session
$adCmdLets = 'Get-ADUser', 'Set-ADUser'
$adSession = New-PSSession -ComputerName $DomainController -Credential $ADCredential
Import-PSSession -Session $adSession -Module ActiveDirectory -CommandName $adCmdLets -AllowClobber | Out-Null

# Imported Functions
. '.\lib\Add-Log.ps1'
. '.\lib\Invoke-SqlCommand.PS1'

# Processing

# Database Connection
$query = Get-Content -Path '.\sql\active-employees.sql' -Raw
$dbResults = Invoke-SqlCommand $SQLServer $SQLDatabase $SQLCredential $query

$aDParams = @{
 Filter     = { (mail -like "*@*") -and (employeeID -like "*") }
 Searchbase = $ActiveDirectorySearchBase
 <# Limiting the properties reduces script run-time
  The extra AD properties are pulled from the SQL result column Aliases. #>
 Properties = ($dbResults | Get-Member -MemberType Properties).name
}
$aDStaff = Get-Aduser @aDParams

# Update AD Attributes
foreach ( $row in $dbResults ) {
 # Process Rows
 $user = $aDStaff.where( { $_.employeeID -eq $row.employeeID })
 Write-Debug  "Process? $($row.employeeid) | $($user.SamAccountName) | $($row.givenname) $($row.sn)"
 if ( $user ) {
  # Begin Check if $user exists
  foreach ( $prop in  (($row | Get-Member -MemberType Properties).name) ) {
   # Begin parse the db column names
   if ( ($row."$prop") -and ($row."$prop" -notmatch "^\s{1,}") ) {
    # Begin Check if value is present
    $value = $row."$prop"
    # Begin case-sensitive compare data between AD and DB
    if ($user."$prop" -cnotmatch $value) {
     Add-Log update ("{0},{1},{2} => {3}" -f $user.SamAccountName, $prop, $($user."$prop"), $value)
     # Set-ADUSer -Replace works for updating most common attributes.
     Set-ADUser -Identity $user.ObjectGUID -Replace @{$prop = $value } -WhatIf:$WhatIf
    } # End  compare data between AD and DB
   } # End Check if value is present
  } # End parse the db column names
  # Fix O365 Global Address Book enrty
  if ( (Get-ADuser -Identity $user.ObjectGUID).msExchHideFromAddressLists -eq $true ) {
   Add-Log update ("{0},msExchHideFromAddressLists = FALSE" -f $user.SamAccountName)
   Set-ADUser -Identity $user.ObjectGUID -Replace @{msExchHideFromAddressLists=$false} -Whatif:$WhatIf
  }
  # Renames the user object if name change detected
  $displayName = $user.GivenName+' '+$user.Surname
  $refreshedUserData = Get-ADUser $user.ObjectGUID -Properties *
  if ( ($refreshedUserData.displayname -notmatch $user.GivenName) -or ($refreshedUserData.displayname -notmatch $user.SurName) ){
   $newDisplayName = $refreshedUserData.GivenName+' '+$refreshedUserData.Surname
   Set-ADuser -Identity $user.ObjectGUID -DisplayName $newDisplayName -Confirm:$false -WhatIf:$WhatIf
   Rename-ADObject -Identity $user.ObjectGUID -NewName $newDisplayName -Confirm:$false -WhatIf:$WhatIf
   Add-Log rename ('{0} renamed to {1}' -f $displayName, $newDisplayName) -Whatif:$WhatIf
  }
 } # End Check if $user exists
} # End Process Rows

Write-Verbose "Tearing down sessions"
Get-PSSession | Remove-PSSession -WhatIf:$false