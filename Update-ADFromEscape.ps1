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
 [Alias('wi')]
 [SWITCH]$WhatIf
)

Get-PSSession | Remove-PSSession -WhatIf:$false

# AD Domain Controller Session
$adCmdLets = 'Get-ADUser', 'Set-ADUser', 'Rename-ADObject'
$adSession = New-PSSession -ComputerName $DomainController -Credential $ADCredential
Import-PSSession -Session $adSession -Module ActiveDirectory -CommandName $adCmdLets -AllowClobber | Out-Null

# Imported Functions
. '.\lib\Add-Log.ps1'
. '.\lib\Invoke-SqlCommand.PS1'

# Processing
for ($i = 1; $i -lt 5; $i++) {
 Add-Log loop $i
 # Database Connection
 $query = Get-Content -Path '.\sql\active-employees.sql' -Raw
 $dbResults = Invoke-SqlCommand $SQLServer $SQLDatabase $SQLCredential $query
 Add-Log dbresults $dbResults.count

 $aDParams = @{
  Filter     = { (mail -like "*@*") -and (employeeID -like "*") }
  Searchbase = $ActiveDirectorySearchBase
  <# Limiting the properties reduces script run-time
   The extra AD properties are pulled from the SQL result column Aliases. #>
  Properties = ($dbResults | Get-Member -MemberType Properties).name
 }
 $aDStaff = Get-Aduser @aDParams
 Add-Log adStaff $aDStaff.count

 # Update AD Attributes
 foreach ( $row in $dbResults ) {
  # Process Rows
  $user = $aDStaff.where( { $_.employeeID -eq $row.employeeID })
  Write-Debug  "Process? $($row.employeeid) | $($user.SamAccountName) | $($row.givenname) $($row.sn)"
  if ( $user ) {
   # Begin Check if $user exists
   foreach ( $prop in  (($row | Get-Member -MemberType Properties).name) ) {
    # Begin parse the db column names
    # if ( ($row."$prop") -and ($row."$prop" -notmatch "^\s{1,}") ) {
    if ( ($row."$prop") -and ($row."$prop" -ne '') ) {
     # Begin Check if rowProp is present in AD Object
     $rowProp = $row."$prop"
     # Begin case-sensitive compare data between AD and DB
     if ( $user."$prop" -cnotcontains $rowProp ) {
      Add-Log update ("{0},{1},{2} => {3}" -f $user.SamAccountName, $prop, $($user."$prop"), $rowProp)
      # Set-ADUSer -Replace works for updating most common attributes.
      Set-ADUser -Identity $user.ObjectGUID -Replace @{$prop = $rowProp } -WhatIf:$WhatIf
     } # End  compare data between AD and DB
    } # End Check if rowProp is present
   } # End parse the db column names
   # Fix O365 Global Address Book enrty
   # msExchHideFromAddressLists is not a header in the list of rows from the DB query
   if ( (Get-ADuser -Identity $user.ObjectGUID).msExchHideFromAddressLists -eq $true ) {
    Add-Log addressbook ("{0},msExchHideFromAddressLists = FALSE" -f $user.SamAccountName)
    Set-ADUser -Identity $user.ObjectGUID -Replace @{msExchHideFromAddressLists=$false} -Whatif:$WhatIf
   }
   # Renames the user object if name change detected
   # In the event of a name change this will overwrite custom Display Name data in AD
   $displayName = $user.GivenName+' '+$user.Surname
   if ( ($row.GivenName -cnotcontains $user.GivenName) -or ($row.sn -cnotcontains $user.SurName) ){
    $newDisplayName = $row.GivenName+' '+$row.sn
    Set-ADuser -Identity $user.ObjectGUID -DisplayName $newDisplayName -Confirm:$false -WhatIf:$WhatIf
    Rename-ADObject -Identity $user.ObjectGUID -NewName $newDisplayName -Confirm:$false -WhatIf:$WhatIf
    Add-Log rename ('{0},{1} renamed to {2}' -f $row.employeeID, $displayName, $newDisplayName) -Whatif:$WhatIf
    # read-host
   }
  } # End Check if $user exists
 } # End Process Rows
 $waitSeconds = if ($WhatIf) { 1 } else { (60*60*2) }
 Add-Log wait ( 'Running again at {0}' -f $((Get-Date).AddSeconds($waitSeconds)) )
 Start-Sleep $waitSeconds
}

Write-Verbose "Tearing down sessions"
Get-PSSession | Remove-PSSession -WhatIf:$false