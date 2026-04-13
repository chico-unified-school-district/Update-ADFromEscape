<#
.SYNOPSIS
  Pull data from Employee Database (Escape Online) and
  update Active Direcrtory user object attrbutes using employeeId as the foreign key.
.DESCRIPTION
.EXAMPLE
.\Update-ADFromEscape.ps1 -DomainController DC1.our.org -ADCredential $adCredObj -SearchBase 'OU=Employees,DC=our,DC=org' -SQLServer EscapeDBServer.our.org -SQLDatabase EscapeOnline -SQLCred $sqlCredObj
.EXAMPLE
  .\Update-ADFromEscape.ps1 -DomainController DC1.our.org -ADCredential $adCredObj -SearchBase 'OU=Employees,DC=our,DC=org' -SQLServer EscapeDBServer.our.org -SQLDatabase EscapeOnline -SQLCred $sqlCredObj -WhatIf -Verbose -Debug
.INPUTS
  Common parameters are used as inputs.
.OUTPUTS
.NOTES
#>

[cmdletbinding()]
param (
 [Parameter(Mandatory = $True)][Alias('DCs')][string[]]$DomainControllers,
 [Parameter(Mandatory = $True)][System.Management.Automation.PSCredential]$ADCredential,
 [Parameter(Mandatory = $True)][Alias('SearchBase')][string]$ActiveDirectorySearchBase,
 [Parameter(Mandatory = $True)][string]$EmployeeServer,
 [Parameter(Mandatory = $True)][string]$EmployeeDatabase,
 [Parameter(Mandatory = $True)][System.Management.Automation.PSCredential]$EmployeeCredential,
 [Parameter(Mandatory = $True)][int]$GracePeriodMonths,
 [int[]]$SkipPersonIds,
 [Alias('wi')][SWITCH]$WhatIf
)

function Clear-ADExpireDate {
 process {
  if (!$_.clearExpiration) { return $_ }
  $msg = $MyInvocation.MyCommand.Name, $_.userInfo
  Write-Host ('{0},{1},Clearing expire date' -f $msg) -f DarkCyan
  Set-ADUser -Identity $_.ad.ObjectGUID -AccountExpirationDate $null -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Enable-ADAccount {
 process {
  # Enabling accounts is a sensitive action.
  if ($_.ad.Enabled) { return $_ }
  if ($_.staleSub) {
   Write-Host ('{0},{1},Stale Sub account detected. Skipping enable.' -f $MyInvocation.MyCommand.Name, $_.userInfo) -f Yellow
   return $_
  }
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.userInfo) -f Green
  Set-ADUser -Identity $_.ad.ObjectGUID -Enabled:$true -WhatIf:$WhatIf
 }
}

function Get-ADStaffData ($ou, $properties) {
 Write-Verbose ('{0}' -f $MyInvocation.MyCommand.Name)
 $adParams = @{
  Filter     = 'mail -like "*@*" -and
    EmployeeID -like "*"
    '
  SearchBase = $ou
  Properties = $properties
 }
 Get-ADUser @adParams | Where-Object { $_.EmployeeID -notmatch '[A-Za-z]' }
}

function Get-EmployeeData ($instance) {
 $sql = Get-Content .\sql\active-employees.sql -Raw
 $data = New-SqlOperation -Server $instance -Query $sql | ConvertTo-Csv | ConvertFrom-Csv
 Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, @($data).count) -f Green
 if ($WhatIf) { Start-Sleep -Seconds 5 } # A slight pause on test runs to allow time to read the count before processing begins.
 $data
}

function New-Obj {
 process {
  $usrInf = $_.EmpId + ',' + $_.NameLast + ',' + $_.NameFirst + ',' +
  $_.EmailWork + ',' + $_.EmploymentStatusDescr + ',' + $_.EmploymentStatusCode
  $obj = [PSCustomObject]@{
   ad              = $null
   clearExpiration = $null
   desc            = $null
   emp             = $_
   propertyList    = $null
   site            = $null
   staleSub        = $null
   userInfo        = '[' + $usrInf.trim() + ']'
  }
  Write-Verbose ($MyInvocation.MyCommand.name, $obj | Out-String)
  $obj
 }
}

function Set-ADData ($data) {
 process {
  $id = $_.emp.EmpID
  $_.ad = $data.Where({ $_.EmployeeID -eq $id })
  if (!$_.ad) { return }
  $_
 }
}

function Set-ClearExpiration {
 process {
  $_.clearExpiration = if
  # Expire date occurs today or in the future
  ($_.ad.AccountExpirationDate -ge [DateTime]::Today ) { $false }
  # Student teacher accounts with expiration
  # elseif (($_.ad.AccountExpirationDate -is [datetime]) -and ($_.ad.Description -like '*student*teacher*')) { $false }
  elseif (($_.ad.AccountExpirationDate -is [datetime]) -and ($_.emp.PersonTypeId -eq '6')) { $false }
  # Stale sub with expiration date already set
  elseif ($_.staleSub -and ($_.ad.AccountExpirationDate -is [datetime])) { $false }
  # All others
  elseif ($_.ad.AccountExpirationDate -isnot [datetime]) { $false }
  # What about already expired stale subs?
  else { $true } # clear expire date
  $_
 }
}

function Set-Description {
 process {
  $_.desc = switch ($_) {
   # No Change for expiring accounts
   { $_.ad.AccountExpirationDate -is [datetime] -and ($_.clearExpiration -eq $false) } { $_.ad.Description; break }
   # Remove Expiration date info from description for accounts that will have expire date cleared.
   # This is to prevent confusion and preserve relevant description info.
   { $_.clearExpiration -eq $false -and ($_.ad.Description -match 'Expiration Date') } { ($_.ad.Description.Split('<')[0]) -replace '\s+', ' '; break }
   # Set Description with site and job class info when available.
   { ($_.emp.JobClassDescr -match '[A-Za-z]') -or ($_.site) } { ($_.site.SiteDescrShort + ' ' + $_.emp.JobClassDescr) -replace '\s+', ' '; break }
   default { $_.ad.Description }
  }
  $_
 }
}

function Set-PropertyListData {
 begin {
  function Remove-ExtraSpaces ($string) { $string -replace '\s+', ' ' }
  function Test-Null ($obj) { if ($obj -match '[A-Za-z0-9]') { $obj.Trim() } else { $null } }
 }
 process {
  $initials = if ($_.emp.NameMiddle -match '\w') { $_.emp.NameMiddle.SubString(0, 1) }
  $_.propertyList = [PSCustomObject]@{
   Company                    = 'Chico Unified School District'
   Department                 = Test-Null $_.emp.JobCategoryDescr
   departmentNumber           = Test-Null $_.emp.siteId
   Description                = Test-Null $_.desc
   extensionAttribute1        = Test-Null $_.emp.BargUnitID
   employeeType               = $_.emp.EmploymentStatusCode.Trim()
   GivenName                  = Remove-ExtraSpaces $_.emp.NameFirst
   initials                   = $initials
   middleName                 = Test-Null (Remove-ExtraSpaces $_.emp.NameMiddle)
   physicalDeliveryOfficeName = Test-Null $_.site.SiteDesc
   sn                         = Remove-ExtraSpaces $_.emp.NameLast
   Title                      = Test-Null $_.emp.JobClassDescr
  }
  Write-Verbose ($MyInvocation.MyCommand.name, $_ | Out-String)
  $_
 }
}

function Set-SiteData ($instance) {
 begin {
  $sql = 'SELECT DISTINCT SiteId,SiteDescr,SiteDescrShort FROM OrgSite'
  $siteData = New-SqlOperation -Server $instance -Query $sql | ConvertTo-Csv | ConvertFrom-Csv
  # Write-Verbose ($MyInvocation.MyCommand.Name, $siteData | Out-String)
 }
 process {
  if ($_.emp.SiteId -notmatch '\d') { return $_ }
  $siteId = $_.emp.SiteId
  $_.site = $siteData.Where({ $_.SiteId -eq $siteId })
  if (!$_.site) { Write-Host ('{0},{1},Site not found for SiteId {2}' -f $MyInvocation.MyCommand.Name, $_.userInfo, $siteId) -f Red }
  $_
 }
}

function Set-StaleSubStatus ([int]$months) {
 begin { $pastDate = (Get-Date).AddMonths(-$months) }
 process {
  $_.staleSub = if (
   ($_.emp.EmploymentStatusCode -match 'S') -and
   (($_.ad.LastLogonDate -is [datetime]) -and ($_.ad.LastLogonDate -lt $pastDate) -or
   (($_.ad.LastLogonDate -isnot [datetime]) -and ($_.WhenCreated -lt $pastDate))
   )
  ) {
   Write-Host ('{0},{1},Stale Sub Detected' -f $MyInvocation.MyCommand.Name, $_.userInfo) -f Yellow
   $true
  }
  else { $false }
  $_
 }
}

function Show-Object {
 begin {
  $i = 0
 }
 process {
  $i++
  Write-Verbose ($MyInvocation.MyCommand.name, $_ | Out-String)
  # Read-Host ('{0}' -f ('x' * 50))
 }
 end {
  Write-Host ('{0},Total Processed: {1}' -f $MyInvocation.MyCommand.Name, $i) -f Green
 }
}

function Skip-Ids ([int[]]$ids) {
 process {
  # Skip specific PersonTypeIds. This was done to preserve ad info for student workers
  if ($ids -contains [int]$_.emp.PersonTypeId) {
   Write-Host ('{0},{1},Skipping PersonTypeId {2}' -f $MyInvocation.MyCommand.Name, $_.userInfo, $_.emp.PersonTypeId) -f Yellow
   return
  }
  $_
 }
}

function Update-ADAttributes {
 begin {
  $clearProps = 'extensionAttribute1' # Only clear these attributes when cleared in Escape.
  function skipAttribute ($adObj, $attribName) {
   $rules = Get-Content .\json\customRules.json -Raw | ConvertFrom-Json | Where-Object { $_.customRules.Enabled }
   foreach ($rule in $rules.customRules) {
    if (!$rule.enabled -or $rule.type -ne 'AD Attribute') { continue } # Skip disabled rules
    if ($adObj.info -notlike $rule.keyPhrase) { continue } # Skip if condition not met.
    # if (!($adObj.info -like '*custom*Name*')) { continue } # Skip if condition not met.
    if ($rule.attributes -contains $attribName) {
     Write-Host ('{0},{1},[{2}] Skipping attribute update based on custom rule' -f $MyInvocation.MyCommand.Name, $adObj.SamAccountName, $attribName) -F Magenta
     return $true
    }
   }
  }
 }
 process {
  # Write-Verbose ( $MyInvocation.MyCommand.Name, $_.ad | Format-List | Out-String )
  foreach ($propName in $_.propertyList.PSObject.Properties.Name) {
   if (skipAttribute -adObj $_.ad -attribName $propName) { continue }
   $propValue = $_.propertyList.$propName
   if ( $_.ad.$propName -eq $propValue) { continue } # Skip if value is the same. This prevents unnecessary AD updates and preserves AD data integrity.
   $msgVars = $MyInvocation.MyCommand.Name, $_.userInfo, $propName, $_.ad.$propName, $propValue
   if ($propValue) {
    Write-Host ('{0},{1},{2},[{3}] => [{4}]' -f $msgVars) -Fore Blue
    Set-ADUser -Identity $_.ad.ObjectGUID -Replace @{$propName = $propValue } -WhatIf:$WhatIf
   }
   else {
    if ($clearProps -notcontains $propName) { continue } # Only clear properties that are in the $clearProps list. This prevents accidental clearing of AD attributes.
    Write-Host ('{0},{1},{2},[{3}] => [{4}]' -f $msgVars) -Fore Cyan
    Set-ADUser -Identity $_.ad.ObjectGUID -Clear $propName -WhatIf:$WhatIf
   }
  }
  $_
 }
}

# =======================================================================================
Import-Module -Name CommonScriptFunctions
Import-Module -Name dbatools -Cmdlet 'Invoke-DbaQuery', 'Set-DbatoolsConfig', 'Connect-DbaInstance', 'Disconnect-DbaInstance'

Show-BlockInfo main
if ($WhatIf) { Show-TestRun }
Clear-SessionData

$gam = 'C:\GAM7\gam.exe'

$cmdlets = 'Get-ADuser', 'Set-ADuser', 'Rename-ADObject', 'Clear-ADAccountExpiration'
Connect-ADSession -DomainControllers $DomainControllers -Cmdlets $cmdLets -Credential $ADCredential

$aDProperties = @(
 'AccountExpirationDate'
 'Company'
 'Department'
 'departmentNumber'
 'Description'
 'EmployeeID'
 'employeeType'
 'extensionAttribute1'
 'gecos'
 'GivenName'
 'HomePage'
 'info'
 'initials'
 'LastLogonDate'
 'middlename'
 'physicalDeliveryOfficeName'
 'sn'
 'Title'
 'WhenCreated'
)
$aDStaffData = Get-ADStaffData $ActiveDirectorySearchBase $aDProperties

$empSQLInstance = Connect-DbaInstance -SqlInstance $EmployeeServer -Database $EmployeeDatabase -SqlCredential $EmployeeCredential

# $sqlParams = @{
#  Server     = $EmployeesServer
#  Database   = $EmployeesDatabase
#  Credential = $EmployeesCredential
# }

Get-EmployeeData $empSQLInstance |
 New-Obj |
  Skip-Ids $SkipPersonIds |
   Set-ADData $aDStaffData |
    Set-SiteData $empSQLInstance |
     Set-StaleSubStatus $GracePeriodMonths |
      Set-ClearExpiration |
       Set-Description |
        Set-PropertyListData |
         Clear-ADExpireDate |
          Update-ADAttributes |
           Enable-ADAccount |
            Show-Object

Clear-SessionData
if ($WhatIf) { Show-TestRun }
Show-BlockInfo end