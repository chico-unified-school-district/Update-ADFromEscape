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
 [Parameter(Mandatory = $True)][string[]]$DomainControllers,
 [Parameter(Mandatory = $True)][System.Management.Automation.PSCredential]$ADCredential,
 [Parameter(Mandatory = $True)][Alias('SearchBase')][string]$ActiveDirectorySearchBase,
 [Parameter(Mandatory = $True)][string]$EmployeeServer,
 [Parameter(Mandatory = $True)][string]$EmployeeDatabase,
 [Parameter(Mandatory = $True)][System.Management.Automation.PSCredential]$EmployeeCredential,
 [Parameter(Mandatory = $True)][string]$SiteRefServer,
 [Parameter(Mandatory = $True)][string]$SiteRefDatabase,
 [Parameter(Mandatory = $True)][System.Management.Automation.PSCredential]$SiteRefCredential,
 [Parameter(Mandatory = $True)][string]$SiteRefTable,
 [Parameter(Mandatory = $True)][int]$GracePeriodMonths,
 [int[]]$SkipPersonIds,
 [switch]$Wait,
 [Alias('wi')][SWITCH]$WhatIf
)

function Clear-ADExpireDate {
 process {
  if (!$_.clearExpiration) { return $_ }
  Write-Host ('{0},{1},Clearing AD account expiration date' -f $MyInvocation.MyCommand.Name, $_.userInfo) -f DarkCyan
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

function Get-EmployeeData ($instance) {
 $sql = Get-Content .\sql\active-employees.sql -Raw
 $data = New-SqlOperation -Server $instance -Query $sql | ConvertTo-Csv | ConvertFrom-Csv
 Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, @($data).count) -f Green
 if ($WhatIf) { Start-Sleep -Seconds 2 } # A slight pause on test runs to allow time to read the count before processing begins.
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
  # Write-Verbose ($MyInvocation.MyCommand.name, $obj | Out-String)
  $obj
 }
}

function Set-ADData ($ou, $properties) {
 begin {
  $adData = Get-ADUser -Filter "EmployeeId -like '*' -and Mail -like '*@*'" -SearchBase $ou -Properties $properties
 }
 process {
  $empId = $_.emp.EmpId
  $_.ad = $adData.Where({ $_.EmployeeId -eq $empId })
  if (!$_.ad) { return }
  $_
 }
}

function Set-ClearExpiration {
 process {
  $_.clearExpiration = switch ($_) {
   { $_.ad.AccountExpirationDate -isnot [datetime] } { break } # No need to clear empty value
   { $_.emp.PersonTypeId -eq '6' } { break } # No need to clear student teacher accounts
   { $_.staleSub } { break } # No need to clear stale subs
   default { $true }
  }
  $_
 }
}

function Set-Description {
 process {
  if ($_.staleSub) { return $_ }
  $jobData = (($_.site.siteAbbrv + ' ' + $_.emp.JobClassDescr) -replace '\s+', ' ')
  $_.desc = switch ($_.emp.EmploymentStatusCode) {
   { $_ -match 'S' } { 'Substitute ' + $jobData; break }
   { $_ -match 'A' } { $jobData; break }
   default { 'Chico Unified Employee Account' }
  }
  return $_
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
  # Write-Verbose ($MyInvocation.MyCommand.name, $_ | Out-String)
  $_
 }
}

function Set-SiteData ($instance, $table) {
 begin {
  $sql = 'SELECT * FROM {0}' -f $table
  $siteData = New-SqlOperation -Server $instance -Query $sql | ConvertTo-Csv | ConvertFrom-Csv
  # Write-Verbose ($MyInvocation.MyCommand.Name, $siteData | Out-String)s
 }
 process {
  if ($_.emp.SiteId -notmatch '\d') { return $_ }
  $siteId = $_.emp.SiteId
  $_.site = $siteData.Where({ [int]$_.SiteCode -eq [int]$siteId })
  if (!$_.site) { Write-Host ('{0},{1},Site not found for SiteId {2}' -f $MyInvocation.MyCommand.Name, $_.userInfo, $siteId) -f Red }
  $_
 }
}

function Set-StaleSubStatus ([int]$months) {
 begin { $cutOffDate = (Get-Date).AddMonths(-$months) }
 process {
  if ($_.emp.EmploymentStatusCode -notmatch 'S') { return $_ } # Skip non-subs
  $lastUsed = if ($_.ad.LastLogonDate) { $_.ad.LastLogonDate } else { $_.ad.WhenCreated }
  if ($lastUsed -le $cutOffDate) { $_.staleSub = $true }
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
  if ($Wait) { Read-Host ('{0}' -f ('x' * 50)) }
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
  $clearProps = 'extensionAttribute1' # Only clear these attributes when cleared in Escape. Added by request.
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
Clear-Host
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

$empSQLInstance = Connect-DbaInstance -SqlInstance $EmployeeServer -Database $EmployeeDatabase -SqlCredential $EmployeeCredential
$intSQLInstance = Connect-DbaInstance -SqlInstance $SiteRefServer -Database $SiteRefDatabase -SqlCredential $SiteRefCredential

Get-EmployeeData $empSQLInstance |
 New-Obj |
  Skip-Ids $SkipPersonIds |
   Set-ADData -ou $ActiveDirectorySearchBase -properties $aDProperties |
    Set-SiteData $intSQLInstance $SiteRefTable |
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