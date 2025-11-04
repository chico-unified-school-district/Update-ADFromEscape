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
 [Parameter(Mandatory = $True)][string]$EmployeesServer,
 [Parameter(Mandatory = $True)][string]$EmployeesDatabase,
 [Parameter(Mandatory = $True)][System.Management.Automation.PSCredential]$EmployeesCredential,
 [Parameter(Mandatory = $True)][string]$SiteRefServer,
 [Parameter(Mandatory = $True)][string]$SiteRefDatabase,
 [Parameter(Mandatory = $True)][System.Management.Automation.PSCredential]$SiteRefCredential,
 [Parameter(Mandatory = $True)][string]$SiteRefTable,
 [Parameter(Mandatory = $True)][int]$GracePeriodMonths,
 [int[]]$SkipPersonIds,
 [Alias('wi')][SWITCH]$WhatIf
)

function Add-ADData ($data) {
 process {
  $id = $_.emp.EmpID
  $_.ad = $data.Where({ $_.EmployeeID -eq $id })
  if (!$_.ad) { return }
  $_
 }
}

function Add-GSuiteLicense {
 begin {
  $licenses = @(
   1010310009 # Google Workspace for Education Plus (Staff)
  )
 }
 process {
  $gSuite = $_.ad.HomePage
  foreach ($license in $licenses) {
   Write-Host ('{0},{1},Adding GSuite License {2}' -f $MyInvocation.MyCommand.Name, $_.userInfo, $license) -f DarkCyan
   $ErrorActionPreference = 'Continue'
   if (!$WhatIf) { & $gam user $gSuite add license $license }
   $ErrorActionPreference = 'Stop'
  }
  $_
 }
}

function Add-Description {
 process {
  $_.desc = if ($_.ad.AccountExpirationDate -is [datetime]) { $_.ad.Description } # No Change for expiring accounts
  elseif ($_.site.siteAbbrv -match '[A-Za-z]') { ($_.site.siteAbbrv + ' ' + $_.emp.JobClassDescr) -replace '\s+', ' ' }
  elseif ($_.emp.JobClassDescr -match '[A-Za-z]') { $_.emp.JobClassDescr }
  else { $_.ad.Description }
  $_
 }
}

function Add-ClearExpiration {
 process {
  $_.clearExpiration = if
  # Expire date occurs today or in the future
  ($_.ad.AccountExpirationDate -ge [DateTime]::Today ) { $false }
  # Student teacher accounts with expiration
  elseif (($_.ad.AccountExpirationDate -is [datetime]) -and ($_.ad.Description -like '*student*teacher*')) { $false }
  # Stale sub with expiration date already set
  elseif ($_.staleSub -and ($_.ad.AccountExpirationDate -is [datetime])) { $false }
  # All others
  elseif ($_.ad.AccountExpirationDate -isnot [datetime]) { $false }
  # What about already expired stale subs?
  else { $true } # clear expire date
  $_
 }
}

function Add-SiteData ($refData) {
 process {
  if ($_.emp.SiteID -notmatch '\d') { return $_ }
  $siteID = $_.emp.SiteID
  $_.site = $refData.Where({ [int]$_.SiteCode -eq [int]$siteID })
  $_
 }
}

function Clear-ADExpireDate {
 process {
  if (!$_.clearExpiration) { return $_ }
  $msg = $MyInvocation.MyCommand.Name, $_.userInfo
  Write-Host ('{0},{1},Clearing expire date' -f $msg) -f DarkCyan
  Set-ADUser -Identity $_.ad.ObjectGUID -AccountExpirationDate $null -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Get-ADActiveStaff ($ou, $properties) {
 Write-Verbose ('{0}' -f $MyInvocation.MyCommand.Name)
 $adParams = @{
  Filter     = 'mail -like "*@*" -and
    Enabled -eq $true -and
    EmployeeID -like "*"
    '
  SearchBase = $ou
  Properties = $properties
 }
 Get-ADUser @adParams | Where-Object { $_.EmployeeID -notmatch '[A-Za-z]' }
}

function Add-PropertyListData {
 begin {
  function Remove-ExtraSpaces ($string) { $string -replace '\s+', ' ' }
  function Test-Null ($obj) { if ($obj -match '[A-Za-z0-9]') { $obj.Trim() } else { $null } }
 }
 process {
  $initials = if ($_.emp.NameMiddle -match '\w') { $_.emp.NameMiddle.SubString(0, 1) }
  $_.propertyList = [PSCustomObject]@{
   Company                    = 'Chico Unified School District'
   Department                 = Test-Null $_.emp.JobCategoryDescr
   departmentNumber           = Test-Null $_.emp.siteID
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
  $_
 }
}

function Get-EmployeeData {
 $sqlParams = @{
  Server     = $EmployeesServer
  Database   = $EmployeesDatabase
  Credential = $EmployeesCredential
  Query      = (Get-Content .\sql\active-employees.sql -Raw)
 }
 $data = New-SqlOperation @sqlParams | ConvertTo-Csv | ConvertFrom-Csv
 Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, @($data).count) -f Green
 $data
}

function New-Obj {
 process {
  $usrInf = $_.EmpId + ' ' + $_.NameLast + ' ' + $_.NameFirst + ' ' +
  $_.EmailWork + ' ' + $_.EmploymentStatusDescr + ' ' + $_.EmploymentStatusCode
  [PSCustomObject]@{
   ad              = $null
   clearExpiration = $null
   desc            = $null
   emp             = $_
   propertyList    = $null
   site            = $null
   staleSub        = $null
   userInfo        = $usrInf.trim()
  }
 }
}

function Show-Object {
 process {
  Write-Verbose ($MyInvocation.MyCommand.name, $_ | Out-String)
  # Read-Host ('{0}' -f ('x' * 50))
 }
}

function Skip-Ids ([int[]]$ids) {
 process {
  # Skip specific PersonTypeIds. This was done to preserve ad info for student workers
  if ($ids -contains [int]$_.emp.PersonTypeId) { return }
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
  $_
 }
}

function Update-ADAttributes {
 begin {
  $clearProps = 'extensionAttribute1'
 }
 process {
  foreach ($propName in $_.propertyList.PSObject.Properties.Name) {
   $propValue = $_.propertyList.$propName
   # Write-Host ('{0},{1},{2},{3}' -f $MyInvocation.MyCommand.Name, $_.ad.SamAccountName, $propName, $propValue) -F Green
   # Begin case-sensitive compare data between AD and DB.
   if ( $_.ad.$propName -cnotcontains $propValue) {
    $msgVars = $MyInvocation.MyCommand.Name, $_.userInfo, $propName, $_.ad.$propName, $propValue
    if ($propValue) {
     Write-Host ('{0},{1},{2},[{3}] => [{4}]' -f $msgVars) -Fore Blue
     Set-ADUser -Identity $_.ad.ObjectGUID -Replace @{$propName = $propValue } -WhatIf:$WhatIf
    }
    else {
     if ($clearProps -notcontains $propName) { return $_ }
     Write-Host ('{0},{1},{2},[{3}] => [{4}]' -f $msgVars) -Fore Cyan
     Set-ADUser -Identity $_.ad.ObjectGUID -Clear $propName -WhatIf:$WhatIf
    }
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
$gam = 'C:\GAM7\gam.exe'

$cmdlets = 'Get-ADuser', 'Set-ADuser', 'Rename-ADObject', 'Clear-ADAccountExpiration'
Connect-ADSession -DomainControllers $DomainControllers -Cmdlets $cmdLets -Credential $ADCredential

$siteRefParams = @{
 Server     = $SiteRefServer
 Database   = $SiteRefDatabase
 Credential = $SiteRefCredential
 Query      = 'SELECT * FROM {0}' -f $SiteRefTable
}
$siteRefDate = New-SqlOperation @siteRefParams | ConvertTo-Csv | ConvertFrom-Csv

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
 'initials'
 'LastLogonDate'
 'middlename'
 'physicalDeliveryOfficeName'
 'sn'
 'Title'
 'WhenCreated'
)
$ActiveADStaff = Get-ADActiveStaff $ActiveDirectorySearchBase $aDProperties

Get-EmployeeData |
 New-Obj |
  Skip-Ids $SkipPersonIds |
   Add-ADData $ActiveADStaff |
    Add-SiteData $siteRefDate |
     Add-Description |
      Add-PropertyListData |
       Update-ADAttributes |
        Set-StaleSubStatus $GracePeriodMonths |
         Add-ClearExpiration |
          Clear-ADExpireDate |
           Show-Object

if ($WhatIf) { Show-TestRun }
Show-BlockInfo end