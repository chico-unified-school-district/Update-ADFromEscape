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

function Add-Description {
 process {
  $_.desc = if ($_.site.siteAbbrv -match '[A-Za-z]') {
   ($_.site.siteAbbrv + ' ' + $_.emp.JobClassDescr) -replace '\s+', ' '
  }
  elseif ($_.emp.JobClassDescr -match '[A-Za-z]') { $_.emp.JobClassDescr }
  else {
   $_.ad.Description
  }
  $_
 }
}

function Add-ExpirationDate {
 process {
  $adExpDate = $_.ad.AccountExpirationDate
  # Skip student teacher account expiration changes
  $_.expirationDate = if ($_.ad.Description -like '*student*teacher*') { 'skip' }
  # Skip already expiring/expired stale subs
  elseif ($_.staleSub -and ($adExpDate -is [datetime])) { 'skip' }
  # Expire Stale Sub
  elseif ($_.staleSub -and ($adExpDate -isnot [datetime])) { (Get-Date).AddDays(14) }
  # Clear expire date for recently returned staff, as HR has se the employee to Active status.
  # Office staff can enable these accounts and reset passwords as needed.
  elseif (!$_.isSub -and ($adExpDate -is [datetime]) -and ($adExpDate -lt (Get-Date)) ) { $null }
  else { 'skip' }
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

function Get-ADData ($ou, $properties) {
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
  [PSCustomObject]@{
   ad             = $null
   emp            = $_
   site           = $null
   desc           = $null
   propertyList   = $null
   isSub          = $null
   staleSub       = $null
   expirationDate = $null
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

function Set-SubStatus {
 begin {
  # 'name,description' | Out-File .\stale-subs.csv
 }
 process {
  # Skip non-sub accounts
  if ($_.emp.EmploymentStatusCode -notmatch 'S') { return $_ }
  $_.isSub = $true
  $pastDate = (Get-Date).AddMonths(-6)
  # Skip newer, unused sub accounts
  if ($_.WhenCreated -gt $pastDate -and $_.LastLogonDate -isnot [datetime]) { return $_ }
  # Skip subs that have used their account with the 6 month grace period
  if ($_.ad.LastLogonDate -and ($_.ad.LastLogonDate -gt $pastDate)) { return $_ }
  Write-Verbose ('{0},{1},Stale Sub Account Detected!' -f $MyInvocation.MyCommand.Name, $_.ad.SamAccountName)
  $_.staleSub = $true
  # $_.ad.Name + ',' + $_.ad.description | Out-File .\stale-subs.csv -Append
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
    $msgVars = $MyInvocation.MyCommand.Name, $_.ad.SamAccountName, $propName, $_.ad.$propName, $propValue
    if ($propValue) {
     # $propValue = if ($propValue -match '[A-Za-z0-9]') { $propValue.Trim() } else { $propValue }
     Write-Host ('{0},{1},{2},[{3}] => [{4}]' -f $msgVars) -Fore Blue
     Set-ADUser -Identity $_.ad.ObjectGUID -Replace @{$propName = $propValue } -WhatIf:$WhatIf
    }
    else {
     if ($clearProps -notcontains $propName) { return $_ }
     Write-Host ('{0},{1},{2},[{3}] => [{4}]' -f $msgVars) -Fore Blue
     Set-ADUser -Identity $_.ad.ObjectGUID -Clear $propName -WhatIf:$WhatIf
    }
   }
  }
  $_
 }
}

function Update-ADExpireDate {
 process {
  if ($_.expirationDate -eq 'skip') { return $_ }
  $msg = if ($_.staleSub) { 'Stale Sub Account - Adding Expire Date' } else { 'Clearing Expire Date' }
  Write-Host ('{0},{1},{2},Emp Status: {3}' -f $MyInvocation.MyCommand.Name, $_.ad.SamAccountName, $msg, $_.emp.EmploymentStatusCode) -f DarkCyan
  $expireDate = if ($_.staleSub) { (Get-Date).AddDays(14) } else { $null }
  Set-ADUser -Identity $_.ad.ObjectGUID -AccountExpirationDate $expireDate -Confirm:$false -WhatIf:$WhatIf
  Read-Host '============================================='
 }
}


# =======================================================================================
Import-Module -Name CommonScriptFunctions
Import-Module -Name dbatools -Cmdlet 'Invoke-DbaQuery', 'Set-DbatoolsConfig', 'Connect-DbaInstance', 'Disconnect-DbaInstance'

Show-BlockInfo main
if ($WhatIf) { Show-TestRun }

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
 'Company'
 'Department'
 'departmentNumber'
 'Description'
 'EmployeeID'
 'employeeType'
 'extensionAttribute1'
 'gecos'
 'GivenName'
 'initials'
 'middlename'
 'physicalDeliveryOfficeName'
 'sn'
 'Title'
 'AccountExpirationDate'
 'LastLogonDate'
 'WhenCreated'
)
$ADData = Get-ADData $ActiveDirectorySearchBase $aDProperties

Get-EmployeeData | New-Obj | Skip-Ids $SkipPersonIds | Add-ADData $ADData |
 Add-SiteData $siteRefDate |
  Add-Description |
   Add-PropertyListData |
    Update-ADAttributes |
     Set-SubStatus |
      Add-ExpirationDate |
       # Update-ADExpireDate |
       Show-Object

if ($WhatIf) { Show-TestRun }
Show-BlockInfo end