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
  function Test-Null ($obj) { if ($obj -match '[A-Za-z0-9]') { $obj } else { $null } }
 }
 process {
  $initials = if ($_.emp.NameMiddle -match '\w') { $_.emp.NameMiddle.SubString(0, 1) }
  $_.propertyList = [PSCustomObject]@{
   Company                    = 'Chico Unified School District'
   Department                 = Test-Null $_.emp.JobCategoryDescr
   departmentNumber           = Test-Null $_.emp.siteID
   Description                = Test-Null $_.desc
   extensionAttribute1        = Test-Null $_.emp.BargUnitID
   GivenName                  = Remove-ExtraSpaces $_.emp.NameFirst
   initials                   = $initials
   middleName                 = Test-Null (Remove-ExtraSpaces $_.emp.NameMiddle)
   physicalDeliveryOfficeName = Test-Null $_.site.SiteDesc
   sn                         = Remove-ExtraSpaces $_.emp.NameLast
   Title                      = Test-Null $_.emp.JobClassDescr
   AccountExpirationDate      = $null
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
   ad           = $null
   emp          = $_
   site         = $null
   desc         = $null
   propertyList = $null
  }
 }
}

function Show-Object {
 process {
  Write-Verbose ($MyInvocation.MyCommand.name, $_ | Out-String)
  Read-Host ('{0}' -f ('x' * 50))
 }
}

function Skip-Ids ([int[]]$ids) {
 process {
  # Skip specific PersonTypeIds. This was done to preserve ad info for student workers
  if ($ids -contains [int]$_.emp.PersonTypeId) { return }
  $_
 }
}

function Update-ADAttributes {
 process {
  foreach ($propName in $_.propertyList.PSObject.Properties.Name) {
   $propValue = $_.propertyList.$propName
   # Write-Host ('{0},{1},{2},{3}' -f $MyInvocation.MyCommand.Name, $_.ad.SamAccountName, $propName, $propValue) -F Green
   # Begin case-sensitive compare data between AD and DB.
   if ( $_.ad.$propName -cnotcontains $propValue) {
    $msgVars = $MyInvocation.MyCommand.Name, $_.ad.SamAccountName, $propName, $_.ad.$propName, $propValue
    Write-Host ('{0},{1},{2},[{3}] => [{4}]' -f $msgVars) -Fore Blue
    if (!$propValue) {
     Set-ADUser -Identity $_.ad.ObjectGUID -Clear $propName -WhatIf:$WhatIf
    }
    else {
     Set-ADUser -Identity $_.ad.ObjectGUID -Replace @{$propName = $propValue } -WhatIf:$WhatIf
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
 'extensionAttribute1'
 'GivenName'
 'initials'
 'middlename'
 'physicalDeliveryOfficeName'
 'sn'
 'Title'
 'AccountExpirationDate'
)
$ADData = Get-ADData $ActiveDirectorySearchBase $aDProperties

Get-EmployeeData | New-Obj | Skip-Ids $SkipPersonIds | Add-ADData $ADData |
 Add-SiteData $siteRefDate |
  Add-Description |
   Add-PropertyListData |
    Update-ADAttributes |
     Show-Object

if ($WhatIf) { Show-TestRun }
Show-BlockInfo end