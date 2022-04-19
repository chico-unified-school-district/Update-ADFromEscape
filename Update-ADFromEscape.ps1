#Requires -Version 5.0
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
  A log file is generated when using -Log [switch]
.NOTES
#>

[cmdletbinding()]
param (
 [Parameter(Mandatory = $True)]
 [Alias('DC')]
 [ValidateScript( { Test-Connection -ComputerName $_ -Quiet -Count 3 })]
 [string]$DomainController,
 [Parameter(Mandatory = $True)]
 [Alias('ADCred')]
 [System.Management.Automation.PSCredential]$ADCredential,
 # String formatted as 'server\database'
 [Parameter(Mandatory = $True)]
 [Alias('SearchBase')]
 [string]$ActiveDirectorySearchBase,
 [Parameter(Mandatory = $True)]
 [ValidateScript( { Test-Connection -ComputerName $_ -Quiet -Count 3 })]
 [string]$SQLServer,
 [Parameter(Mandatory = $True)]
 [string]$SQLDatabse,
 [Parameter(Mandatory = $True)]
 [System.Management.Automation.PSCredential]$SQLCredential,
 [Alias('wi')]
 [SWITCH]$WhatIf
)

function Compare-Data ($escapeData, $adData, $properties) {
 Write-Verbose $MyInvocation.MyCommand.name
 Write-Verbose ('{0},Escape Count: {1}, AD Count {2}' -f $MyInvocation.MyCommand.name, $escapeData.count, $adData.count )
 $compareParams = @{
  ReferenceObject  = $escapeData
  DifferenceObject = $adData
  Property         = $properties
  Debug            = $false
 }
 $results = Compare-Object @compareParams | Where-Object { ($_.sideindicator -eq '=>') }
 # $output = foreach ($item in $results) { $AeriesData.Where({ $_.employeeId -eq $item.employeeId }) }
 Write-Verbose ( '{0},Count: {1}' -f $MyInvocation.MyCommand.name, $results.count)
 $results
}
filter Find-DuplicateIds {
 $id = $_.employeeId
 $adObj = $global:adData.Where({ $_.employeeId -eq $id })
 if ($adObj.count -gt 1) {
  Write-Warning ('{0},{1},Multiple AD objects detected' -f $MyInvocation.MyCommand.Name, $id)
  Start-Sleep 20
  return
 }
 $_
}
filter Find-ActiveEscapeUser {
 $id = $_.employeeId
 $escObj = $global:escapeData.Where({ $_.employeeId -eq $id })
 if (-not$escObj) {
  Write-Host ('{0},{1},Not active in Escape' -f $MyInvocation.MyCommand.Name, $id) -Fore Yellow
  return
 }
 $_
}
function Get-ADData ($properties) {
 Write-Verbose ('{0}' -f $MyInvocation.MyCommand.Name)
 $adParams = @{
  Filter     = 'mail -like "*@*" -and Enabled -eq $true'
  SearchBase = $ActiveDirectorySearchBase
  Properties = $properties
 }
 Get-ADuser @adParams | Where-Object { $_.employeeId -match '^\d{4,5}$' }
}
function Get-EscapeData {
 Write-Verbose ('{0}' -f $MyInvocation.MyCommand.Name)
 $sql = Get-Content .\sql\active-employees.sql -Raw
 $sqlParams = @{
  Server     = $SQLServer
  Database   = $SQLDatabse
  Credential = $SQLCredential
  Query      = $sql
 }
 Invoke-Sqlcmd @sqlParams | ConvertTo-Csv | ConvertFrom-Csv
}
function Update-ADAttributes {
 begin {
  $count = 1
 }
 process {
  $id = $_.employeeId
  $adObj = $global:adData.Where({ $_.employeeId -eq $id })
  $escObj = $global:escapeData.Where({ $_.employeeId -eq $id })
  Write-Verbose ('{0},{1}' -f $MyInvocation.MyCommand.Name, $adObj.SamAccountName)
  foreach ($prop in $global:escapeRowNames) {
   # Begin parse the db column names
   $propData = $escObj."$prop"
   if ( $propData -match '[A-Za-z0-9]') {
    # Begin case-sensitive compare data between AD and DB
    if ( $adObj."$prop" -cnotcontains $propData ) {
     Write-Host ("{0},{1},[{2}] => [{3}]" -f $adObj.SamAccountName, $prop, $($adObj."$prop"), $propData) -Fore Blue
     Write-Debug 'Set?'
     Set-ADUser -Identity $adObj.ObjectGUID -Replace @{$prop = $propData } -WhatIf:$WhatIf
    } # End  compare data between AD and DB
   }
   else {
    Write-Verbose ("{0},{1},No Escape property data" -f $adObj.SamAccountName, $prop)
   }
  }
  # Write-Debug 'ok'
  Write-Verbose $count
  $count++
 }
}
function Start-ADSession {
 $adSession = New-PSSession -ComputerName $DomainController -Credential $ADCredential
 Import-PSSession -Session $adSession -Module ActiveDirectory -AllowClobber | Out-Null
}
function Start-CheckUserInfo {
 Write-Host ('{0}' -f $MyInvocation.MyCommand.Name)
 if ($WhatIf) { Show-TestRun }
 Clear-SessionData
 'SQLServer' | Load-Module
 Start-ADSession

 $global:escapeData = Get-EscapeData
 $global:escapeRowNames = ($global:escapeData | Get-Member -MemberType Properties).name

 # Start-ADSession
 $global:adData = Get-ADData $global:escapeRowNames

 $results = Compare-Data $global:escapeData $global:adData $global:escapeRowNames
 $results | Find-DuplicateIds | Find-ActiveEscapeUser | Update-ADAttributes
 Clear-SessionData
 if ($WhatIf) { Show-TestRun }
}

# imported
. .\lib\Clear-SessionData.ps1
. .\lib\Load-Module.ps1
. .\lib\Show-TestRun.ps1
# process
Start-CheckUserInfo