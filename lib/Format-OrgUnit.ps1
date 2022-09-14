
[cmdletbinding()]
param(
 [Parameter(Mandatory = $True)]
 [Alias('DC')]
 [string]$DomainController,
 [Parameter(Mandatory = $True)]
 [Alias('ADCred')]
 [System.Management.Automation.PSCredential]$ADCredential,
 [Parameter(Mandatory = $True)]
 [string]$EmployeeOrgUnit,
 [Parameter(Mandatory = $True)]
 [string]$SQLRefServer,
 [Parameter(Mandatory = $True)]
 [string]$SQLRefDatabse,
 [Parameter(Mandatory = $True)]
 [System.Management.Automation.PSCredential]$SQLRefCredential,
 [Parameter(Mandatory = $True)]
 $EmployeeObj
)

function Add-SiteInfo {
 process {
  $siteData = Get-Site $siteRef $_.siteCode
  if (-not$siteData) {
   $msgVars = $MyInvocation.MyCommand.Name, $_.empId, $_.siteCode
   Write-Host ('{0},{1},{2},No Site Data' -f $msgVars)
   return
  }
  $_ | Add-Member -MemberType NoteProperty -Name siteData -Value $siteData
  $_
 }
}

function Get-UserOrgUnit {
 process {
  # Site codes 500 and under are school sites
  if ($_.siteCode -ge 500) {
   Get-ServiceOrgUnit
  }
  else {
   $_ | Get-SiteOrgUnit
  }
 }
}

function Get-ServiceOrgUnit {
 $ouName = $_.siteDescription
 Get-ADOrganizationalUnit -Filter * -SearchBase $EmployeeOrgUnit |
 Where-Object { $_.DistinguishedName -match "^OU=$ouName," }
}

function Add-OrgLevel {
 begin {
  $orgLevels = @{CSEA = 'Admin'; CUMA = 'Admin'; CUTA = 'Teacher' }
 }
 process {
  $myOrgLevel = $orgLevels[$_.BargUnitId]
  if (-not$myOrgLevel) {
   $msgVars = $MyInvocation.MyCommand.Name, $_.empId, $_.BargUnitId
   Write-Host ('{0},{1},{2},No matching BargUnitId. No org level set.' -f $msgVars)
   return
  }
  $_ | Add-Member -MemberType NoteProperty -Name orgLevel -Value $myOrgLevel
  $_
 }
}

function Get-Site ($siteRef, $siteCode) {
 $siteRef.Where({ $_.siteCode -eq $siteCode })
}

# function Get-SiteOrgUnit {
#  begin {
#   $employeeOUs = Get-ADOrganizationalUnit -filter * -SearchBase $EmployeeOrgUnit
#  }
#  process {
#   # $ouLevel = $_.BargUnitId | Set-OrgUnitLevel
#   # $ouName = $_.siteCode | Set-OrgUnitName
#   if ($null -eq $ouLevel -or $null -eq $ouSite) { return }
#   $employeeOus.Where({ ($_.DistinguishedName -match "^OU=$ouName,") -and
#   ($_.DistinguishedName -match "\bOU=$ouLevel,") })
#  }
# }

function Get-OrgUnit {
 begin {
  $employeeOUs = Get-ADOrganizationalUnit -filter * -SearchBase $EmployeeOrgUnit
 }
 process {
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.empId, $_.siteCode)
  $orgUnit = $_.siteData.orgUnit
  $ouLevel = $_.orgLevel
  $targetOrgUnit = if ($_.siteCode -ge 500){
   $employeeOus.Where({ ($_.DistinguishedName -match "^OU=$orgUnit,")})
  } else {
   $employeeOus.Where({ ($_.DistinguishedName -match "^OU=$orgUnit,") -and
   ($_.DistinguishedName -match "\bOU=$ouLevel,") })
  }
  $targetOrgUnit.DistinguishedName
 }
}

function Get-SiteRefData {
 $siteRefSQLParams = @{
  Server     = $SQLRefServer
  Database   = $SQLRefDatabse
  Credential = $SQLRefCredential
  Query      = 'SELECT * FROM zone_ref;'
 }
 Invoke-Sqlcmd @siteRefSQLParams
}

$siteRef = Get-SiteRefData
$EmployeeObj | Add-SiteInfo | Add-OrgLevel | Get-OrgUnit