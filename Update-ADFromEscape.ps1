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
  [Parameter(Mandatory = $True)]
  [Alias('DCs')]
  [string[]]$DomainControllers,
  [Parameter(Mandatory = $True)]
  [Alias('ADCred')]
  [System.Management.Automation.PSCredential]$ADCredential,
  # String formatted as 'server\database'
  [Parameter(Mandatory = $True)]
  [Alias('SearchBase')]
  [string]$ActiveDirectorySearchBase,
  [Parameter(Mandatory = $True)]
  [string]$SQLServerEmployees,
  [Parameter(Mandatory = $True)]
  [string]$SQLDatabaseEmployees,
  [Parameter(Mandatory = $True)]
  [System.Management.Automation.PSCredential]$SQLCredentialEmployees,
  [Parameter(Mandatory = $True)]
  [string]$SQLServerSiteRef,
  [Parameter(Mandatory = $True)]
  [string]$SQLDatabaseSiteRef,
  [Parameter(Mandatory = $True)]
  [System.Management.Automation.PSCredential]$SQLCredentialSiteRef,
  [Alias('wi')]
  [SWITCH]$WhatIf
)

# filter Find-DuplicateIds {
#   $id = $_.employeeId
#   $adObj = $adData.Where({ $_.employeeId -eq $id })
#   if ($adObj.count -gt 1) {
#     Write-Warning ('{0},{1},Multiple AD objects detected' -f $MyInvocation.MyCommand.Name, $id)
#     Start-Sleep 20
#     return
#   }
#   $_
# }

function Get-ADData ($properties) {
  Write-Verbose ('{0}' -f $MyInvocation.MyCommand.Name)
  $adParams = @{
    Filter     = 'mail -like "*@*" -and
    Enabled -eq $true -and
    EmployeeID -like "*"
    '
    SearchBase = $ActiveDirectorySearchBase
    Properties = $properties
  }
  Get-ADuser @adParams | Where-Object { $_.employeeId -match '^\d{4,5}$' }
}

function Get-EscapeData {
  Write-Verbose ('{0}' -f $MyInvocation.MyCommand.Name)
  $sqlParams = @{
    Server                 = $SQLServerEmployees
    Database               = $SQLDatabaseEmployees
    Credential             = $SQLCredentialEmployees
    TrustServerCertificate = $true
    Query                  = (Get-Content .\sql\active-employees.sql -Raw)
  }
  Invoke-Sqlcmd @sqlParams | ConvertTo-Csv | ConvertFrom-Csv
}

function Get-SiteRefData {
  $params = @{
    Server                 = $SQLServerSiteRef
    Database               = $SQLDatabaseSiteRef
    Credential             = $SQLCredentialSiteRef
    TrustServerCertificate = $true
    Query                  = 'SELECT * FROM zone_ref;'
  }
  Invoke-Sqlcmd @params
}

function Get-Site ($siteRef, $siteCode) {
  $siteRef.Where({ $_.siteCode -eq $siteCode })
}

function New-PropertyObj {
  begin {
    $siteRef = Get-SiteRefData
  }
  process {
    $siteData = Get-Site $siteRef $_.SiteID
    $desc = if ($siteData -and $siteData.siteAbbrv.length -gt 1) {
   ($siteData.siteAbbrv + ' ' + $_.JobClassDescr) -replace '\s+', ' '
    }
    else {
      $_.JobClassDescr
    }
    $initial = if ($_.NameMiddle -match '[A-Za-z]') { $_.NameMiddle.SubString(0, 1) }
    # $name = ($_.NameFirst + ' ' + $_.NameLast) -replace ('\s+', ' ')
    $obj = [PSCustomObject]@{
      givenname                  = $_.NameFirst
      sn                         = $_.NameLast
      middlename                 = $_.NameMiddle
      initials                   = $initial
      # displayName                = $name
      employeeID                 = $_.EmpID
      company                    = 'Chico Unified School District'
      title                      = $_.JobClassDescr
      description                = $desc
      department                 = $_.JobCategoryDescr
      departmentnumber           = $_.SiteID
      physicalDeliveryOfficeName = $_.SiteDescr
      extensionAttribute1        = $_.BargUnitID
    }
    $obj
  }
}

function Update-ADAttributes ($adData, $properties) {
  begin {
    $count = $adData.count
    function Remove-ExtraSpaces ($string) {
      if ($string -match '\s{2,}') {
        $string -replace '\s+', ' '
      }
    }
  }
  process {
    $id = $_.EmployeeID
    $adObj = $adData.Where({ $_.EmployeeID -eq $id })
    if (-not$adObj) { return }
    Write-Verbose ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.EmployeeID, $adObj.name)
    Write-Verbose ($_ | Out-String)
    # fixed double spaces in name/displayName
    $fixedName = Remove-ExtraSpaces $adObj.name
    if ($fixedName) {
      $msgVars = $MyInvocation.MyCommand.Name, $adObj.name, $fixedName
      Write-Host ('{0},Removing extra spaces: [{1}] => [{2}]' -f $msgVars) -Fore Green
      Rename-ADObject -Identity $adObj.ObjectGUID -NewName $fixedName -WhatIf:$WhatIf
      Set-Aduser -Identity $adObj.ObjectGUID -DisplayName $fixedName -WhatIf:$WhatIf
    }
    foreach ($prop in $properties) {
      $propData = $_.$prop
      # Ensure property has data
      if ( $propData -match '[A-Za-z0-9]') {
        # Begin case-sensitive compare data between AD and DB
        if ( $adObj.$prop -cnotcontains $propData ) {
          $msgVars = $count, $MyInvocation.MyCommand.Name, $id, $adObj.SamAccountName, $prop, $($adObj.$prop), $propData
          Write-Host ("{0},{1},{2},{3},{4},[{5}] => [{6}]" -f $msgVars) -Fore Blue
          Write-Debug 'Set?'
          Set-ADUser -Identity $adObj.ObjectGUID -Replace @{$prop = $propData } -WhatIf:$WhatIf
        }
      }
      else {
        Write-Verbose ("{0},{1},No Escape property data" -f $adObj.SamAccountName, $prop)
      }
    }
    Write-Debug 'ok'
    $count--
  }
}

function Start-ADSession {
  $dc = Select-DomainController $DomainControllers
  $cmdlets = 'Get-ADuser', 'Set-ADuser', 'Rename-ADObject', 'Clear-ADAccountExpiration'
  $adSession = New-PSSession -ComputerName $dc -Credential $ADCredential
  Import-PSSession -Session $adSession -Module ActiveDirectory -CommandName $cmdlets -AllowClobber | Out-Null
}


function Update-AccountExpiration {
  process {
    $id = $_.EmployeeID
    $adObj = Get-ADUser -Filter "EmployeeId -eq $id" -Properties *
    if ($adObj) {
      $msgVars = $MyInvocation.MyCommand.Name, $id, $adObj.SamAccountName
      if (($adObj.AccountExpirationDate -is [datetime]) -and ($adObj.AccountExpirationDate -gt (Get-Date))) {
        Write-Host ('{0},{1},{2},Clearing Expiration Date' -f $msgVars) -Fore DarkCyan
        # Clear-ADAccountExpiration -Identity $adObj.ObjectGUID -Confirm:$false -WhatIf:$WhatIf
      }
    }
  }
}

# ======================== Main ===========================
# Imported
. .\lib\Clear-SessionData.ps1
. .\lib\Load-Module.ps1
. .\lib\Select-DomainController.ps1
. .\lib\Show-TestRun.ps1

Show-TestRun
Clear-SessionData

Start-ADSession

$aDProperties = @(
  'givenname'
  'sn'
  'middlename'
  'initials'
  'employeeID'
  'company'
  'title'
  'description'
  'department'
  'departmentnumber'
  'physicalDeliveryOfficeName'
  'extensionAttribute1'
)
# $aDProperties = $userObjects[0].PSObject.Properties.Name
$adData = Get-ADData $aDProperties
$userObjects = Get-EscapeData | New-PropertyObj
$userObjects | Update-ADAttributes -adData $adData -properties $aDProperties
$userObjects | Update-AccountExpiration