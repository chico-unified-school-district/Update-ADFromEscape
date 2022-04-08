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
 [string]$SQLDatabse,
 # Credential object with database select permission.
 [Parameter(Mandatory = $True)]
 [Alias('SQLCred')]
 [System.Management.Automation.PSCredential]$SQLCredential,
 [Alias('wi')]
 [SWITCH]$WhatIf
)
