
[cmdletbinding()]
param(
 [Parameter(Mandatory = $True)]
 [Alias('DC')]
 [string]$DomainController,
 [Parameter(Mandatory = $True)]
 [Alias('ADCred')]
 [System.Management.Automation.PSCredential]$ADCredential,
 [Parameter(Mandatory = $True)]
 [string]$SQLServer,
 [Parameter(Mandatory = $True)]
 [string]$SQLDatabse,
 [Parameter(Mandatory = $True)]
 [System.Management.Automation.PSCredential]$SQLCredential
)

if ([int]$iteCode -ge 500) {
 $ouLookup
}
