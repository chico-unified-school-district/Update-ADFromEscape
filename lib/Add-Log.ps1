function Add-Log {
	[cmdletbinding()]
	Param ( 
		[Parameter(Position=0,Mandatory=$True)]
		[STRING]$Type,
  [Parameter(Position=1,Mandatory=$True)]
  [Alias("Msg")]
		[STRING]$Message,
  [Parameter(Position=2,Mandatory=$false)]
  [SWITCH]$WhatIf )
 $date = Get-Date -Format s
 $type = "[$($type.toUpper())]"
 $testString = if ($WhatIf){"[WhatIf] "}
 $logMsg = "$testString$date,$type,$message"
 Write-Output $logMsg
}