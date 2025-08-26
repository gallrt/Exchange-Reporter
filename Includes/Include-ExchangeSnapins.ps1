# Die Funktion "Remove-WriteConsole" dient nur dazu die Ausgabe der Exchange Shell
# unterdrücken
#--------------------------------------------------------------------------------------

function Remove-WriteConsole {
	[CmdletBinding(DefaultParameterSetName = 'FromPipeline')]
	param(
		[Parameter(ValueFromPipeline = $true, ParameterSetName = 'FromPipeline')] [object] $InputObject,
		[Parameter(Mandatory = $true, ParameterSetName = 'FromScriptblock', Position = 0)] [ScriptBlock] $ScriptBlock
	)

	begin {
		function Cleanup {
			Remove-Item function:\write-host -ea 0
			Remove-Item function:\write-verbose -ea 0
		}

		function ReplaceWriteConsole([string] $Scope) {
			Invoke-Expression "function ${scope}:Write-Host { }"
			Invoke-Expression "function ${scope}:Write-Verbose { }"
		}

		Cleanup

		if ($pscmdlet.ParameterSetName -eq 'FromPipeline') {
			ReplaceWriteConsole -Scope 'global'
		}
	}

	process {
		if ($pscmdlet.ParameterSetName -eq 'FromScriptBlock') {
			. ReplaceWriteConsole -Scope 'local'
			& $scriptblock
		} else {
			$InputObject
		}
	}

	end {
		Cleanup
	}  
}

# Lade Exchange Snapins und Verbinde zu Exchange Server
#--------------------------------------------------------------------------------------

if ((Get-PSSnapin | Where-Object { $_.name }) -notmatch "Microsoft.Exchange.Management.PowerShell") {
	if ($emsversion -match "2010") {
		$repspath = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath + "bin\RemoteExchange.ps1"
		$snapins = . $repspath | Remove-WriteConsole
		$connect = Connect-ExchangeServer -auto | Remove-WriteConsole
	}
	if ($emsversion -match "2013") {
		$repspath = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath + "bin\RemoteExchange.ps1"
		$snapins = . $repspath | Remove-WriteConsole
		$connect = Connect-ExchangeServer -auto | Remove-WriteConsole
	}
	if ($emsversion -match "2016") {
		$repspath = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath + "bin\RemoteExchange.ps1"
		$snapins = . $repspath | Remove-WriteConsole
		$connect = Connect-ExchangeServer -auto | Remove-WriteConsole
	}
	if ($emsversion -match "2019") {
		$repspath = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath + "bin\RemoteExchange.ps1"
		$snapins = . $repspath | Remove-WriteConsole
		$connect = Connect-ExchangeServer -auto | Remove-WriteConsole
	}
	if ($emsversion -notmatch "2010" -and $emsversion -notmatch "2013" -and $emsversion -notmatch "2016" -and $emsversion -notmatch "2019") {
		$version = (Get-ChildItem HKLM:\SOFTWARE\Microsoft\ExchangeServer\v1* -ea 0 | Sort-Object -Descending | Select-Object -First 1).pschildname
		$repspath = (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\$version\Setup).MsiInstallPath + "bin\RemoteExchange.ps1"
		$snapins = . $repspath | Remove-WriteConsole
		$connect = Connect-ExchangeServer -auto | Remove-WriteConsole
	}
}
