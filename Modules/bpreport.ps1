$bpreport = Generate-ReportHeader "bpreport.png" "$l_bp_header "

# Übersicht Fehler
$cells = @("$l_bp_srvname", "$l_bp_svcname", "$l_bp_errorcount")
$bpreport += Generate-HTMLTable "$l_bp_t1header" $cells

$start = (Get-Date).AddDays(-$reportinterval)


foreach ($exserver in $exservers) {
	$eventsrv = $exserver.Name
	$bperrors = Get-WinEvent -ComputerName $eventsrv -FilterHashtable @{Logname = "application"; StartTime = [datetime]$start; level = 2 } -ea 0 | Where-Object { $_.ProviderName -match "exchange" -and $_.ID -match "1500*" } | Select-Object Message, Id, TimeCreated, ProviderName

	if ($bperrors) {
		$bperrorgroups = $bperrors | Group-Object ProviderName
		foreach ($bperrorgroup in $bperrorgroups) {
			$providername = $bperrorgroup.Name
			$errorcount = $bperrorgroup.Count

			$cells = @("$eventsrv", "$providername", "$errorcount")
			$bpreport += New-HTMLTableLine $cells
		}

		$bperrormessages = $bperrors | Select-Object Message -Unique
		foreach ($bperrormessage in $bperrormessages) {
			$message = $bperrormessage.Message
			$cells = @("$eventsrv", "$message")
			$bpdetailreport += New-HTMLTableLine $cells
		}
	}
}

$bpreport += End-HTMLTable

# Übersicht Warnungen
$cells = @("$l_bp_srvname", "$l_bp_svcname", "$l_bp_warncount")
$bpreport += Generate-HTMLTable "$l_bp_t2header" $cells

foreach ($exserver in $exservers) {
	$eventsrv = $exserver.Name
	$bpwarnings = Get-WinEvent -ComputerName $eventsrv -FilterHashtable @{Logname = "application"; StartTime = [datetime]$start; level = 3 } -ea 0 | Where-Object { $_.ProviderName -match "exchange" -and $_.ID -match "1500*" } | Select-Object Message, Id, TimeCreated, ProviderName

	if ($bpwarnings) {
		$bpwarninggroups = $bpwarnings | Group-Object ProviderName
		foreach ($bpwarninggroup in $bpwarninggroups) {
			$providername = $bpwarninggroup.Name
			$warningcount = $bpwarninggroup.Count

			$cells = @("$eventsrv", "$providername", "$warningcount")
			$bpreport += New-HTMLTableLine $cells
		}

		$bpwarningmessages = $bpwarnings | Select-Object Message -Unique
		foreach ($bpwarningmessage in $bpwarningmessages) {
			$message = $bpwarningmessage.Message
			$cells = @("$eventsrv", "$message")
			$bpdetailreport += New-HTMLTableLine $cells
		}
	}
}

$bpreport += End-HTMLTable

# Details
$cells = @("$l_bp_srvname", "$l_bp_discription")
$bpreport += Generate-HTMLTable "$l_bp_t3header" $cells

$bpreport += $bpdetailreport

$bpreport += End-HTMLTable

$bpreport | Set-Content "$tmpdir\serverinfo.html"
$bpreport | Add-Content "$tmpdir\report.html"