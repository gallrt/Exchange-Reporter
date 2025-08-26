$pfreport = Generate-ReportHeader "pfreport.png" "$l_pf_header"

if ($emsversion -match "2010") {
	$cells = @("$l_pf_db", "$l_pf_server", "$l_pf_size", "$l_pf_lastbackup")
	$pfreport += Generate-HTMLTable "$l_pf_header2" $cells

	$pfdbs = Get-PublicFolderDatabase -Status
	foreach ($pfdb in $pfdbs) {
		$pfdbname = $pfdb.Name
		$pfdbserver = $pfdb.Server
		$pfdbsize = $pfdb.DatabaseSize
		$pflastbackup = $pfdb.LastFullBackup
		if ($pflastbackup) {
			$pflastbackup = Get-Date $pflastbackup -UFormat "%d.%m.%Y %R"
		} else {
			$pflastbackup = "Nie"
		}

		$cells = @("$pfdbname", "$pfdbserver", "$pfdbsize", "$pflastbackup")
		$pfreport += New-HTMLTableLine $cells
	}
	$pfreport += End-HTMLTable

	$cells = @("$l_pf_name", "$l_pf_db", "$l_pf_size", "$l_pf_elementcount")
	$pfreport += Generate-HTMLTable "$l_pf_header3" $cells

	$pfs = Get-PublicFolderStatistics -ResultSize Unlimited -ea 0 | Sort-Object TotalItemSize -Descending | Select-Object -First 200
	foreach ($pf in $pfs) {
		$pfname = $pf.AdminDisplayName
		$pfdb = $pf.DatabaseName
		$pfsize = $pf.TotalItemSize
		$pfitemcount = $pf.ItemCount

		$cells = @("$pfname", "$pfdb", "$pfsize", "$pfitemcount")
		$pfreport += New-HTMLTableLine $cells
	}
	$pfreport += End-HTMLTable
}

if ($emsversion -match "2013" -or $emsversion -match "2016" -or $emsversion -match "2019") {
	$cells = @("$l_pf_mbx", "$l_pf_server", "$l_pf_size", "$l_pf_db")
	$pfreport += Generate-HTMLTable "$l_pf_header4" $cells

	$pfmbxs = Get-Mailbox -PublicFolder
	foreach ($pfmbx in $pfmbxs) {
		$pfmbxname = $pfmbx.Name
		$pfmbxserver = $pfmbx.ServerName
		$pfmbxsize = (Get-MailboxStatistics $pfmbx).TotalItemSize.value
		$pfmbxdatabase = $pfmbx.Database.Name
		$cells = @("$pfmbxname", "$pfmbxserver", "$pfmbxsize", "$pfmbxdatabase")
		$pfreport += New-HTMLTableLine $cells
	}
	$pfreport += End-HTMLTable

	$cells = @("$l_pf_name", "$l_pf_db", "$l_pf_size", "$l_pf_elementcount")
	$pfreport += Generate-HTMLTable "$l_pf_header3" $cells

	$pfs = Get-PublicFolderStatistics -ResultSize Unlimited -ea 0 | Sort-Object TotalItemSize -Descending | Select-Object -First 200
	foreach ($pf in $pfs) {
		$pfname = $pf.Name
		$pfid = $pf.EntryId
		$pfdb = (Get-PublicFolder $pfid).ContentMailboxName
		$pfsize = $pf.TotalItemSize
		$pfitemcount = $pf.ItemCount

		$cells = @("$pfname", "$pfdb", "$pfsize", "$pfitemcount")
		$pfreport += New-HTMLTableLine $cells
	}
	$pfreport += End-HTMLTable
}

$pfreport | Set-Content "$tmpdir\pfreport.html"
$pfreport | Add-Content "$tmpdir\report.html"
