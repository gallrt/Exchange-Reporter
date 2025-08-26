$warninglevel = 200
#-----------------------------------------------------------
$mbxreport = Generate-ReportHeader "mbxreport.png" "$l_mbx_header"

$cells = @("$l_mbx_name", "$l_mbx_size", "$l_mbx_database")
$mbxreport += Generate-HTMLTable "$l_mbx_topmbx ($DisplayTop)" $cells

$mbxexclude = ($excludelist | Where-Object { $_.Setting -match "mbxreport" }).Value
if ($mbxexclude) {
	[array]$mbxexclude = $mbxexclude.split(",")
	$mailboxes = Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics -ea 0 -wa 0
	foreach ($entry in $mbxexclude) {
		$mailboxes = $mailboxes | Where-Object { $_.displayname -notmatch $entry -or $_.alias -notmatch $entry }
	}
	$mailboxes = $mailboxes | Sort-Object TotalItemSize -Descending | Select-Object -First $DisplayTop
} else {
	$mailboxes = Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics -ea 0 -wa 0 | Sort-Object TotalItemSize -Descending | Select-Object -First $DisplayTop
}

foreach ($mailbox in $mailboxes) {
	$mbxname = $mailbox.DisplayName
	$mbxsize = $mailbox.TotalItemSize
	$db = $mailbox.Database

	$cells = @("$mbxname", "$mbxsize", "$db")
	$mbxreport += New-HTMLTableLine $cells
}
$mbxreport += End-HTMLTable

$databases = Get-MailboxDatabase

$mbxlimits = @()
foreach ($database in $databases) {
	$mbxdatabase = $database.Name
	$dblimit = $database.ProhibitSendQuota
	if ($dblimit -match "Unlimited") {
		$dblimitstate = "Inactive"
		$dblimitvalue = "Unlimited"
	} else {
		$dblimitstate = "Active"
		$dblimitvalue = $dblimit.value.toMB()
	}
	$mailboxesindb = Get-Mailbox -Database $database -ResultSize Unlimited | Sort-Object
	foreach ($mailbox in $mailboxesindb) {
		$mbxname = $mailbox.Name
		$mbxalias = $mailbox.Alias
		$mbxsize = (Get-MailboxStatistics $mailbox -wa 0).TotalItemSize
		if (!$mbxsize) {
			$mbxsize = 0
		} else {
			$mbxsize = $mbxsize.Value.toMB()
		}
		$mbxlimit = $mailbox.ProhibitSendQuota
		$mbxdefault = $mailbox.UseDatabaseQuotaDefaults
		if ($mbxlimit -match "Unlimited") {
			$mbxlimitstate = "Inactive"
			$mbxlimitvalue = "Unlimited"
		} else {
			$mbxlimitstate = "Active"
			$mbxlimitvalue = $mbxlimit.Value.toMB()
		}
		[array]$mbxlimits += New-Object PSObject -Property @{Mailbox = "$mbxname"; DBlimit = "$dblimitstate"; DBLimitValue = "$dblimitvalue"; MBXLimit = "$mbxlimitstate"; MBXLimitValue = "$mbxlimitvalue"; MBXSize = "$mbxsize"; MBXAlias = "$mbxalias"; MBXUseDBDefault = "$mbxdefault"; Database = $mbxdatabase }
	}
}

$reportlimits = @()
foreach ($mailbox in $mbxlimits) {
	$mbxname = $mailbox.mailbox
	$mbxalias = $mailbox.mbxalias
	[double]$mbxsize = $mailbox.mbxsize
	$mbxdatabase = $mailbox.database

	$mbxlimit = $mailbox.mbxlimit
	$dblimit = $mailbox.dblimit
	$mbxusedbdefault = $mailbox.MBXUseDBDefault

	#es gilt das Limit der Datenbank
	if ($mbxusedbdefault -eq "True" -and $dblimit -eq "Active") {
		[double]$limitsize = $mailbox.dblimitvalue
		$warningactive = $mbxsize -ge ($limitsize - $warninglevel)
		$limittype = "Database"
		[array]$reportlimits += New-Object PSObject -Property @{Mailbox = "$mbxname"; MBXAlias = "$mbxalias"; LimitType = "$Limittype"; LimitSize = "$limitsize"; MailboxSize = "$mbxsize"; WarningActive = "$warningactive"; Database = $mbxdatabase }
	}
	#es gilt das Limit des Postfachs
	if ($mbxusedbdefault -eq "False" -and $mbxlimit -eq "Active") {
		[double]$limitsize = $mailbox.MBXLimitValue
		$warningactive = $mbxsize -ge ($limitsize - $warninglevel)
		$limittype = "Mailbox"
		[array]$reportlimits += New-Object PSObject -Property @{Mailbox = "$mbxname"; MBXAlias = "$mbxalias"; LimitType = "$Limittype"; LimitSize = "$limitsize"; MailboxSize = "$mbxsize"; WarningActive = "$warningactive"; Database = $mbxdatabase }
	}
}
$reportlimits = $reportlimits | Where-Object { $_.WarningActive -match "True" }

$cells = @("$l_mbx_name", "$l_mbx_size", "$l_mbx_limit", "$l_mbx_database", "$l_mbx_limittype")
$mbxreport += Generate-HTMLTable "$l_mbx_mbxlimit" $cells
if ($reportlimits) {
	foreach ($mbx in $reportlimits) {
		$mbxname = $mbx.mailbox
		$mbxsize = $mbx.mailboxsize
		$mbxlimit = $mbx.limitsize
		$mbxdb = $mbx.database
		$limittype = $mbx.limittype
		$cells = @("$mbxname", "$mbxsize", "$mbxlimit", "$mbxdb", "$limittype")
		$mbxreport += New-HTMLTableLine $cells
	}
} else {
	$cells = @("$l_mbx_nolimit")
	$mbxreport += New-HTMLTableLine $cells
}

$mbxreport += End-HTMLTable

#Getrennte Mailboxen

$cells = @("$l_mbx_name", "$l_mbx_database", "$l_mbx_size", "$l_mbx_disconnected", "$l_mbx_id")
$mbxreport += Generate-HTMLTable "$l_mbx_dismbx" $cells

$dismbxs = Get-MailboxDatabase | Get-MailboxStatistics -wa 0 -ea 0 | Where-Object { $_.DisconnectDate -ne $null } | Select-Object DisplayName, Identity, DisconnectDate, Database, TotalItemSize
foreach ($dismbx in $dismbxs) {
	$dismbxname = $dismbx.DisplayName
	$disdb = $dismbx.Database
	$dissize = $dismbx.TotalItemSize
	[string]$disdate = $dismbx.DisconnectDate | Get-Date -UFormat %d.%m.%Y
	$disid = $dismbx.Identity

	$cells = @("$dismbxname", "$disdb ", "$dissize", "$disdate", "$disid")
	$mbxreport += New-HTMLTableLine $cells
}
$mbxreport += End-HTMLTable


#Inaktive Mailboxen

$cells = @("$l_mbx_name", "$l_mbx_database", "$l_mbx_size", "$l_mbx_lastlogin", "$l_mbx_lastloginfrom")
$mbxreport += Generate-HTMLTable "$l_mbx_maybeinactive" $cells

$logonstats = Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics -wa 0 -ea 0 | Select-Object DisplayName, Database, TotalItemSize, LastLoggedOnUserAccount, LastLogonTime | Where-Object { $_.LastLogonTime -lt ((Get-Date).AddDays(-120)) } | Sort-Object LastLogonTime
foreach ($entry in $logonstats) {
	$ianame = $entry.DisplayName
	$iadb = $entry.Database
	$iasize = $entry.TotalItemSize
	$iall = $entry.LastLogonTime
	$iauser = $entry.LastLoggedOnUserAccount
	if (!$iall) {
		$iastate = "$l_mbx_userdeactivated"
	}
	if (!$iall -and $iaobj -notmatch "Disabled") {
		$iastate = "$l_mbx_unknown"
	}
	if ($iall) {
		$iastate = $iall
	}

	$cells = @("$ianame", "$iadb", "$iasize", "$iastate", "$iauser")
	$mbxreport += New-HTMLTableLine $cells
}
$mbxreport += End-HTMLTable

$mbxreport | Set-Content "$tmpdir\mbxreport.html"
$mbxreport | Add-Content "$tmpdir\report.html"