#--------------------------------------------------------------------------------------
# Abstand zum Limit in MB, ab dem ein Postfach in die Liste der
# "Postfächer nahe am Sende-Limit" aufgenommen wird.
$warninglevel = 200

$mbxreport = Generate-ReportHeader "mbxreport.png" "$l_mbx_header"


#--------------------------------------------------------------------------------------
# größte Postfächer

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


#--------------------------------------------------------------------------------------
# Postfächer nahe am Sende-Limit

$databases = Get-MailboxDatabase
$reportlimits = @()
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

		[bool] $limited = $false
		if ($mbxdefault -eq "True" -and $dblimitstate -eq "Active") {
			# es gilt das Limit der Datenbank
			$limited = $true
			[double]$limitsize = $dblimitvalue
			$limittype = "Database"
		}
		if ($mbxdefault -eq "False" -and $mbxlimitstate -eq "Active") {
			# es gilt das Limit des Postfachs
			$limited = $true
			[double]$limitsize = $MBXLimitValue
			$limittype = "Mailbox"
		}
		if ($limited -and $mbxsize -ge ($limitsize - $warninglevel)) {
			# Postfach hat Limit (DB oder PF) und ist nahe an ihrem Limit
			[array]$reportlimits += New-Object PSObject -Property @{
				Mailbox     = "$mbxname";
				MBXAlias    = "$mbxalias";
				LimitType   = "$Limittype";
				LimitSize   = "$limitsize";
				MailboxSize = "$mbxsize";
				Database    = $mbxdatabase
			}
		}
	}
}

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


#--------------------------------------------------------------------------------------
# Getrennte Mailboxen

$cells = @("$l_mbx_name", "$l_mbx_database", "$l_mbx_size", "$l_mbx_disconnected", "$l_mbx_id")
$mbxreport += Generate-HTMLTable "$l_mbx_dismbx" $cells

$dismbxs = Get-MailboxDatabase | Get-MailboxStatistics -wa 0 -ea 0 | Where-Object { $_.DisconnectDate -ne $null } |
			Select-Object DisplayName, Identity, DisconnectDate, Database, TotalItemSize
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


#--------------------------------------------------------------------------------------
# Inaktive Mailboxen

$cells = @("$l_mbx_name", "$l_mbx_database", "$l_mbx_size", "$l_mbx_lastlogin", "$l_mbx_lastloginfrom")
$mbxreport += Generate-HTMLTable "$l_mbx_maybeinactive" $cells

$logonstats = Get-Mailbox -ResultSize Unlimited | Get-MailboxStatistics -wa 0 -ea 0 |
			Select-Object DisplayName, Database, TotalItemSize, LastLoggedOnUserAccount, LastLogonTime |
			Where-Object { $_.LastLogonTime -lt ((Get-Date).AddDays(-120)) } |
			Sort-Object LastLogonTime
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