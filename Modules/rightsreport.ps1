$rightsreport = Generate-ReportHeader "rightsreport.png" "$l_perm_header"

$cells = @("$l_perm_mbx", "$l_perm_database", "$l_perm_user", "$l_perm_permission")
$rightsreport += Generate-HTMLTable "$l_perm_header2" $cells

$allmbx = Get-Mailbox -ResultSize Unlimited
foreach ($mailbox in $allmbx) {
	$mbxname = $mailbox.DisplayName
	$mbxdb = $mailbox.Database
	$rights = Get-MailboxPermission $mailbox | Where-Object { $_.IsInherited -match "False" -and $_.User -notmatch "Selbst" -and $_.User -notmatch "Self" -and $_.Deny -match "False" }
	if ($rights) {
		foreach ($right in $rights) {
			$username = $right.User.RawIdentity
			$accessright = "$l_perm_fuccaccess"

			$cells = @("$mbxname", "$mbxdb", "$username", "$accessright")
			$rightsreport += New-HTMLTableLine $cells
		}
	}
	$sendOB = $mailbox.GrantSendOnBehalfTo
	if ($sendOB) {
		foreach ($right in $sendOB) {
			$username = $right.Name
			$accessright = "$l_perm_sendonbehalf"

			$cells = @("$mbxname", "$mbxdb", "$username", "$accessright")
			$rightsreport += New-HTMLTableLine $cells
		}
	}
	$sendas = Get-ADPermission $mailbox.DistinguishedName | Where-Object { ($_.ExtendedRights -like "*Send-As*") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF") -and -not ($_.User -like "NT-AUTORITÄT\SELBST") }
	if ($sendas) {
		foreach ($right in $sendas) {
			$username = $right.User.RawIdentity
			$accessright = "$l_perm_sendas"

			$cells = @("$mbxname", "$mbxdb", "$username", "$accessright")
			$rightsreport += New-HTMLTableLine $cells
		}
	}
}

$rightsreport += End-HTMLTable

$rightsreport | Set-Content "$tmpdir\rightsreport.html"
$rightsreport | Add-Content "$tmpdir\report.html"
