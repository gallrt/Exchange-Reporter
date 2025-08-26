$redirectreport = Generate-ReportHeader "redirectreport.png" "$l_redir_header "

$cells = @("$l_redir_mbx", "$l_redir_rulename", "$l_redir_type", "$l_redir_targetaddr", "$l_redir_active")
$redirectreport += Generate-HTMLTable "$l_redir_header2" $cells

$rules = Get-Mailbox -ResultSize Unlimited | ForEach-Object { Get-InboxRule -Mailbox $PSItem.Id } | Where-Object { $_.ForwardTo -or $_.RedirectTo }
foreach ($rule in $rules) {
	$mbxname = $rule.MailboxOwnerId.Name
	$rulename = $rule.Name
	$ruleactive = $rule.Enabled
	if ($rule.ForwardTo -and !$rule.RedirectTo) {
		$type = "$l_redir_forward"
		$target = $rule.ForwardTo.DisplayName
	}
	if ($rule.RedirectTo -and !$rule.ForwardTo) {
		$type = "$l_redir_redir"
		$target = $rule.RedirectTo.DisplayName
	}
	if ($rule.RedirectTo -and $rule.ForwardTo) {
		$type = "$l_redir_forandredir"
		$target = $rule.RedirectTo.DisplayName
		$target += $rule.ForwardTo.DisplayName
	}

	$cells = @("$mbxname", "$rulename", "$type", "$target", "$ruleactive")
	$redirectreport += New-HTMLTableLine $cells
}

$redirectreport += End-HTMLTable

$cells = @("$l_redir_mbx", "$l_redir_forwardto", "$l_redir_targetaddr")
$redirectreport += Generate-HTMLTable "$l_redir_header3" $cells

$rules = Get-Mailbox -ResultSize Unlimited | Where-Object { $_.ForwardingAddress -ne $NULL } | Sort-Object -Property Name
foreach ($rule in $rules) {
	$mbxname = $rule.Name
	if ($rule.DeliverToMailboxAndForward -match "False") {
		$type = "$l_redir_onlytarget"
	} else {
		$type = "$l_redir_mbxandtarget"
	}

	$canname = $rule.ForwardingAddress
	$dn = ConvertFrom-Canonical $canname
	try {
		$adobj = Get-ADObject $dn -Properties ProxyAddresses
		[string]$target = ($adobj.ProxyAddresses | Select-String SMTP -CaseSensitive:$true)
		$target = $target.replace("SMTP:", "")
		if (!$target) {
			$target = "$dn"
		}
	} catch {
		$target = "$dn"
	}
	$cells = @("$mbxname", "$type", "$target")
	$redirectreport += New-HTMLTableLine $cells
}
$redirectreport += End-HTMLTable

$redirectreport | Add-Content "$tmpdir\report.html"