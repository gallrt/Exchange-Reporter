# Tage die eine Verteilerliste nicht benutzt wurde
#---------------------------------------------------
$unuseddays = 14
#---------------------------------------------------

$dgreport = Generate-ReportHeader "dgreport.png" "$l_dg_header"

$cells = @("$l_dg_name", "$l_dg_email", "$l_dg_member")
$dgreport += Generate-HTMLTable "$l_dg_t1header $unuseddays $l_dg_t1header2" $cells

$end = Get-Date
$dgstart = $end.AddDays(-$unuseddays)

if ($emsversion -match "2010") {
	$distributiongroups = Get-DistributionGroup -ResultSize Unlimited | Select-Object PrimarySmtpAddress | Sort-Object PrimarySmtpAddress
	$counts = Get-TransportServer | Get-MessageTrackingLog -EventId Expand -ResultSize Unlimited -start $dgstart -end $end | Sort-Object RelatedRecipientAddress | Group-Object RelatedRecipientAddress | Sort-Object Name | Select-Object @{label = "PrimarySmtpAddress"; expression = { $_.Name } }, Count
}
if ($emsversion -match "2013") {
	$distributiongroups = Get-DistributionGroup -ResultSize Unlimited | Select-Object PrimarySmtpAddress | Sort-Object PrimarySmtpAddress
	$counts = Get-Transportservice | Get-MessageTrackingLog -EventId Expand -ResultSize Unlimited -start $dgstart -end $end | Sort-Object RelatedRecipientAddress | Group-Object RelatedRecipientAddress | Sort-Object Name | Select-Object @{label = "PrimarySmtpAddress"; expression = { $_.Name } }, Count
}
if ($emsversion -match "2016" -or $emsversion -match "2019") {
	$distributiongroups = Get-DistributionGroup -ResultSize Unlimited | Select-Object PrimarySmtpAddress | Sort-Object PrimarySmtpAddress
	$counts = Get-Transportservice | Get-MessageTrackingLog -EventId Expand -ResultSize Unlimited -Start $dgstart -End $end | Sort-Object RelatedRecipientAddress | Group-Object RelatedRecipientAddress | Sort-Object Name | Select-Object @{label = "PrimarySmtpAddress"; expression = { $_.Name } }, Count
}

if ($distributiongroups -and $counts) {
	$unuseddls = Compare-Object $distributiongroups $counts -SyncWindow 1000 -Property PrimarySmtpAddress -PassThru | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object -Property PrimarySmtpAddress | Sort-Object
}

if ($unuseddls) {
	foreach ($unuseddl in $unuseddls) {
		[string]$smtpaddress = $unuseddl.PrimarySmtpAddress
		$dg = Get-DistributionGroup $smtpaddress -ResultSize Unlimited
		$dgname = $dg.DisplayName
		$members = Get-DistributionGroupMember -Identity $dg -ResultSize Unlimited | Select-Object Name | ForEach-Object { $_.Name }
		if ($members) {
			$hasmembers = "$l_dg_memberyes"
		} else {
			$hasmembers = "$l_dg_memberno"
		}
		$cells = @("$dgname", "$smtpaddress", "$hasmembers")
		$dgreport += New-HTMLTableLine $cells
	}
} else {
	$cells = @("$l_dg_nounuseddg")
}

$dgreport += End-HTMLTable

$dgreport | Set-Content "$tmpdir\dgreport.html"
$dgreport | Add-Content "$tmpdir\report.html"