$mailreport = Generate-ReportHeader "mailreport.png" "$l_mail_header"

$cells = @("$l_mail_sendcount", "$l_mail_reccount", "$l_mail_volsend", "$l_mail_volrec")
$mailreport += Generate-HTMLTable "$l_mail_header2 $ReportInterval $l_mail_days" $cells

$mailexclude = ($excludelist | Where-Object { $_.Setting -match "mailreport" }).Value
if ($mailexclude) {
	[array]$mailexclude = $mailexclude.split(",")
}

if ($emsversion -match "2016" -or $emsversion -match "2019") {
	$transportservers = Get-TransportService
	$SendMails = Get-TransportService | Get-MessageTrackingLog -Start $Start -End $End -EventId Send -ea 0 -ResultSize Unlimited | Where-Object { $_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.Source -match "SMTP" } | Select-Object Sender, Recipients, Timestamp, TotalBytes, ClientHostname
	$ReceivedMails = Get-TransportService | Get-MessageTrackingLog -Start $Start -End $End -EventId Receive -ea 0 -ResultSize Unlimited | Where-Object { $_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.Source -match "SMTP" } | Select-Object Sender, Recipients, Timestamp, TotalBytes, ServerHostname
}

if ($emsversion -match "2013") {
	$transportservers = Get-TransportService
	$SendMails = Get-TransportService | Get-MessageTrackingLog -Start $Start -End $End -EventId Send -ea 0 -ResultSize Unlimited | Where-Object { $_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.Source -match "SMTP" } | Select-Object Sender, Recipients, Timestamp, TotalBytes, ClientHostname
	$ReceivedMails = Get-TransportService | Get-MessageTrackingLog -Start $Start -End $End -EventId Receive -ea 0 -ResultSize Unlimited | Where-Object { $_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.Source -match "SMTP" } | Select-Object Sender, Recipients, Timestamp, TotalBytes, ServerHostname
}

if ($emsversion -match "2010") {
	$transportservers = Get-TransportServer
	$SendMails = Get-TransportServer | Get-MessageTrackingLog -Start $Start -End $End -EventId Send -ea 0 -ResultSize Unlimited | Where-Object { $_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.Source -match "SMTP" } | Select-Object Sender, Recipients, Timestamp, TotalBytes, ClientHostname
	$ReceivedMails = Get-TransportServer | Get-MessageTrackingLog -Start $Start -End $End -EventId Receive -ea 0 -ResultSize Unlimited | Where-Object { $_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.Source -match "SMTP" } | Select-Object Sender, Recipients, Timestamp, TotalBytes, ServerHostname
}

if ($mailexclude) {
	foreach ($entry in $mailexclude) {
		$SendMails = $SendMails | Where-Object { $_.Sender -notmatch $entry -and $_.Recipients -notmatch $entry }
	}
	foreach ($entry in $mailexclude) {
		$ReceivedMails = $ReceivedMails | Where-Object { $_.Sender -notmatch $entry -and $_.Recipients -notmatch $entry }
	}
}


#--------------------------------------------------------------------------------------
# Total

$totalsendmail = $sendmails | Measure-Object TotalBytes -Sum
$totalreceivedmail = $receivedmails | Measure-Object TotalBytes -Sum

$totalsendvol = $totalsendmail.Sum
$totalreceivedvol = $totalreceivedmail.Sum
$totalsendvol = $totalsendvol / 1024 / 1024
$totalreceivedvol = $totalreceivedvol / 1024 / 1024
$totalsendvol = [System.Math]::Round($totalsendvol , 2)
$totalreceivedvol = [System.Math]::Round($totalreceivedvol , 2)

$totalsendcount = $totalsendmail.Count
$totalreceivedcount = $totalreceivedmail.Count

$totalmail = @{$l_mail_send = $totalsendcount }
$totalmail += @{$l_mail_received = $totalreceivedcount }

New-CylinderChart 500 400 "$l_mail_overallcount" Mails "$l_mail_count" $totalmail "$tmpdir\totalmailcount.png"

$totalmail = @{$l_mail_send = $totalsendvol }
$totalmail += @{$l_mail_received = $totalreceivedvol }

New-CylinderChart 500 400 "$l_mail_overallvolume" Mails "$l_mail_size" $totalmail "$tmpdir\totalmailvol.png"

$cells = @("$totalsendcount", "$totalreceivedcount", "$totalsendvol", "$totalreceivedvol")
$mailreport += New-HTMLTableLine $cells
$mailreport += End-HTMLTable
$mailreport += Include-HTMLInlinePictures "$tmpdir\totalmail*.png"


#--------------------------------------------------------------------------------------
# Je Server

if ($transportservers.count -gt 1) {
	$cells = @("$l_mail_servername", "$l_mail_overallcount", "$l_mail_overallvolume", "$l_mail_sendcount", "$l_mail_reccount", "$l_mail_volsend", "$l_mail_volrec")
	$mailreport += Generate-HTMLTable "$l_mail_header2 $ReportInterval $l_mail_days $l_mail_perserver" $cells

	$perserverstats = @()
	foreach ($transportserver in $transportservers) {
		$tpsname = $transportserver.Name
		$tpssend = $sendmails | Where-Object { $_.ClientHostname -match "$tpsname" } | Measure-Object TotalBytes -Sum
		$tpsreceive = $ReceivedMails | Where-Object { $_.ServerHostname -match "$tpsname" } | Measure-Object TotalBytes -Sum
		$tpssendcount = $tpssend.Count
		$tpsreceivecount = $tpsreceive.Count

		$tpssendvol = $tpssend.Sum
		$tpssendvol = $tpssendvol / 1024 / 1024
		$tpssendvol = [System.Math]::Round($tpssendvol , 2)
		$tpsreceivevol = $tpsreceive.Sum
		$tpsreceivevol = $tpsreceivevol / 1024 / 1024
		$tpsreceivevol = [System.Math]::Round($tpsreceivevol , 2)


		$tpstotalvol = $tpsreceivevol + $tpssendvol
		$tpstotalcount = $tpsreceivecount + $tpssendcount

		$cells = @("$tpsname", "$tpstotalcount", "$tpstotalvol", "$tpssendcount", "$tpsreceivecount", "$tpssendvol", "$tpsreceivevol")
		$mailreport += New-HTMLTableLine $cells

		$perserverstats += New-Object PSObject -Property @{Name = "$tpsname"; TotalCount = $tpstotalcount; SendCount = $tpssendcount; ReceiveCount = $tpsreceivecount; ToltalVolume = $tpstotalvol; SendVolume = $tpssendvol; ReceiveVolume = $tpsreceivevol }
	}
	$mailreport += End-HTMLTable

	foreach ($tpserver in $perserverstats) {
		$tpsname = $tpserver.Name
		$tpstotalvol = $tpserver.ToltalVolume
		$tpstotalcount = $tpserver.TotalCount
		$tpssendvol = $tpserver.SendVolume
		$tpsreceivedvol = $tpserver.ReceiveVolume
		$tpssendcount = $tpserver.SendCount
		$tpsreceivedcount = $tpserver.ReceiveCount

		$tpsvoldata += [ordered]@{$tpsname = $tpstotalvol }
		$tpscountdata += [ordered]@{$tpsname = $tpstotalcount }

		$tpsrscountdata += [ordered]@{"$tpsname $l_mail_send" = $tpssendcount }
		#$tpsrscountdata += @{"$tpsname $l_mail_received"=$tpsreceivedcount}

		$tpsrsvoldata += [ordered]@{"$tpsname $l_mail_send" = $tpssendvol }
		#$tpsrsvoldata += @{"$tpsname $l_mail_received"=$tpsreceivedvol}
	}

	foreach ($tpserver in $perserverstats) {
		$tpsname = $tpserver.Name
		$tpsreceivedvol = $tpserver.ReceiveVolume
		$tpsreceivedcount = $tpserver.ReceiveCount

		#$tpsrscountdata += @{"$tpsname $l_mail_send"=$tpssendcount}
		$tpsrscountdata += [ordered]@{"$tpsname $l_mail_received" = $tpsreceivedcount }

		#$tpsrsvoldata += @{"$tpsname $l_mail_send"=$tpssendvol}
		$tpsrsvoldata += [ordered]@{"$tpsname $l_mail_received" = $tpsreceivedvol }
	}

	New-CylinderChart 500 400 "$l_mail_overallcount" Mails "$l_mail_size $l_mail_overall" $tpsvoldata "$tmpdir\pertpsvol.png"
	New-CylinderChart 500 400 "$l_mail_overallcount" Mails "$l_mail_count $l_mail_overall" $tpscountdata "$tmpdir\pertpscount.png"
	New-CylinderChart 500 400 "$l_mail_overallcount" Mails "$l_mail_size" $tpsrsvoldata "$tmpdir\pertpsvolrs.png"
	New-CylinderChart 500 400 "$l_mail_overallcount" Mails "$l_mail_coun" $tpsrscountdata "$tmpdir\pertpscountrs.png"

	$mailreport += Include-HTMLInlinePictures "$tmpdir\pertps*.png"
}
$total += New-Object PSObject -Property @{Name = "$name"; Volume = $volume }


#--------------------------------------------------------------------------------------
# Days

$cells = @("$l_mail_date", "$l_mail_sendcount", "$l_mail_reccount", "$l_mail_volsend", "$l_mail_volrec")
$mailreport += Generate-HTMLTable "$l_mail_overviewperday" $cells

$daycounter = 1
do {
	$dayendcounter = $daycounter - 1
	$daystart = (Get-Date -Hour 00 -Minute 00 -Second 00).AddDays(-$daycounter)
	$dayend = (Get-Date -Hour 00 -Minute 00 -Second 00).AddDays(-$dayendcounter)

	$DayReceivedMails = $ReceivedMails | Where-Object { $_.Timestamp -ge $daystart -and $_.Timestamp -le $dayend }
	$DaySendMails = $sendmails | Where-Object { $_.Timestamp -ge $daystart -and $_.Timestamp -le $dayend }

	$daytotalsendmail = $daysendmails | Measure-Object TotalBytes -Sum
	$daytotalreceivedmail = $dayreceivedmails | Measure-Object TotalBytes -Sum

	$daytotalsendvol = $daytotalsendmail.Sum
	$daytotalreceivedvol = $daytotalreceivedmail.Sum
	$daytotalsendvol = $daytotalsendvol / 1024 / 1024
	$daytotalreceivedvol = $daytotalreceivedvol / 1024 / 1024
	$daytotalsendvol = [System.Math]::Round($daytotalsendvol , 2)
	$daytotalreceivedvol = [System.Math]::Round($daytotalreceivedvol , 2)

	$daytotalsendcount = $daytotalsendmail.Count
	$daytotalreceivedcount = $daytotalreceivedmail.Count

	$day = $daystart | Get-Date -Format "dd.MM.yy"

	$daystotalmailvol += [ordered]@{$day = $daytotalreceivedvol }
	$daystotalmailcount += [ordered]@{$day = $daytotalreceivedcount }

	$cells = @("$day", "$daytotalsendcount", "$daytotalreceivedcount", "$daytotalsendvol", "$daytotalreceivedvol")
	$mailreport += New-HTMLTableLine $cells

	$daycounter++
} while ($daycounter -le $reportinterval)

New-CylinderChart 500 400 "$l_mail_daycount" Mails "$l_mail_count" $daystotalmailcount "$tmpdir\dailymailcount.png"
New-CylinderChart 500 400 "$l_mail_daysize" Mails "$l_mail_size" $daystotalmailvol "$tmpdir\dailymailvol.png"

$mailreport += End-HTMLTable
$mailreport += Include-HTMLInlinePictures "$tmpdir\dailymail*.png"


#--------------------------------------------------------------------------------------
# Sender (by mail count)

$SendMailsSender = $sendmails.Sender
$topsenders = $SendMailsSender | Group-Object -NoElement |
		Sort-Object Count -Descending | Select-Object -First $DisplayTop

$cells = @("$l_mail_sender", "$l_mail_count")
$mailreport += Generate-HTMLTable "Top $DisplayTop $l_mail_sender ($l_mail_count)" $cells
foreach ($topsender in $topsenders) {
	$tsname = $topsender.Name
	$tscount = $topsender.Count

	$cells = @("$tsname", "$tscount")
	$mailreport += New-HTMLTableLine $cells
}
$mailreport += End-HTMLTable


#--------------------------------------------------------------------------------------
# Recipients  (by mail count)

$ReceivedMailsRecipients = $ReceivedMails.Recipients 
$toprecipients = $ReceivedMailsRecipients | Group-Object -NoElement |
		Sort-Object Count -Descending | Select-Object -First $DisplayTop

$cells = @("$l_mail_recipient", "$l_mail_count")
$mailreport += Generate-HTMLTable "Top $DisplayTop $l_mail_recipient ($l_mail_count)" $cells
foreach ($toprecipient in $toprecipients) {
	$trname = $toprecipient.Name
	$trcount = $toprecipient.Count

	$cells = @("$trname", "$trcount")
	$mailreport += New-HTMLTableLine $cells
}
$mailreport += End-HTMLTable


#--------------------------------------------------------------------------------------
# Sender (by volume)

$cells = @("$l_mail_sender", "$l_mail_sizemb")
$mailreport += Generate-HTMLTable "Top $DisplayTop $l_mail_sender ($l_mail_size)" $cells

$sendstatgroup = $SendMails | Group-Object Sender
$total = @()
foreach ($group in $sendstatgroup) {
	$name = ($group.Group | Select-Object -First 1).Sender
	$volume = ($group.Group | Measure-Object TotalBytes -Sum).Sum
	$total += New-Object PSObject -Property @{Name = "$name"; Volume = $volume }
}
$toptensendersvol = $total | Sort-Object Volume -Descending | Select-Object -First $DisplayTop

foreach ($topsender in $toptensendersvol) {
	$tsname = $topsender.Name
	$tsvolume = $topsender.Volume
	$tsvolume = $tsvolume / 1024 / 1024
	$tsvolume = [System.Math]::Round($tsvolume , 2)
	$cells = @("$tsname", "$tsvolume")
	$mailreport += New-HTMLTableLine $cells
}
$mailreport += End-HTMLTable


#--------------------------------------------------------------------------------------
# Recipients (by volume)

$cells = @("$l_mail_recipient", "$l_mail_sizemb")
$mailreport += Generate-HTMLTable "Top $DisplayTop $l_mail_recipient ($l_mail_size)" $cells

$receivedstat = @()
$total = @()
foreach ($receivedMail in $ReceivedMails) {
	foreach ($MailRecipient in $receivedMail.Recipients) {
		$receivedstat += New-Object PSObject -Property @{
			Recipient = $MailRecipient;
			TotalBytes = $receivedMail.TotalBytes;
		}
	}
}
$receivedstatgroup = $receivedstat | Group-Object Recipient

foreach ($group in $receivedstatgroup) {
	$name = ($group.Group | Select-Object -First 1).Recipient
	$volume = ($group.Group | Measure-Object TotalBytes -Sum).Sum
	$total += New-Object PSObject -Property @{Name = "$name"; Volume = $volume }
}
$toptenrecipientsvol = $total | Sort-Object Volume -Descending | Select-Object -First $DisplayTop

foreach ($toprecipient in $toptenrecipientsvol) {
	$trname = $toprecipient.Name
	$trvolume = $toprecipient.Volume
	$trvolume = $trvolume / 1024 / 1024
	$trvolume = [System.Math]::Round($trvolume , 2)
	$cells = @("$trname", "$trvolume")
	$mailreport += New-HTMLTableLine $cells
}
$mailreport += End-HTMLTable


#--------------------------------------------------------------------------------------
# Durchschnitt

try {
	$usercount = (Get-Mailbox -ResultSize Unlimited | Select-Object Alias).Count

	$dsend = $totalsendcount / $usercount
	$dsend = [System.Math]::Round($dsend , 2)
	$dreceived = $totalreceivedcount / $usercount
	$dreceived = [System.Math]::Round($dreceived , 2)
	$dsendvol = $totalsendvol / $usercount
	$dsendvol = [System.Math]::Round($dsendvol , 2)
	$dreceivedvol = $totalreceivedvol / $usercount
	$dreceivedvol = [System.Math]::Round($dreceivedvol  , 2)
	$dmailsizesend = $totalsendvol / $totalsendcount
	$dmailsizesend = [System.Math]::Round($dmailsizesend , 2)
	$dmailsizereceived = $totalreceivedvol / $totalreceivedcount
	$dmailsizereceived = [System.Math]::Round($dmailsizereceived , 2)

	$cells = @("$l_mail_average", "$l_mail_value")
	$mailreport += Generate-HTMLTable "$l_mail_averagevalue" $cells

	$cells = @("$l_mail_avmbxsendcount", "$dsend")
	$mailreport += New-HTMLTableLine $cells

	$cells = @("$l_mail_avmbxreccount", "$dreceived")
	$mailreport += New-HTMLTableLine $cells

	$cells = @("$l_mail_avmbxsendsize", "$dsendvol MB")
	$mailreport += New-HTMLTableLine $cells

	$cells = @("$l_mail_avmbxrecsize", "$dreceivedvol MB")
	$mailreport += New-HTMLTableLine $cells

	$cells = @("$l_mail_avmailsendsize", "$dmailsizesend MB")
	$mailreport += New-HTMLTableLine $cells

	$cells = @("$l_mail_avmailrecsize", "$dmailsizereceived MB")
	$mailreport += New-HTMLTableLine $cells

	$mailreport += End-HTMLTable
} catch {}

$mailreport | Set-Content "$tmpdir\mailreport.html"
$mailreport | Add-Content "$tmpdir\report.html"