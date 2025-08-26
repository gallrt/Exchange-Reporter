$admodule = Get-Module -ListAvailable | Where-Object { $_.name -match "ActiveDirectory" }
if ($admodule) {
	Import-Module ActiveDirectory
} else {
	if ($errorlog -match "yes") {
		"Fehler: PowerShell ActiveDirectory Modul nicht gefunden" | Add-Content "$installpath\ErrorLog.txt"
	}
	exit 0
}

$adreport = Generate-ReportHeader "addsinfo.png" "$l_adds_header"

$cells = @("$l_adds_fname", "$l_adds_sversion", "$l_adds_sname", "$l_adds_ffl", "$l_adds_sites")
$adreport += Generate-HTMLTable "$l_adds_overview" $cells

$schemaversion = (Get-ADObject (Get-ADRootDSE).schemaNamingContext -Property objectVersion).objectVersion

if ($schemaversion -eq "13") {
	$schemaname = "Windows Server 2000"
}
if ($schemaversion -eq "30") {
	$schemaname = "Windows Server 2003"
}
if ($schemaversion -eq "31") {
	$schemaname = "Windows Server 2003 R2"
}
if ($schemaversion -eq "44") {
	$schemaname = "Windows Server 2008"
}
if ($schemaversion -eq "47") {
	$schemaname = "Windows Server 2008 R2"
}
if ($schemaversion -eq "52") {
	$schemaname = "Windows Server 2012 Beta"
}
if ($schemaversion -eq "56") {
	$schemaname = "Windows Server 2012"
}
if ($schemaversion -eq "69") {
	$schemaname = "Windows Server 2012 R2"
}
if ($schemaversion -eq "87") {
	$schemaname = "Windows Server 2016"
}
if ($schemaversion -ge "88") {
	$schemaname = "Windows Server 2019/2022"
}

$adforest = Get-ADForest
$forestmode = $adforest.ForestMode
[string]$forestsites = $adforest.Sites
$forestsites = $forestsites.Replace(" ", ", ")
$forestname = $adforest.RootDomain

$cells = @("$forestname", "$schemaversion", "$schemaname", "$forestmode", "$forestsites")
$adreport += New-HTMLTableLine $cells

$adreport += End-HTMLTable

$cells = @("$l_adds_dname", "$l_adds_nbtname", "$l_adds_tdomain", "$l_adds_dfl", "$l_adds_dc")
$adreport += Generate-HTMLTable "$l_adds_adoverview" $cells

$addomains = $adforest.Domains
foreach ($addomain in $addomains) {
	$adds = Get-ADDomain $addomain
	$domainname = $addomain
	$netbiosname = $adds.NetBIOSName
	$parentdomain = $adds.ParentDomain
	if (!$parentdomain) {
		$parentdomain = "$l_adds_nothing"
	}
	$domainmode = $adds.DomainMode
	[string]$adcontrollers = $adds.ReplicaDirectoryServers
	$adcontrollers = $adcontrollers.Replace(" ", ", ")

	$cells = @("$domainname", "$netbiosname", "$parentdomain", "$domainmode", "$adcontrollers")
	$adreport += New-HTMLTableLine $cells
}

$adreport += End-HTMLTable

$cells = @("$l_adds_name", "$l_adds_contributer", "$l_adds_model", "$l_adds_os", "$l_adds_ram", "$l_adds_uptime")
$adreport += Generate-HTMLTable "$l_adds_dcoverview" $cells

foreach ($domaincontroller in $domaincontrollers) {
	try {
		$computername = $domaincontroller.Name
		$computerSystem = Get-WmiObject Win32_ComputerSystem -ComputerName $computername
		$computerOS = Get-WmiObject Win32_OperatingSystem -ComputerName $computername

		$hardware = $computerSystem.Manufacturer
		$model = $computerSystem.Model
		$os = $computerOS.Caption + ", SP: " + $computerOS.ServicePackMajorVersion
		$os = $os.replace("Microsoft Windows ", "")
		$ram = $computerSystem.TotalPhysicalMemory / 1gb
		$ram = [System.Math]::Round($ram, 2)
		$lastboot = $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
		$lastboot = Get-Date $lastboot -UFormat "%d.%m.%Y %R"

		$cells = @("$computername", "$hardware", "$model", "$os", "$ram", "$lastboot")
		$adreport += New-HTMLTableLine $cells
	} catch {
		$cells = @("WMI Error")
		$adreport += New-HTMLTableLine $cells
	}
}

$adreport += End-HTMLTable


foreach ($domaincontroller in $domaincontrollers) {
	try {
		$eventsrv = $domaincontroller.Name
		$cells = @("$l_adds_source", "$l_adds_timestamp", "$l_adds_id", "$l_adds_count", "$l_adds_message")
		$adreport += Generate-HTMLTable "$eventsrv - $l_adds_replerror" $cells

		$eventgroups = Get-WinEvent -ComputerName $eventsrv -FilterHashtable @{Logname = "*replication*"; StartTime = [datetime]$start; level = 2, 3 } -ea 0 | Select-Object message, id, timecreated | Group-Object id

		if ($eventgroups) {
			foreach ($eventgroup in $eventgroups) {
				$eventOne = $eventgroup.Group | Select-Object -First 1
				$eventcount = $eventgroup.Count
				$eventsource = $eventOne.ProviderName
				$eventid = $eventOne.Id
				$eventtime = $eventOne.TimeCreated
				$eventtime = $eventtime | Get-Date -Format "dd.MM.yy hh:mm:ss"
				$eventmessage = $eventOne.Message
				$eventmeslength = $eventmessage.Length
				if ($eventmeslength -gt 200) {
					$eventcontent = $eventmessage.Substring(0, 200)
					$eventcontent = $eventcontent + "..."
				} else {
					$eventcontent = $eventmessage
				}

				$cells = @("$eventsource", "$eventtime", "$eventid", "$eventcount", "$eventcontent")
				$adreport += New-HTMLTableLine $cells
			}
		} else {
			$cells = @("$l_adds_noerror")
			$adreport += New-HTMLTableLine $cells
		}
	} catch {
		$cells = @("WMI Error")
		$adreport += New-HTMLTableLine $cells
	}
	$adreport += End-HTMLTable
}

$adreport | Set-Content "$tmpdir\adreport.html"
$adreport | Add-Content "$tmpdir\report.html"