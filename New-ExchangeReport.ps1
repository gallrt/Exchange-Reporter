#Requires -RunAsAdministrator
#--------------------------------------------------------------------------------------
# Exchange Reporter 3.13
# for Exchange Server 2010/2013/2016/2019
# www.frankysweb.de
#
# Generating Exchange Reports
# by Frank Zoechling
#
#--------------------------------------------------------------------------------------

param(
	[Parameter(Mandatory = $false)][string]$Installpath = $PSScriptRoot,
	[Parameter(Mandatory = $false)][string]$ExchangeVersion,
	[Parameter(Mandatory = $false)][string]$ConfigFile = "settings.ini"
)

# Konsole Header
#--------------------------------------------------------------------------------------

$reporterversion = "3.13"

if ($ExchangeVersion) {
	$EMSVersion = $ExchangeVersion
}
$otitle = $host.ui.RawUI.WindowTitle
$host.ui.RawUI.WindowTitle = "Exchange Reporter $ReporterVersion - www.FrankysWeb.de"


Write-Host "
------------------------------------------------------------------------------------------"
Write-Host "
   _____         _                             ______                      _
  |  ___|       | |                            | ___ \                    | |
  | |____  _____| |__   __ _ _ __   __ _  ___  | |_/ /___ _ __   ___  _ __| |_ ___ _ __
  |  __\ \/ / __| '_ \ / _`` | '_ \ / _`` |/ _ \ |    // _ \ '_ \ / _ \| '__| __/ _ \ '__|
  | |___>  < (__| | | | (_| | | | | (_| |  __/ | |\ \  __/ |_) | (_) | |  | ||  __/ |
  \____/_/\_\___|_| |_|\__,_|_| |_|\__, |\___| \_| \_\___| .__/ \___/|_|  \__ \___|_|
                                    __/ |                | |
                                   |___/                 |_|
" -ForegroundColor cyan
Write-Host "
           for Exchange Server 2010 / 2013 / 2016 / 2019 / Office365

                                     www.FrankysWeb.de

           Version: $ReporterVersion

------------------------------------------------------------------------------------------
"
# Prüfen ob PowerShell 4.0 vorhanden
#--------------------------------------------------------------------------------------

Write-Host " Checking Powershell Version:" -NoNewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
$psversion = (Get-Host).version.major

if ($psversion -ge "4") {
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "OK (PowerShell $psversion)" -ForegroundColor green
} else {
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "Error" -ForegroundColor red
	exit 0
	Write-Host ""
}

# Laden der Funktionen aus "Include-Functions.ps1"
#--------------------------------------------------------------------------------------

Write-Host " Loading functions from Include-Functions.ps1:" -NoNewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
$functionfile = Test-Path "$installpath\Includes\Include-Functions.ps1"
if ($functionfile) {
	. "$installpath\Includes\Include-Functions.ps1"
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "Done" -ForegroundColor green
} else {
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "Error (not found)" -ForegroundColor red
	exit 0
	Write-Host ""
}

# settings.ini einlesen
#--------------------------------------------------------------------------------------

try {
	Write-Host " Loading settings from $ConfigFile`:" -NoNewline
	$origpos = $host.UI.RawUI.CursorPosition
	$origpos.X = 70
	$globalsettingsfile = "$installpath\$ConfigFile"
	$inifile = get-inicontent "$globalsettingsfile"
} catch {
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "Error" -ForegroundColor red
	exit 0
	Write-Host ""
}
$host.UI.RawUI.CursorPosition = $origpos
Write-Host "Done" -ForegroundColor green

# Settings verarbeiten
#--------------------------------------------------------------------------------------

# INI sections
$activemoduleshash = $inifile["Modules"]
$3rdPartyactivemoduleshash = $inifile["3rdPartyModules"]
$reportsettingshash = $inifile["Reportsettings"]
$reportsettings = convert-hashtoobject $reportsettingshash
$languagehash = $inifile["LanguageSettings"]
$excludehash = $inifile["ExcludeList"]
$languagesettings = convert-hashtoobject $languagehash
$excludelist = convert-hashtoobject $excludehash
$activemodules = convert-hashtoobject $activemoduleshash
$excludelist = $excludelist | Where-Object { $_.setting -notmatch "Comment" -and $_.setting -notmatch ";" }
$activemodules = $activemodules | Where-Object { $_.setting -notmatch "Comment" -and $_.setting -notmatch ";" } | Sort-Object setting
$3rdPartyactivemodules = convert-hashtoobject $3rdPartyactivemoduleshash
$3rdPartyactivemodules = $3rdPartyactivemodules | Where-Object { $_.setting -notmatch "Comment" -and $_.setting -notmatch ";" } | Sort-Object setting

# Einstellungen:
#--------------------------------------------------------------------------------------

$ReportInterval = ($reportsettings | Where-Object { $_.Setting -eq "Interval" }).Value
$CleanTMPFolder = ($reportsettings | Where-Object { $_.Setting -eq "CleanTMPFolder" }).Value
$Errorlog = ($reportsettings | Where-Object { $_.Setting -eq "WriteErrorLog" }).Value
$AddPDFFileToMail = ($reportsettings | Where-Object { $_.Setting -eq "AddPDFFileToMail" }).Value
$SMTPAuth = ($reportsettings | Where-Object { $_.Setting -eq "SMTPServerAuth" }).Value

if ($SMTPAuth -match "yes") {
	$SMTPServerUser = ($reportsettings | Where-Object { $_.Setting -eq "SMTPServerUser" }).Value
	$SMTPServerPass = ($reportsettings | Where-Object { $_.Setting -eq "SMTPServerPass" }).Value

	$secpasswd = ConvertTo-SecureString $SMTPServerPass -AsPlainText -Force
	$smtpcreds = New-Object System.Management.Automation.PSCredential ($SMTPServerUser, $secpasswd)
}

$Recipient = ($reportsettings | Where-Object { $_.Setting -eq "Recipient" }).Value
[array]$Recipient = $Recipient.split(",")
$Sender = ($reportsettings | Where-Object { $_.Setting -eq "Sender" }).Value
$Mailserver = ($reportsettings | Where-Object { $_.Setting -eq "Mailserver" }).Value
$Subject = ($reportsettings | Where-Object { $_.Setting -eq "Subject" }).Value
[int]$DisplayTop = ($reportsettings | Where-Object { $_.Setting -eq "DisplayTop" }).Value
$language = ($languagesettings | Where-Object { $_.Setting -eq "Language" }).Value

# Errorlog schreiben
#--------------------------------------------------------------------------------------

if ($errorlog -match "yes") {
	$logtime = Get-Date
	"-Start-- $logtime ----------------------------------------------------------------------------------" | Add-Content "$installpath\ErrorLog.txt"
}

# Sprache anzeigen
#--------------------------------------------------------------------------------------

try {
	Write-Host " Setting Report Language:" -NoNewline
	$origpos = $host.UI.RawUI.CursorPosition
	$origpos.X = 70
} catch {
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "Error" -ForegroundColor red
	if ($errorlog -match "yes") {
		$error[0] | Add-Content "$installpath\ErrorLog.txt"
	}
	exit 0
	Write-Host ""
}
$host.UI.RawUI.CursorPosition = $origpos
Write-Host "$language" -ForegroundColor green


# Lade Exchange Snapin
#--------------------------------------------------------------------------------------

Write-Host " Loading Exchange Management SnapIn:" -NoNewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
try {
	. "$installpath\Includes\Include-ExchangeSnapins.ps1"
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "Done" -ForegroundColor green
} catch {
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "Error" -ForegroundColor red
	if ($errorlog -match "yes") {
		$error[0] | Add-Content "$installpath\ErrorLog.txt"
	}
	exit 0
	Write-Host ""
}

# Exchange Version ermitteln
#--------------------------------------------------------------------------------------

Write-Host " Checking Exchange Management Shell:" -NoNewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
if (!$emsversion) {
	$emsversion = Get-ExchangeVersionByRegistry
}
if ($emsversion -match "2010") {
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "OK (Exchange 2010)" -ForegroundColor green
}
if ($emsversion -match "2013") {
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "OK (Exchange 2013)" -ForegroundColor green
}
if ($emsversion -match "2016") {
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "OK (Exchange 2016)" -ForegroundColor green
}
if ($emsversion -match "2019") {
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "OK (Exchange 2019)" -ForegroundColor green
}
if (!$emsversion) {
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "Error (EMS not found)" -ForegroundColor red
	exit 0
	Write-Host ""
}
if ($emsversion -notmatch "2010" -and $emsversion -notmatch "2013" -and $emsversion -notmatch "2016" -and $emsversion -notmatch "2019") {
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "Error (Wrong EMS Version)" -ForegroundColor red
	if ($errorlog -match "yes") {
		"Wrong EMS Version" | Add-Content "$installpath\ErrorLog.txt"
	}
	exit 0
	Write-Host ""
}

# Temporäres Verzeichnis erstellen
#--------------------------------------------------------------------------------------

Write-Host " Generating temp. Directory:" -NoNewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
if (Test-Path "$installpath\TEMP") { Remove-Item "$installpath\TEMP" -Force -Recurse }
$tmpdir = New-Item "$installpath\TEMP" -Type directory -ea 0
$tmpdir = $tmpdir.fullname
if ($tmpdir) {
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "Done" -ForegroundColor green
} else {
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "Error" -ForegroundColor red
	if ($errorlog -match "yes") {
		$error[0] | Add-Content "$installpath\ErrorLog.txt"
	}
	exit 0
	Write-Host ""
}

# Lade .NET Assembly
#--------------------------------------------------------------------------------------

Write-Host " Loading .NET Assemblies:" -NoNewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
try {
	[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "Done" -ForegroundColor green
} catch {
	if ($errorlog -match "yes") {
		$error[0] | Add-Content "$installpath\ErrorLog.txt"
	}
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "Error" -ForegroundColor red
	exit 0
	Write-Host ""
}

# Häufig genutzte Variablen
#--------------------------------------------------------------------------------------

Write-Host " Loading global Variables:" -NoNewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
try {
	$mbxservers = Get-MailboxServer -ea 0
	if ($emsversion -match "2010") {
		$casservers = Get-ClientAccessServer -ea 0
	}
	if ($emsversion -match "2013") {
		$casservers = Get-ClientAccessServer -ea 0
	}
	if ($emsversion -match "2016") {
		$casservers = Get-ClientAccessService -ea 0
	}
	if ($emsversion -match "2019") {
		$casservers = Get-ClientAccessService -ea 0
	}

	$ExServers = Get-ExchangeServer -ea 0 | Where-Object { $_.admindisplayversion.major -ge 14 }
	$ExDomains = $ExServers | Select-Object Domain -Unique
	$DomainControllers = @()
	foreach ($ExDomain in $ExDomains) {
		$DomainName = $ExDomain.domain
		$DomainControllers += Get-DomainController -DomainName $DomainName -ea 0
	}
	$OrgName = (Get-OrganizationConfig).Name
	#$emsversion = Get-ExchangeVersionByRegistry
	$files = @()
	$host.UI.RawUI.CursorPosition = $origpos
	$modulpath = "$installpath" + "\modules"
	$languagefilepath = "$installpath" + "\Language\" + "$language"
	$Start = (Get-Date -Hour 00 -Minute 00 -Second 00).AddDays(-$ReportInterval)
	$End = (Get-Date -Hour 00 -Minute 00 -Second 00)
	$Today = Get-Date | Convert-Date
	Write-Host "Done" -ForegroundColor green
	$EntireForrest = Set-ADServerSettings -ViewEntireForest $True
} catch {
	if ($errorlog -match "yes") {
		$error[0] | Add-Content "$installpath\ErrorLog.txt"
	}
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "Error" -ForegroundColor red
	exit 0
	Write-Host ""
}
Write-Host ""
Write-Host "------------------------------------------------------------------------------------------"
Write-Host ""


# MODULE
#--------------------------------------------------------------------------------------

# HTML Datei vorbereiten
$htmlheader = New-HTMLHeader ExchangeReporter
$htmlheader | Set-Content "$tmpdir\report.html"

foreach ($activemodule in $activemodules) {
	$module = $activemodule.Value
	Write-Host " Working on Module '$module':" -NoNewline
	$origpos = $host.UI.RawUI.CursorPosition
	$origpos.X = 70
	try {
		. "$LanguageFilepath\$module"
		. "$ModulPath\$module"
		$host.UI.RawUI.CursorPosition = $origpos
		Write-Host "Done" -ForegroundColor green
	} catch {
		if ($errorlog -match "yes") {
			$module | Add-Content "$installpath\ErrorLog.txt"
			$error[0] | Add-Content "$installpath\ErrorLog.txt"
		}
		$host.UI.RawUI.CursorPosition = $origpos
		Write-Host "Error" -ForegroundColor red
		Write-Host ""
	}
}

foreach ($3rdPartyactivemodule in $3rdPartyactivemodules) {
	$module = $3rdPartyactivemodule.Value
	Write-Host " Working on 3rd Party Module '$module':" -NoNewline
	$origpos = $host.UI.RawUI.CursorPosition
	$origpos.X = 70
	try {
		. "$modulpath\3rdParty\$module"
		$host.UI.RawUI.CursorPosition = $origpos
		Write-Host "Done" -ForegroundColor green
	} catch {
		if ($errorlog -match "yes") {
			$module | Add-Content "$installpath\ErrorLog.txt"
			$error[0] | Add-Content "$installpath\ErrorLog.txt"
		}
		$host.UI.RawUI.CursorPosition = $origpos
		Write-Host "Error" -ForegroundColor red
		Write-Host ""
	}
}

# Report vorbereiten
#--------------------------------------------------------------------------------------

Generate-ReportFooter | Add-Content "$tmpdir\report.html"
$mailbody = Get-Content "$tmpdir\report.html" | Out-String

foreach ($activemodule in $activemodules) {
	$module = $activemodule.Value
	$pngfile = $module.replace(".ps1", ".png")
	$files += Get-ChildItem "$Installpath\Images\$pngfile" -Recurse | Where-Object { -not $_.PSIsContainer } | ForEach-Object { $_.fullname }
}

foreach ($3rdPartyactivemodule in $3rdPartyactivemodules) {
	$module = $3rdPartyactivemodule.Value
	$pngfile = $module.replace(".ps1", ".png")
	$files += Get-ChildItem "$Installpath\Images\$pngfile" -Recurse | Where-Object { -not $_.PSIsContainer } | ForEach-Object { $_.fullname }
}

$files += Get-ChildItem "$Installpath\Images\reportheader.png" | Where-Object { -not $_.PSIsContainer } | ForEach-Object { $_.fullname }
$files += Get-ChildItem "$Installpath\TEMP\*.png" | Where-Object { -not $_.PSIsContainer } | ForEach-Object { $_.fullname }

# PDF File erzeugen
#--------------------------------------------------------------------------------------

if ($AddPDFFileToMail -match "yes") {
	Write-Host " Saving PDF Report:" -NoNewline
	$origpos = $host.UI.RawUI.CursorPosition
	$origpos.X = 70
	if (Test-Path "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe") {
		try {
			$pdfreport = $mailbody.replace("cid:", "")
			$pdfreport | Set-Content "$installpath\TEMP\PDFReport.htm"
			$pdfpath = "$installpath\TEMP"
			$pdffile = "$installpath\TEMP\Report.pdf"
			foreach ($file in $files) {
				Copy-Item $file -Destination $pdfpath -Force -ea 0
			}
			$pdf = &"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe" --quiet --enable-local-file-access "$installpath\TEMP\PDFReport.htm" "$installpath\TEMP\Report.pdf" | Out-Null
			$files += "$installpath\TEMP\Report.pdf"
			$host.UI.RawUI.CursorPosition = $origpos
			Write-Host "Done" -ForegroundColor green
		} catch {
			if ($errorlog -match "yes") {
				$error[0] | Add-Content "$installpath\ErrorLog.txt"
			}
			$host.UI.RawUI.CursorPosition = $origpos
			Write-Host "Error" -ForegroundColor red
			Write-Host ""
		}
	} else {
		$host.UI.RawUI.CursorPosition = $origpos
		Write-Host "Error (WKHTML not found)" -ForegroundColor red
		Write-Host ""
	}
}

# Report per Mail verschicken
#--------------------------------------------------------------------------------------

Write-Host ""
Write-Host " Sending Report:" -NoNewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
try {
	if ($SMTPAuth -match "yes") {
		Send-MailMessage -Encoding UTF8 -From "Exchange Reporter - www.FrankysWeb.de <$sender>" -To "$Recipient"  -Subject "$subject" -SmtpServer $mailserver -BodyAsHtml -Body $mailbody -Attachments $files -Credential $smtpcreds
		$host.UI.RawUI.CursorPosition = $origpos
		Write-Host "Done" -ForegroundColor green
	} else {
		Send-MailMessage -Encoding UTF8 -From "Exchange Reporter - www.FrankysWeb.de <$sender>" -To $Recipient  -Subject "$subject" -SmtpServer $mailserver -BodyAsHtml -Body $mailbody -Attachments $files
		$host.UI.RawUI.CursorPosition = $origpos
		Write-Host "Done" -ForegroundColor green
		Write-Host ""
	}
} catch {
	if ($errorlog -match "yes") {
		$error[0] | Add-Content "$installpath\ErrorLog.txt"
	}
	$host.UI.RawUI.CursorPosition = $origpos
	Write-Host "Error" -ForegroundColor red
	Write-Host ""
}

# Report per FTP Hochladen
#--------------------------------------------------------------------------------------

if ($FTPUpload -match "yes") {
	Write-Host ""
	Write-Host "------------------------------------------------------------------------------------------"
	Write-Host " Uploading files to FTP Server:" -NoNewline
	$origpos = $host.UI.RawUI.CursorPosition
	$origpos.X = 70
	try {
		$FTPServer = ($reportsettings | Where-Object { $_.Setting -eq "FTPServer" }).Value
		$FTPUser = ($reportsettings | Where-Object { $_.Setting -eq "FTPUser" }).Value
		$FTPPass = ($reportsettings | Where-Object { $_.Setting -eq "FTPPass" }).Value
		$FTPLocalFolder = ($reportsettings | Where-Object { $_.Setting -eq "FTPLocalFolder" }).Value

		$webclient = New-Object System.Net.WebClient
		$webclient.Credentials = New-Object System.Net.NetworkCredential($FTPUser, $FTPPass)

		foreach ($item in (Get-ChildItem $FTPLocalFolder)) {
			"Uploading $item..."
			$uri = New-Object System.Uri($FTPServer + $item.Name)
			$webclient.UploadFile($uri, $item.FullName)
		}

		Write-Host "Done" -ForegroundColor green
		Write-Host ""
	} catch {
		if ($errorlog -match "yes") {
			$error[0] | Add-Content "$installpath\ErrorLog.txt"
		}
		$host.UI.RawUI.CursorPosition = $origpos
		Write-Host "Error" -ForegroundColor red
		Write-Host ""
	}
}

# Aufräumen
#--------------------------------------------------------------------------------------

if ($CleanTMPFolder -match "yes") {
	Write-Host ""
	Write-Host "------------------------------------------------------------------------------------------"
	Write-Host " Cleaning up temp. Directory:" -NoNewline
	$origpos = $host.UI.RawUI.CursorPosition
	$origpos.X = 70
	try {
		$delTMPdir = Remove-Item $tmpdir -Recurse -Force
		$host.UI.RawUI.CursorPosition = $origpos
		Write-Host "Done" -ForegroundColor green
		Write-Host ""
	} catch {
		if ($errorlog -match "yes") {
			$error[0] | Add-Content "$installpath\ErrorLog.txt"
		}
		$host.UI.RawUI.CursorPosition = $origpos
		Write-Host "Error" -ForegroundColor red
		Write-Host ""
	}
}


# Errorlog schliessen
#--------------------------------------------------------------------------------------

if ($errorlog -match "yes") {
	$logtime = Get-Date
	"-End--- $logtime ----------------------------------------------------------------------------------" | Add-Content "$installpath\ErrorLog.txt"
}

$host.ui.RawUI.WindowTitle = $otitle