# Exchange-Reporter
 Exchange Reporter creates email reports with statistics for Exchange Server

## Character Encoding
This project uses different Encodings for different files. Some are
`Windows-1252`, some are `UTF-8`, some are `utf-8 with BOM`, and some are still
different.

Powershell expects all files to be in his default encoding.

> PowerShell's default encoding varies depending on version:
> - In PowerShell 6+, the default encoding is UTF-8 without BOM on all platforms.
> - In Windows PowerShell, the default encoding is usually Windows-1252,
>   which is an extension of latin-1 (also known as ISO 8859-1).
> 
> Source: [learn.microsoft.com](https://learn.microsoft.com/en-us/powershell/scripting/dev-cross-plat/vscode/understanding-file-encoding)

You can find your encoding using `.\Test-Encoding.ps1`

You can find your Powershell Version using `$PSVersionTable.PSVersion`

### Windows Powershell 5.1
If you run this code on Windows with Powershell 5.1 (eg on
Windows Server 2016, 2019, 2022, 2025 or Windows 10 and 11),
you should change the encoding either to `UTF-8 with BOM`
or to your `default encoding` (depends on your locale).

However, for the file `settings.ini` it is necessary to be in
UTF-8 encoding, for non-ASCII letters (e.g. Umlaut in German)
to work correctly.

Encoding MUST be changed for following files to work:
- `Modules\rightsreport.ps1`
- `Includes\Include-Functions.ps1`

```powershell
$FilePath = ".\Modules\rightsreport.ps1"

$Content = Get-Content $FilePath -Raw -Encoding UTF8
Set-Content -Path $FilePath -Encoding Default -Value "$Content"
```

### Powershell 7
If you would run this code with PowerShell 6 or higher,
you dont't need to change anything for the files in `UTF-8`
or in `UTF-8 with BOM`.

All files in folder `Language` use the `Windows-1252` encoding.

Encoding MUST be changed for following files to work:
- All files in folder `Language\DE`
- `Includes\Include-Functions.ps1`

```powershell
$MyPath = ".\Includes\Include-Functions.ps1"
$FileInfo = New-Item -Force -Path $MyPath -Value (Get-Content -Raw -Path $MyPath -Encoding Windows1252)
```


## Usage
See Blog for usage: [Exchange Reporter](https://www.frankysweb.de/exchange-reporter-2013/)

English manual: Exchange Reporter Handbuch (EN).docx

German manual: Exchange Reporter Handbuch (DE).docx

## Website
 [FrankysWeb](https://www.frankysweb.de/)
