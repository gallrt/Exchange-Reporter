$badBytes = [byte[]]@(0xC3, 0x80)
$utf8Str = [System.Text.Encoding]::UTF8.GetString($badBytes)
$bytes = [System.Text.Encoding]::ASCII.GetBytes('Write-Output "') + [byte[]]@(0xC3, 0x80) + [byte[]]@(0x22)
$path = Join-Path ([System.IO.Path]::GetTempPath()) 'encodingtest.ps1'

try {
    [System.IO.File]::WriteAllBytes($path, $bytes)

    switch (& $path) {
        $utf8Str {
            return 'UTF-8'
            break
        }
        default {
            return 'Windows-1252'
            break
        }
    }
} finally {
    Remove-Item $path
}