[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$ownedKeys = @(
    'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\PptGcs.Connect',
    'HKCU:\Software\Classes\PptGcs.Connect',
    'HKCU:\Software\Classes\CLSID\{1B728F4B-5AF0-41FD-8589-86E44A7A5501}',
    'HKCU:\Software\Classes\PptGcs.ConstraintPaneHost',
    'HKCU:\Software\Classes\CLSID\{7FA35AF7-2CF8-4B2A-8F5D-2A4BCC1E21A0}'
)
foreach ($key in $ownedKeys) {
    if (-not (Test-Path -LiteralPath $key)) { continue }
    $resolvedPath = (Get-Item -LiteralPath $key).Name
    if ($resolvedPath -notmatch '^HKEY_CURRENT_USER\\Software\\(Microsoft\\Office\\PowerPoint\\Addins\\PptGcs\.Connect|Classes\\PptGcs\.(Connect|ConstraintPaneHost)|Classes\\CLSID\\\{(1B728F4B-5AF0-41FD-8589-86E44A7A5501|7FA35AF7-2CF8-4B2A-8F5D-2A4BCC1E21A0)\})$') {
        throw "Refusing to delete unexpected registry path: $resolvedPath"
    }
    Remove-Item -LiteralPath $key -Recurse -Force
}
Write-Output 'PPT GCS registration removed for the current Windows user. Restart PowerPoint to unload it.'
