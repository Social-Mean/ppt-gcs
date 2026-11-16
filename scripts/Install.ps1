[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitProcess) { throw 'Run this script in 64-bit Windows PowerShell for 64-bit PowerPoint.' }
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$assemblyPath = Join-Path $projectRoot 'bin\PptGcs.AddIn.dll'
if (-not (Test-Path -LiteralPath $assemblyPath)) { throw 'Run scripts\Build.ps1 first.' }
$paneHostPath = Join-Path $projectRoot 'bin\PptGcs.PaneHost.dll'
if (-not (Test-Path -LiteralPath $paneHostPath)) { throw 'Run scripts\Build.ps1 first to build the native task pane host.' }
$assemblyIdentity = [Reflection.AssemblyName]::GetAssemblyName($assemblyPath)
if ($assemblyIdentity.Name -ne 'PptGcs.AddIn') { throw 'Unexpected add-in assembly identity.' }
$classId = '{1B728F4B-5AF0-41FD-8589-86E44A7A5501}'
$progId = 'PptGcs.Connect'
$className = 'PptGcs.AddIn.Connect'
$classKey = "HKCU:\Software\Classes\CLSID\$classId"
$serverKey = "$classKey\InprocServer32"
$versionKey = "$serverKey\$($assemblyIdentity.Version)"
$progIdKey = "HKCU:\Software\Classes\$progId"
$officeKey = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$progId"
$codeBase = ([Uri]::new($assemblyPath)).AbsoluteUri

foreach ($key in @($classKey, $serverKey, $versionKey, "$classKey\ProgId", $progIdKey, "$progIdKey\CLSID", $officeKey)) {
    New-Item -Path $key -Force | Out-Null
}
Set-Item -LiteralPath $classKey -Value $className
Set-Item -LiteralPath $serverKey -Value 'mscoree.dll'
New-ItemProperty -LiteralPath $serverKey -Name ThreadingModel -Value Both -PropertyType String -Force | Out-Null
foreach ($key in @($serverKey, $versionKey)) {
    New-ItemProperty -LiteralPath $key -Name Class -Value $className -PropertyType String -Force | Out-Null
    New-ItemProperty -LiteralPath $key -Name Assembly -Value $assemblyIdentity.FullName -PropertyType String -Force | Out-Null
    New-ItemProperty -LiteralPath $key -Name RuntimeVersion -Value 'v4.0.30319' -PropertyType String -Force | Out-Null
    New-ItemProperty -LiteralPath $key -Name CodeBase -Value $codeBase -PropertyType String -Force | Out-Null
}
Set-Item -LiteralPath "$classKey\ProgId" -Value $progId
Set-Item -LiteralPath $progIdKey -Value $className
Set-Item -LiteralPath "$progIdKey\CLSID" -Value $classId
New-ItemProperty -LiteralPath $officeKey -Name FriendlyName -Value 'PPT GCS' -PropertyType String -Force | Out-Null
New-ItemProperty -LiteralPath $officeKey -Name Description -Value 'Interactive layout constraints for PowerPoint shapes' -PropertyType String -Force | Out-Null
New-ItemProperty -LiteralPath $officeKey -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null

# A native ATL control hosts the managed view. The WinForms UserControl itself
# is created by the add-in, not registered or activated through COM.
$paneClassId = '{7FA35AF7-2CF8-4B2A-8F5D-2A4BCC1E21A0}'
$paneProgId = 'PptGcs.ConstraintPaneHost'
$paneClassKey = "HKCU:\Software\Classes\CLSID\$paneClassId"
$paneProgIdKey = "HKCU:\Software\Classes\$paneProgId"
foreach ($key in @($paneClassKey, "$paneClassKey\InprocServer32", "$paneClassKey\ProgId", "$paneClassKey\Control", "$paneClassKey\Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}", $paneProgIdKey, "$paneProgIdKey\CLSID")) {
    New-Item -Path $key -Force | Out-Null
}
Set-Item -LiteralPath $paneClassKey -Value 'PPT GCS native task pane host'
Set-Item -LiteralPath "$paneClassKey\InprocServer32" -Value $paneHostPath
New-ItemProperty -LiteralPath "$paneClassKey\InprocServer32" -Name ThreadingModel -Value Apartment -PropertyType String -Force | Out-Null
Set-Item -LiteralPath "$paneClassKey\ProgId" -Value $paneProgId
Set-Item -LiteralPath $paneProgIdKey -Value 'PPT GCS native task pane host'
Set-Item -LiteralPath "$paneProgIdKey\CLSID" -Value $paneClassId
Write-Output 'PPT GCS registered for the current Windows user. Restart PowerPoint to load it.'
