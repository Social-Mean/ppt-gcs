[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
if (-not [Environment]::Is64BitProcess) { throw 'Run this script in 64-bit Windows PowerShell for 64-bit PowerPoint.' }
$projectRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$assemblyPath = Join-Path $projectRoot 'bin\PptGcs.AddIn.dll'
if (-not (Test-Path -LiteralPath $assemblyPath)) { throw 'Run scripts\Build.ps1 first.' }
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
New-ItemProperty -LiteralPath $officeKey -Name Description -Value 'Two-way same-top constraints for PowerPoint shapes' -PropertyType String -Force | Out-Null
New-ItemProperty -LiteralPath $officeKey -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null
Write-Output 'PPT GCS registered for the current Windows user. Restart PowerPoint to load it.'
