[CmdletBinding()]
param(
    [ValidateSet('Create','Snapshot','Save','Close','Open','SelectSlide')][string]$Action='Snapshot',
    [ValidateRange(1,6)][int]$Slide=1,
    [string]$Path,
    [string]$OutputPath
)
$ErrorActionPreference='Stop'
if (-not $Path) { $Path=Join-Path $PSScriptRoot '..\artifacts\layout-v2.pptx' }
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$cases=@('SameY','SameX','SameWidth','SameHeight','EqualHorizontalGap','EqualVerticalGap')
function Assert-TestPresentation($presentation) {
    if ($null -eq $presentation -or $presentation.Slides.Count -ne $cases.Count) {
        throw 'The active presentation is not the six-slide layout-v2 test document.'
    }
    for ($index=1; $index -le $cases.Count; $index++) {
        if ($presentation.Slides.Item($index).Tags.Item('PPTGCS_CASE') -ne $cases[$index-1]) {
            throw "Unexpected test case marker on slide $index. No document action was performed."
        }
    }
}
if ($Action -eq 'Create') {
    $ppt.COMAddIns.Item('PptGcs.Connect').Object.CreateDemo($null)
} elseif ($Action -eq 'Open') {
    $null=$ppt.Presentations.Open([string][IO.Path]::GetFullPath($Path),0,0,-1)
}
$presentation=$ppt.ActivePresentation
Assert-TestPresentation $presentation
switch ($Action) {
    'Save' {
        $absolutePath=[string][IO.Path]::GetFullPath($Path)
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $absolutePath) | Out-Null
        $presentation.SaveAs($absolutePath,24)
    }
    'Close' {
        if ($presentation.Saved -eq 0) { throw 'Save the test presentation before closing it.' }
        $presentation.Close()
        if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() }
        return
    }
    'SelectSlide' { $ppt.ActiveWindow.View.GotoSlide($Slide) }
}
$snapshot=$ppt.COMAddIns.Item('PptGcs.Connect').Object.GetSnapshot() | ConvertFrom-Json
$result=[pscustomobject][ordered]@{
    Utc=[DateTime]::UtcNow.ToString('O')
    ProcessId=@(Get-Process POWERPNT | Select-Object -ExpandProperty Id)
    PowerPointVersion=$ppt.Version
    presentation=$snapshot.presentation
    slides=$snapshot.slides
    ConstraintXml=@($presentation.CustomXMLParts.SelectByNamespace('urn:ppt-gcs:v2') | ForEach-Object {$_.XML})
} | ConvertTo-Json -Depth 25
if ($OutputPath) {
    $outputDirectory=Split-Path -Parent ([IO.Path]::GetFullPath($OutputPath))
    New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
    $result | Set-Content -LiteralPath $OutputPath -Encoding UTF8
}
$result
