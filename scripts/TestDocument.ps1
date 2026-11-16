param(
    [ValidateSet('Create','Snapshot','Save','Close','Open')][string]$Action='Snapshot',
    [string]$Path=(Join-Path $PSScriptRoot '..\artifacts\two-rectangles.pptx'),
    [string]$OutputPath
)
$ErrorActionPreference='Stop'
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
if ($Action -eq 'Save' -or $Action -eq 'Close') {
    $candidate=$ppt.ActivePresentation
    if ($candidate.Slides.Count -ne 1 -or $candidate.Slides.Item(1).Shapes.Count -ne 2 -or
        $candidate.Slides.Item(1).Shapes.Item(1).Name -ne 'GCS_Rectangle_A' -or
        $candidate.Slides.Item(1).Shapes.Item(2).Name -ne 'GCS_Rectangle_B') {
        throw 'The active presentation is not the two-rectangle test document.'
    }
}
switch($Action) {
    'Create' { $ppt.COMAddIns.Item('PptGcs.Connect').Object.CreateDemo($null) }
    'Save' { $dir=Split-Path -Parent ([IO.Path]::GetFullPath($Path)); New-Item -ItemType Directory -Force -Path $dir | Out-Null; $ppt.ActivePresentation.SaveAs([string][IO.Path]::GetFullPath($Path),24) }
    'Close' { $ppt.ActivePresentation.Close(); if ($ppt.Presentations.Count -eq 0) { $ppt.Quit() }; return }
    'Open' { $null=$ppt.Presentations.Open([string][IO.Path]::GetFullPath($Path),0,0,-1) }
}
$p=$ppt.ActivePresentation
$slide=$p.Slides.Item(1)
$shapes=@(foreach($s in $slide.Shapes) {[pscustomobject]@{Name=$s.Name;Id=$s.Id;Tag=$s.Tags.Item('PPTGCS_ID');Left=$s.Left;Top=$s.Top;Width=$s.Width;Height=$s.Height}})
$result=[pscustomobject]@{Utc=[DateTime]::UtcNow.ToString('O');ProcessId=@(Get-Process POWERPNT | Select-Object -ExpandProperty Id);PowerPointVersion=$ppt.Version;File=$p.FullName;SlideId=$slide.SlideID;Shapes=$shapes;ConstraintXml=@($p.CustomXMLParts.SelectByNamespace('urn:ppt-gcs:v1') | ForEach-Object {$_.XML})} | ConvertTo-Json -Depth 5
if ($OutputPath) { $result | Set-Content -LiteralPath $OutputPath -Encoding UTF8 }
$result
