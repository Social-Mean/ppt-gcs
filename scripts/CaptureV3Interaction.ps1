param([Parameter(Mandatory=$true)][string]$Name)
$ErrorActionPreference='Stop'
if($Name -notmatch '^[A-Za-z0-9-]+$'){throw 'Invalid capture name'}
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$directory=Join-Path $root 'artifacts\v3-interaction'
$ppt=[Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$p=$ppt.ActivePresentation
if($p.FullName -ne (Join-Path $directory 'interaction.pptx') -or $p.Tags.Item('PPTGCS_TEST') -ne 'V3'){throw 'Unexpected interactive fixture'}
$addin=$ppt.COMAddIns.Item('PptGcs.Connect').Object
function Tags($object) {
    @(for($i=1;$i -le $object.Tags.Count;$i++) {if($object.Tags.Name($i) -like 'PPTGCS_*'){[pscustomobject]@{Name=$object.Tags.Name($i);Value=$object.Tags.Value($i)}}}) | Sort-Object Name
}
$state=[pscustomobject]@{
    Utc=[DateTime]::UtcNow.ToString('O');ProcessId=(Get-Process POWERPNT).Id;Path=$p.FullName
    PresentationTags=@(Tags $p);ConstraintXml=$addin.GetConstraintXml();Snapshot=($addin.GetSnapshot() | ConvertFrom-Json)
    Pane=($addin.GetPaneState() | ConvertFrom-Json);Radial=($addin.GetRadialState() | ConvertFrom-Json)
    Slides=@(foreach($slide in $p.Slides){[pscustomobject]@{Id=$slide.SlideID;Name=$slide.Name;Shapes=@(foreach($shape in $slide.Shapes){
        [pscustomobject]@{Id=$shape.Id;Name=$shape.Name;Tags=@(Tags $shape);Left=$shape.Left;Top=$shape.Top;Width=$shape.Width;Height=$shape.Height}
    })}})
}
$state | ConvertTo-Json -Depth 25 | Set-Content -LiteralPath (Join-Path $directory ($Name+'.json')) -Encoding UTF8
$baseline=Get-Content -LiteralPath (Join-Path $directory 'before.json') -Raw | ConvertFrom-Json
$start=[DateTimeOffset]::Parse($baseline.Utc)
$stop=[DateTimeOffset]::Parse($state.Utc)
# Export only this fixture's capture interval, never the full session history.
$lines=@(foreach($line in [IO.File]::ReadLines((Join-Path $env:LOCALAPPDATA 'PptGcs\layout-v2.log'))) {
    $separator=$line.IndexOf(' ')
    if($separator -le 0){continue}
    $time=[DateTimeOffset]::Parse($line.Substring(0,$separator))
    if($time -ge $start -and $time -le $stop){$line}
})
[IO.File]::WriteAllLines((Join-Path $directory 'events.log'),$lines,[Text.UTF8Encoding]::new($false))
Write-Output ('Captured '+$Name)
