$ErrorActionPreference = 'Stop'
$ppt = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
[pscustomobject]@{Version=$ppt.Version;Visible=$ppt.Visible;WindowState=$ppt.WindowState;Presentations=$ppt.Presentations.Count;Left=$ppt.Left;Top=$ppt.Top;Width=$ppt.Width;Height=$ppt.Height;HWND=$ppt.HWND} | ConvertTo-Json
$ppt.COMAddIns.Update()
$addin=$ppt.COMAddIns.Item('PptGcs.Connect')
$addin.Connect=$true
$addin.Object.GetStatus()
$addin.Object.GetDiagnostics()
