# Dot-source this module from Windows PowerShell 5.1 on an STA thread.
# Office may reject an automation call while its UI is busy. The standard OLE
# message filter retries only rejected COM calls, for at most ten seconds.
if (-not ('PptGcs.Testing.OfficeMessageFilterScope' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace PptGcs.Testing {
 [ComImport,Guid("00000016-0000-0000-C000-000000000046"),InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
 public interface IOfficeMessageFilter {
  [PreserveSig] int HandleInComingCall(int type,IntPtr caller,int ticks,IntPtr info);
  [PreserveSig] int RetryRejectedCall(IntPtr callee,int ticks,int reject);
  [PreserveSig] int MessagePending(IntPtr callee,int ticks,int type);
 }
 public sealed class OfficeMessageFilterScope : IOfficeMessageFilter,IDisposable {
  private IOfficeMessageFilter previous;
  private bool disposed;
  public OfficeMessageFilterScope() { Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(this,out previous)); }
  public int HandleInComingCall(int type,IntPtr caller,int ticks,IntPtr info) { return 0; }
  public int RetryRejectedCall(IntPtr callee,int ticks,int reject) { return ticks>=0 && ticks<10000 && (reject==1 || reject==2) ? 100 : -1; }
  public int MessagePending(IntPtr callee,int ticks,int type) { return 2; }
  public void Dispose() {
   if(disposed)return;
   IOfficeMessageFilter replaced;
   Marshal.ThrowExceptionForHR(CoRegisterMessageFilter(previous,out replaced));
   disposed=true;previous=null;
  }
  [DllImport("ole32.dll")] private static extern int CoRegisterMessageFilter(IOfficeMessageFilter value,out IOfficeMessageFilter old);
 }
}
'@
}

function Start-OfficeComSession {
    [CmdletBinding()]
    param()
    if ($PSVersionTable.PSEdition -eq 'Core' -or [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        throw 'Run this Office integration test with Windows PowerShell: powershell.exe -STA -File <script.ps1>.'
    }
    $filter = New-Object PptGcs.Testing.OfficeMessageFilterScope
    try {
        $createdApplication = $false
        try { $application = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
        catch [Runtime.InteropServices.COMException] {
            if ($_.Exception.HResult -ne -2147221021) { throw }
            $application = New-Object -ComObject PowerPoint.Application
            $createdApplication = $true
        }
        $originalWindow = $null
        if ($application.Windows.Count -gt 0) { $originalWindow = $application.ActiveWindow }
        $originalDocuments = @(for ($index = 1; $index -le $application.Presentations.Count; $index++) {
            $item = $application.Presentations.Item($index)
            [pscustomobject]@{ Document = $item; Name = [string]$item.Name; Path = [string]$item.FullName; Saved = [int]$item.Saved }
        })
        return [pscustomobject]@{
            Application = $application; Filter = $filter; OriginalWindow = $originalWindow
            OriginalDocuments = $originalDocuments; CreatedApplication = $createdApplication
        }
    }
    catch { $filter.Dispose(); throw }
}

function Stop-OfficeComSession {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Session)
    # Fixture ownership and window restoration belong to the calling test.
    # Never quit an Office application or close its documents here.
    $Session.Filter.Dispose()
}

function Test-OfficeComIdentity {
    [CmdletBinding()]
    param($First, $Second)
    if ($null -eq $First -or $null -eq $Second) { return $false }
    $left = [Runtime.InteropServices.Marshal]::GetIUnknownForObject($First)
    try {
        $right = [Runtime.InteropServices.Marshal]::GetIUnknownForObject($Second)
        try { return $left -eq $right }
        finally { [void][Runtime.InteropServices.Marshal]::Release($right) }
    }
    finally { [void][Runtime.InteropServices.Marshal]::Release($left) }
}
