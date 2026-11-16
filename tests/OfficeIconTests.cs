using System;
using System.Drawing;
using System.Runtime.InteropServices;
using PptGcs.AddIn;

internal static class OfficeIconTests
{
    private static int passed;
    [STAThread] private static int Main()
    {
        try
        {
            using(var app=new FakeIconApplication())
            using(var provider=new OfficeIconProvider(app,null))
            {
                var image=(Bitmap)provider.Get("SameX",24);
                Color pixel=image.GetPixel(12,12);
                Check(pixel.A==128 && pixel.R==255 && Math.Abs(pixel.G-128)<=1 && pixel.B==0,
                    "Office bitmap premultiplied alpha survives the native handle copy");
                Check(image.GetPixel(0,0).A==0,"transparent icon background remains transparent");
                Check(ReferenceEquals(image,provider.Get("SameX",24)) && app.CommandBars.Calls==1,
                    "repeated access uses one managed image and one Office call");
                Check(ReferenceEquals(image,provider.Get("SameX",24)) && !image.Size.IsEmpty,
                    "returned image remains usable independently of temporary picture objects");
                Check(provider.Get("not-a-command",24)==null && app.CommandBars.Calls==1,
                    "unmapped commands never query Office");
                bool rejected=false;
                try{provider.Get("SameX",15);}catch(ArgumentOutOfRangeException){rejected=true;}
                Check(rejected,"Office image size bounds are enforced");
            }
            int errors=0;
            using(var app=new FakeIconApplication {CommandBars={Fail=true}})
            using(var provider=new OfficeIconProvider(app,(id,error)=>errors++))
            {
                Check(provider.Get("SameY",24)==null && provider.Get("SameY",24)==null &&
                    app.CommandBars.Calls==1 && errors==1,"failed Office glyph is reported once and not retried while painting");
            }
            using(var app=new FakeIconApplication())
            {
                var provider=new OfficeIconProvider(app,null);
                provider.Get("SameX",24);provider.Dispose();provider.Dispose();
                bool rejected=false;
                try{provider.Get("SameX",24);}catch(ObjectDisposedException){rejected=true;}
                Check(rejected,"disposed provider rejects image access and can be disposed twice");
            }
            Console.WriteLine("PASS: "+passed+" Office icon tests");return 0;
        }
        catch(Exception exception){Console.Error.WriteLine(exception);return 1;}
    }
    private static void Check(bool condition,string name)
    {if(!condition)throw new Exception(name);passed++;}
}

// Public types let the production dynamic Office adapter use the same contract.
public sealed class FakeIconApplication : IDisposable
{
    public FakeIconCommandBars CommandBars {get;}=new FakeIconCommandBars();
    public void Dispose(){CommandBars.Dispose();}
}
public sealed class FakeIconCommandBars : IDisposable
{
    public int Calls {get;private set;}
    public bool Fail {get;set;}
    private IntPtr handle;
    public object GetImageMso(string id,int width,int height)
    {
        Calls++;
        if(Fail)throw new InvalidOperationException("Office cannot supply this image");
        if(handle==IntPtr.Zero)
        {
            var header=new BitmapHeader {Size=40,Width=24,Height=-24,Planes=1,BitCount=32};
            IntPtr pixels;
            handle=CreateDIBSection(IntPtr.Zero,ref header,0,out pixels,IntPtr.Zero,0);
            if(handle==IntPtr.Zero)throw new InvalidOperationException("Test DIB creation failed");
            var data=new byte[24*24*4];int middle=(12*24+12)*4;
            data[middle+1]=64;data[middle+2]=128;data[middle+3]=128;
            Marshal.Copy(data,0,pixels,data.Length);
        }
        return new FakeIconPicture {Handle=unchecked((int)handle.ToInt64())};
    }
    public void Dispose(){if(handle!=IntPtr.Zero){DeleteObject(handle);handle=IntPtr.Zero;}}
    [StructLayout(LayoutKind.Sequential)] private struct BitmapHeader
    {public uint Size;public int Width,Height;public ushort Planes,BitCount;public uint Compression,SizeImage;public int X,Y;public uint Used,Important;}
    [DllImport("gdi32.dll")]private static extern IntPtr CreateDIBSection(IntPtr dc,ref BitmapHeader info,uint usage,out IntPtr bits,IntPtr section,uint offset);
    [DllImport("gdi32.dll")]private static extern bool DeleteObject(IntPtr handle);
}
public sealed class FakeIconPicture
{public int Type {get{return 1;}}public int Handle {get;set;}}
