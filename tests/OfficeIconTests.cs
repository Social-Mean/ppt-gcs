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
                Check(pixel.A==79 && pixel.R==255 && pixel.G==135 && pixel.B==45,
                    "real Office straight-alpha edge bytes survive the native handle copy");
                Check(image.GetPixel(0,0).A==0,"transparent icon background remains transparent");
                using(var light=Composite(image,Color.White))
                using(var dark=Composite(image,Color.FromArgb(32,32,32)))
                {
                    Color lightEdge=light.GetPixel(12,12),darkEdge=dark.GetPixel(12,12);
                    Check(Near(lightEdge,255,218,190) && Near(darkEdge,101,64,36),
                        "semi-transparent native edge composites once without bright halos on light and dark surfaces");
                    Check(Near(dark.GetPixel(0,0),32,32,32),
                        "RGB stored under zero alpha does not leak a white fringe into the background");
                }
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
            using(var app=new FakeIconApplication {CommandBars={TransparentOnly=true}})
            using(var provider=new OfficeIconProvider(app,null))
                Check(((Bitmap)provider.Get("SameX",24)).GetPixel(12,12).A==0,
                    "a transparent 32-bit source stays transparent rather than becoming an opaque rectangle");
            Console.WriteLine("PASS: "+passed+" Office icon tests");return 0;
        }
        catch(Exception exception){Console.Error.WriteLine(exception);return 1;}
    }
    private static void Check(bool condition,string name)
    {if(!condition)throw new Exception(name);passed++;}
    private static Bitmap Composite(Image image,Color background)
    {
        var result=new Bitmap(image.Width,image.Height);
        using(var graphics=Graphics.FromImage(result))
        {graphics.Clear(background);graphics.DrawImageUnscaled(image,0,0);}
        return result;
    }
    private static bool Near(Color color,int red,int green,int blue)
    {return Math.Abs(color.R-red)<=1 && Math.Abs(color.G-green)<=1 && Math.Abs(color.B-blue)<=1;}
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
    public bool TransparentOnly {get;set;}
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
            // Exact straight BGRA values observed in the installed Office
            // AspectRatio 16 px artwork, plus white RGB under zero alpha.
            var data=new byte[24*24*4];
            for(int i=0;i<data.Length;i+=4){data[i]=255;data[i+1]=255;data[i+2]=255;}
            if(!TransparentOnly)
            {
                int middle=(12*24+12)*4;
                data[middle]=45;data[middle+1]=135;data[middle+2]=255;data[middle+3]=79;
            }
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
