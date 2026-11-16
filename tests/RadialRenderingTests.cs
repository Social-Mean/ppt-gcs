using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PptGcs.AddIn;
using PptGcs.Core;

internal static class RadialRenderingTests
{
    private const BindingFlags Private = BindingFlags.Instance | BindingFlags.NonPublic;
    private static int passed;

    [STAThread] private static int Main(string[] arguments)
    {
        try
        {
            foreach (float scale in new[] {1f,1.5f,2f})
            using (var fixture = new Fixture(scale))
            using (var image = fixture.Render())
            {
                Check(image.GetPixel(0,0).A == 0 && image.GetPixel(image.Width-1,image.Height-1).A == 0,
                    "outside the wheel remains genuinely transparent at scale " + scale);
                int translucent = 0, invalidPremultiplied = 0, magenta = 0;
                var bits = image.LockBits(new Rectangle(Point.Empty,image.Size),ImageLockMode.ReadOnly,PixelFormat.Format32bppPArgb);
                try
                {
                    var row = new byte[image.Width * 4];
                    for (int y=0;y<image.Height;y++)
                    {
                        Marshal.Copy(IntPtr.Add(bits.Scan0,y*bits.Stride),row,0,row.Length);
                        for(int x=0;x<image.Width;x++)
                        {
                            int offset=x*4, alpha=row[offset+3];
                            if(alpha>0 && alpha<255) translucent++;
                            if(row[offset]>alpha || row[offset+1]>alpha || row[offset+2]>alpha)invalidPremultiplied++;
                            if(alpha>0 && row[offset+2]>row[offset+1]+25 && row[offset]>row[offset+1]+25)magenta++;
                        }
                    }
                }
                finally { image.UnlockBits(bits); }
                Check(translucent>100,"curved edges retain fractional alpha instead of binary color-key clipping at scale "+scale);
                Check(invalidPremultiplied==0,"the complete layered frame uses valid premultiplied BGRA at scale "+scale);
                Check(magenta==0,"no magenta-key edge contamination at scale "+scale);
                int escapedLabels=0;
                for(int y=(int)(100*scale);y<(int)(391*scale);y++)for(int x=(int)(150*scale);x<(int)(450*scale);x++)
                {
                    double distance=Math.Sqrt(Math.Pow(x-300*scale,2)+Math.Pow(y-250*scale,2));
                    if(distance>138*scale && distance<150*scale && image.GetPixel(x,y).A!=0)escapedLabels++;
                }
                Check(escapedLabels==0,"long catalog labels stay inside the wheel perimeter at scale "+scale);
                using(var light=Composite(image,Color.White))
                using(var dark=Composite(image,Color.FromArgb(30,34,40)))
                {
                    Check(light.GetPixel(0,0).ToArgb()==Color.White.ToArgb() && dark.GetPixel(0,0).ToArgb()==Color.FromArgb(30,34,40).ToArgb(),
                        "the same alpha frame composes onto light and dark backgrounds at scale "+scale);
                    Check(image.GetPixel((int)(300*scale),(int)(250*scale)).A==255,
                        "text rendering preserves the opaque wheel face instead of punching out GDI alpha at scale "+scale);
                    if(arguments.Length>0)
                    {
                        Directory.CreateDirectory(arguments[0]);
                        string prefix=Path.Combine(arguments[0],"radial-"+(int)(scale*100));
                        image.Save(prefix+"-transparent.png",ImageFormat.Png);
                        light.Save(prefix+"-light.png",ImageFormat.Png);
                        dark.Save(prefix+"-dark.png",ImageFormat.Png);
                    }
                }
            }

            using(var fixture=new Fixture(1))
            {
                Form hud=fixture.Hud;
                IntPtr foreground=GetForegroundWindow(),capture=GetCapture();
                IntPtr handle=hud.Handle;
                fixture.Present();
                long flags=GetWindowLongPtr(handle,-20).ToInt64();
                Check((flags & 0x00080000)!=0 && hud.TransparencyKey.IsEmpty,"native HUD uses a per-pixel layered window with no transparency key");
                Check((flags & 0x08000020)==0x08000020 && SendMessage(handle,0x0084,IntPtr.Zero,IntPtr.Zero)==new IntPtr(-1),
                    "HUD remains nonactivating and passes pointer hit testing through to the canvas");
                Check(!hud.Visible && GetForegroundWindow()==foreground && GetCapture()==capture,
                    "native UpdateLayeredWindow succeeds on a hidden test window without desktop input or focus/capture changes");
                object frame=fixture.HudField("frame"),native=fixture.HudField("layeredSurface");
                int handles=GetGuiResources(Process.GetCurrentProcess().Handle,0);
                for(int i=0;i<12;i++)fixture.Present();
                Check(ReferenceEquals(frame,fixture.HudField("frame")) && ReferenceEquals(native,fixture.HudField("layeredSurface")) &&
                    GetGuiResources(Process.GetCurrentProcess().Handle,0)==handles,
                    "hover redraw reuses the managed frame and native DIB/DC without GDI-handle growth");
            }
            Console.WriteLine("PASS: "+passed+" radial alpha rendering tests (no PowerPoint or physical input)");
            return 0;
        }
        catch(Exception error){Console.Error.WriteLine(error);return 1;}
    }

    private static Bitmap Composite(Image source,Color color)
    {
        var image=new Bitmap(source.Width,source.Height,PixelFormat.Format32bppArgb);
        using(var graphics=Graphics.FromImage(image)){graphics.Clear(color);graphics.DrawImageUnscaled(source,0,0);}
        return image;
    }
    private static void Check(bool condition,string label)
    {if(!condition)throw new InvalidOperationException(label);passed++;Console.WriteLine("PASS "+label);}

    private sealed class Fixture:IDisposable
    {
        private readonly RadialMenuController controller;
        private readonly List<Image> icons=new List<Image>();
        public Form Hud {get;private set;}
        public Fixture(float scale)
        {
            controller=new RadialMenuController(()=>new ConstraintPaneModel(),kind=>new List<ShapeGeometry>(),kind=>{},more=>{},()=>null,()=>{},(kind,data)=>{},Icon);
            Set("dpiScale",scale);Set("outerRadius",136*scale);Set("innerRadius",38*scale);
            Set("center",new PointF(300*scale,250*scale));Set("selectedCount",2);Set("currentSlot",2);
            var slots=(Array)typeof(RadialMenuController).GetField("slots",Private).GetValue(controller);
            for(int index=0;index<slots.Length;index++)slots.GetValue(index).GetType().GetField("Enabled").SetValue(slots.GetValue(index),index!=5);
            var preview=(List<RectangleF>)typeof(RadialMenuController).GetField("previewBounds",Private).GetValue(controller);
            preview.Add(new RectangleF(20*scale,30*scale,80*scale,50*scale));
            Type type=typeof(RadialMenuController).GetNestedType("RadialHud",BindingFlags.NonPublic);
            Hud=(Form)Activator.CreateInstance(type,new object[]{controller});
            Hud.ClientSize=new Size((int)(680*scale),(int)(540*scale));
            Hud.Location=new Point(30,30);
        }
        private Image Icon(string command,int size)
        {
            // A simple test pattern exercises native-size image composition;
            // it is not an Office glyph or a PowerPoint screen capture.
            var image=new Bitmap(size,size,PixelFormat.Format32bppArgb);
            using(var graphics=Graphics.FromImage(image))
            {graphics.Clear(Color.Transparent);graphics.SmoothingMode=SmoothingMode.AntiAlias;using(var pen=new Pen(Color.FromArgb(35,117,198),2)){graphics.DrawRectangle(pen,3,3,size-7,size-7);graphics.DrawLine(pen,5,size/2,size-5,size/2);}}
            icons.Add(image);return image;
        }
        private void Set(string field,object value){typeof(RadialMenuController).GetField(field,Private).SetValue(controller,value);}
        public object HudField(string field){return Hud.GetType().GetField(field,Private).GetValue(Hud);}
        public Bitmap Render()
        {
            var image=new Bitmap(Hud.ClientSize.Width,Hud.ClientSize.Height,PixelFormat.Format32bppPArgb);
            using(var graphics=Graphics.FromImage(image))Hud.GetType().GetMethod("DrawFrame",Private).Invoke(Hud,new object[]{graphics});
            return image;
        }
        public void Present(){Hud.GetType().GetMethod("UpdateSurface").Invoke(Hud,null);}
        public void Dispose(){Hud.Dispose();controller.Dispose();foreach(var image in icons)image.Dispose();}
    }

    private static IntPtr GetWindowLongPtr(IntPtr window,int index)
    {return IntPtr.Size==8?GetWindowLongPtr64(window,index):new IntPtr(GetWindowLong32(window,index));}
    [DllImport("user32.dll",EntryPoint="GetWindowLongPtrW")]private static extern IntPtr GetWindowLongPtr64(IntPtr window,int index);
    [DllImport("user32.dll",EntryPoint="GetWindowLongW")]private static extern int GetWindowLong32(IntPtr window,int index);
    [DllImport("user32.dll")]private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]private static extern IntPtr GetCapture();
    [DllImport("user32.dll")]private static extern IntPtr SendMessage(IntPtr window,int message,IntPtr wParam,IntPtr lParam);
    [DllImport("user32.dll")]private static extern int GetGuiResources(IntPtr process,int flags);
}
