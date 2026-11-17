using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace PptGcs.AddIn
{
    // A passive editing aid. Reads Office only on its UI thread and never
    // creates document shapes, captures input, activates a pane or starts undo.
    internal sealed class CanvasConstraintOverlay : IDisposable
    {
        private readonly dynamic application;
        private readonly Func<ConstraintPaneModel> readModel;
        private readonly Func<string,int,Image> getIcon;
        private readonly Action<string,string> log;
        private readonly int thread = Thread.CurrentThread.ManagedThreadId;
        private readonly uint nativeThread=GetCurrentThreadId(),processId=GetCurrentProcessId();
        private readonly List<TrackedShape> tracked = new List<TrackedShape>();
        private ConstraintPaneModel model;
        private CanvasConstraintLayout layout;
        private OverlayWindow window;
        private IntPtr root, canvas;
        private long identity,presentationIdentity;
        private int slideId;
        private DateTime nextModelRefresh;
        private string lastSignature, lastError, hiddenReason = "not-selected";
        private Rectangle viewport;
        private CanvasConstraintProjection projection;
        private bool disposed;
        private int renders, modelReads;
        public CanvasConstraintDisplay Display { get; private set; } = CanvasConstraintDisplay.Selection;

        public CanvasConstraintOverlay(object application, Func<ConstraintPaneModel> readModel,
            Func<string,int,Image> getIcon, Action<string,string> log)
        { this.application=application;this.readModel=readModel;this.getIcon=getIcon;this.log=log; }

        public void SetDisplay(CanvasConstraintDisplay value)
        {
            RequireThread();
            if(!Enum.IsDefined(typeof(CanvasConstraintDisplay),value))throw new ArgumentOutOfRangeException("value");
            Display=value;Invalidate();
            if(value==CanvasConstraintDisplay.Hidden)Hide("hidden-by-user");
        }

        public void Invalidate()
        { RequireThread();nextModelRefresh=DateTime.MinValue;lastSignature=null; }

        public void HideForContextChange()
        { RequireThread();Hide("context-changed");Invalidate(); }

        public void CheckVisibility()
        {
            RequireThread();
            if(window!=null && window.Visible && (root==IntPtr.Zero || !IsWindow(root) || !IsWindowVisible(root) ||
                IsIconic(root) || GetAncestor(GetForegroundWindow(),2)!=root))Hide("inactive-window");
        }

        public void Tick(IReadOnlyList<PptGcs.Core.ShapeGeometry> dragGeometry, string driverId, bool radialOpen)
        {
            RequireThread();
            if(disposed)return;
            bool dragging=dragGeometry!=null;
            try
            {
                if(Display==CanvasConstraintDisplay.Hidden){Hide("hidden-by-user");return;}
                if(radialOpen){Hide("radial-menu");return;}
                if((int)application.Presentations.Count==0){ClearContext();Hide("no-document");return;}
                dynamic documentWindow=application.ActiveWindow;
                if(documentWindow==null || (int)documentWindow.ViewType!=9 || (int)documentWindow.ActivePane.ViewType!=1)
                {Hide("not-slide-editor");return;}
                IntPtr owner=GetAncestor(GetForegroundWindow(),2);uint ownerProcess;
                if(owner==IntPtr.Zero || GetWindowThreadProcessId(owner,out ownerProcess)!=nativeThread || ownerProcess!=processId ||
                    !IsWindowVisible(owner) || IsIconic(owner))
                {Hide("inactive-window");return;}
                if((int)documentWindow.Selection.Type==3){Hide("editing-text");return;}
                dynamic slide=documentWindow.View.Slide;
                long currentIdentity=Identity((object)documentWindow);
                long currentPresentation=Identity((object)documentWindow.Presentation);
                int currentSlide=(int)slide.SlideID;
                if(owner!=root || identity!=currentIdentity || presentationIdentity!=currentPresentation || slideId!=currentSlide)
                {
                    ClearContext();root=owner;identity=currentIdentity;presentationIdentity=currentPresentation;slideId=currentSlide;
                    if(window!=null){window.Dispose();window=null;}
                }

                double originX=(double)documentWindow.PointsToScreenPixelsX(0f);
                double originY=(double)documentWindow.PointsToScreenPixelsY(0f);
                double scaleX=((double)documentWindow.PointsToScreenPixelsX(100f)-originX)/100.0;
                double scaleY=((double)documentWindow.PointsToScreenPixelsY(100f)-originY)/100.0;
                if(!Finite(originX)||!Finite(originY)||!Finite(scaleX)||!Finite(scaleY)||scaleX<=0||scaleY<=0)
                {Hide("invalid-projection");return;}
                RectangleF slideBounds=new RectangleF((float)originX,(float)originY,
                    (float)((double)documentWindow.Presentation.PageSetup.SlideWidth*scaleX),
                    (float)((double)documentWindow.Presentation.PageSetup.SlideHeight*scaleY));
                if(canvas==IntPtr.Zero || !IsWindow(canvas) || !IsWindowVisible(canvas) || !IsChild(root,canvas))canvas=FindCanvas(root,slideBounds);
                Rectangle bounds;
                if(canvas==IntPtr.Zero || !ClientBounds(canvas,out bounds) || bounds.Width<40 || bounds.Height<40)
                {Hide("canvas-unavailable");return;}
                viewport=bounds;
                uint dpi=GetDpiForWindow(canvas);
                if(dpi==0){Hide("invalid-dpi");return;}
                projection=new CanvasConstraintProjection {OriginX=originX-bounds.Left,OriginY=originY-bounds.Top,
                    ScaleX=scaleX,ScaleY=scaleY,DpiScale=dpi/96f};

                // Keep the full relation/selection read out of the drag loop.
                // During a drag, the existing controller has already written
                // and read back geometry before this Tick is called.
                if(!dragging && (model==null || DateTime.UtcNow>=nextModelRefresh))
                {
                    model=readModel();modelReads++;tracked.Clear();
                    CacheShapes((object)slide);
                    nextModelRefresh=DateTime.UtcNow.AddMilliseconds(250);
                }
                if(model==null){Hide("waiting-for-scene");return;}
                if(dragging)
                {
                    var geometryById=dragGeometry.ToDictionary(item=>item.Id,StringComparer.Ordinal);
                    foreach(var item in model.Objects)
                    {
                        item.Selected=item.Geometry!=null && item.Geometry.Id==driverId;
                        if(item.Geometry==null)continue;
                        PptGcs.Core.ShapeGeometry geometry;
                        if(geometryById.TryGetValue(item.Geometry.Id,out geometry))item.Geometry=geometry;
                    }
                }
                else foreach(var item in tracked)item.Item.Geometry=PowerPointLayout.Geometry(item.Shape,item.Item.Geometry.Id);

                var signature=Signature();
                if(signature==lastSignature && window!=null && window.Visible)return;
                layout=CanvasConstraintRenderer.Build(model,Display,projection,bounds.Size);
                if(layout.Badges.Count==0){Hide("no-visible-relations");return;}
                int imageSize=Math.Max(16,Math.Min(128,(int)Math.Round(16*projection.DpiScale)));
                var icons=new Dictionary<string,Image>(StringComparer.Ordinal);
                foreach(string kind in layout.Badges.Select(item=>item.Kind).Distinct())icons[kind]=getIcon(kind,imageSize);
                if(window==null)window=new OverlayWindow();
                if(GetAncestor(GetForegroundWindow(),2)!=owner){Hide("inactive-window");return;}
                dynamic latest=application.ActiveWindow;
                if(latest==null || Identity((object)latest)!=identity || Identity((object)latest.Presentation)!=presentationIdentity ||
                    (int)latest.View.Slide.SlideID!=slideId){Hide("context-changed");Invalidate();return;}
                window.Render(bounds,layout,icons,projection.DpiScale);
                if(!window.Visible)window.Show(new Owner(root));
                lastSignature=signature;hiddenReason=null;lastError=null;renders++;
            }
            catch(Exception exception)
            {
                Hide("read-failed");nextModelRefresh=DateTime.MinValue;
                if(lastError!=exception.Message)
                {lastError=exception.Message;try{log("canvas-overlay-error",exception.Message);}catch{}}
            }
        }

        private void CacheShapes(object slideObject)
        {
            dynamic slide=slideObject;
            var selected=new HashSet<string>(model.Objects.Where(item=>item.Selected).Select(item=>item.Id));
            var relevant=model.Relations.Where(item=>Display==CanvasConstraintDisplay.All || item.MemberIds.Any(selected.Contains));
            var members=new HashSet<string>(relevant.SelectMany(item=>item.MemberIds),StringComparer.Ordinal);
            var byNativeId=new Dictionary<int,ConstraintPaneObject>();
            string prefix="shape:"+slideId.ToString(CultureInfo.InvariantCulture)+"/";
            foreach(var item in model.Objects.Where(item=>item.Geometry!=null && members.Contains(item.Id)))
            {
                int nativeId;
                if(item.Id.StartsWith(prefix,StringComparison.Ordinal) && Int32.TryParse(item.Id.Substring(prefix.Length),out nativeId))byNativeId.Add(nativeId,item);
            }
            for(int index=1;index<=(int)slide.Shapes.Count;index++)
            {
                object shape=slide.Shapes[index];ConstraintPaneObject item;
                if(byNativeId.TryGetValue((int)((dynamic)shape).Id,out item))tracked.Add(new TrackedShape {Shape=shape,Item=item});
            }
        }

        private string Signature()
        {
            var value=new StringBuilder().Append(Display).Append('|').Append(viewport).Append('|')
                .Append(projection.OriginX.ToString("R",CultureInfo.InvariantCulture)).Append('|')
                .Append(projection.OriginY.ToString("R",CultureInfo.InvariantCulture)).Append('|')
                .Append(projection.ScaleX.ToString("R",CultureInfo.InvariantCulture)).Append('|')
                .Append(projection.ScaleY.ToString("R",CultureInfo.InvariantCulture)).Append('|').Append(projection.DpiScale);
            foreach(var item in model.Objects)
            {
                value.Append('|').Append(item.Id).Append(':').Append(item.Name).Append(':').Append(item.Selected);
                var geometry=item.Geometry;if(geometry==null)continue;
                value.Append(':').Append(geometry.Left.ToString("R",CultureInfo.InvariantCulture)).Append(':').Append(geometry.Top.ToString("R",CultureInfo.InvariantCulture))
                    .Append(':').Append(geometry.Width.ToString("R",CultureInfo.InvariantCulture)).Append(':').Append(geometry.Height.ToString("R",CultureInfo.InvariantCulture));
            }
            foreach(var item in model.Relations)value.Append('|').Append(item.Id).Append(':').Append(item.Kind).Append(':').Append(item.Label).Append(':').Append(item.Enabled).Append(':').Append(item.Status)
                .Append(':').Append(item.Value).Append(':').Append(item.Second).Append(':').Append(String.Join(",",item.MemberIds));
            return value.ToString();
        }

        public object GetState()
        {
            RequireThread();
            return new {display=Display.ToString(),visible=window!=null&&window.Visible,hiddenReason,lastError,renders,modelReads,
                nativeWindow=window==null||!window.IsHandleCreated?0:window.Handle.ToInt64(),canvas=canvas.ToInt64(),
                viewport=new {viewport.X,viewport.Y,viewport.Width,viewport.Height},
                projection=projection==null?null:new {projection.OriginX,projection.OriginY,projection.ScaleX,projection.ScaleY,projection.DpiScale},
                badges=layout==null?new object[0]:layout.Badges.Select(item=>(object)new {item.RelationId,item.Kind,item.Label,
                    bounds=new {item.Bounds.X,item.Bounds.Y,item.Bounds.Width,item.Bounds.Height}}).ToArray()};
        }

        private void Hide(string reason)
        {if(window!=null && window.IsHandleCreated)window.Hide();hiddenReason=reason;lastSignature=null;layout=null;}
        private void ClearContext()
        {model=null;tracked.Clear();layout=null;canvas=IntPtr.Zero;identity=0;presentationIdentity=0;slideId=0;nextModelRefresh=DateTime.MinValue;lastSignature=null;}
        private void RequireThread()
        {if(Thread.CurrentThread.ManagedThreadId!=thread)throw new InvalidOperationException("Canvas overlay requires Office's UI thread.");}
        public void Dispose()
        {RequireThread();if(disposed)return;disposed=true;if(window!=null){window.Dispose();window=null;}ClearContext();}
        private static bool Finite(double value){return !Double.IsNaN(value)&&!Double.IsInfinity(value);}
        private static long Identity(object value)
        {IntPtr unknown=Marshal.GetIUnknownForObject(value);try{return unknown.ToInt64();}finally{Marshal.Release(unknown);}}

        private IntPtr FindCanvas(IntPtr owner,RectangleF slideBounds)
        {
            IntPtr found=IntPtr.Zero;float best=-1;bool tied=false;
            EnumChildWindows(owner,(handle,data)=>{
                var name=new StringBuilder(128);GetClassName(handle,name,name.Capacity);
                Rectangle bounds;uint candidateProcess;
                if(name.ToString()!="mdiClass" || !IsWindowVisible(handle) || !ClientBounds(handle,out bounds) ||
                    GetWindowThreadProcessId(handle,out candidateProcess)!=nativeThread || candidateProcess!=processId)return true;
                var overlap=RectangleF.Intersect(bounds,slideBounds);float score=Math.Max(0,overlap.Width)*Math.Max(0,overlap.Height);
                if(score>best){best=score;found=handle;tied=false;}
                else if(score==best)tied=true;
                return true;
            },IntPtr.Zero);
            return !tied && best>0?found:IntPtr.Zero;
        }
        private static bool ClientBounds(IntPtr handle,out Rectangle bounds)
        {
            Rect rectangle;var point=new NativePoint();bounds=Rectangle.Empty;
            if(!GetClientRect(handle,out rectangle)||!ClientToScreen(handle,ref point))return false;
            bounds=new Rectangle(point.X,point.Y,rectangle.Right-rectangle.Left,rectangle.Bottom-rectangle.Top);return true;
        }

        private sealed class TrackedShape {public ConstraintPaneObject Item;public object Shape;}
        private sealed class Owner : IWin32Window {public IntPtr Handle{get;private set;}public Owner(IntPtr handle){Handle=handle;}}
        private sealed class OverlayWindow : Form
        {
            private Bitmap frame;private LayeredWindowSurface surface;
            public OverlayWindow()
            {FormBorderStyle=FormBorderStyle.None;ShowInTaskbar=false;StartPosition=FormStartPosition.Manual;AutoScaleMode=AutoScaleMode.None;AccessibleName="PPT GCS 画布约束标记";}
            protected override bool ShowWithoutActivation {get{return true;}}
            protected override CreateParams CreateParams
            {get{var value=base.CreateParams;value.ExStyle|=0x00080000|0x08000000|0x00000080|0x00000020;return value;}}
            protected override void WndProc(ref Message message)
            {if(message.Msg==0x0021){message.Result=new IntPtr(3);return;}if(message.Msg==0x0084){message.Result=new IntPtr(-1);return;}base.WndProc(ref message);}
            protected override void OnPaintBackground(PaintEventArgs args){}
            protected override void OnPaint(PaintEventArgs args){}
            public void Render(Rectangle bounds,CanvasConstraintLayout layout,IReadOnlyDictionary<string,Image> icons,float scale)
            {
                Bounds=bounds;if(!IsHandleCreated)CreateHandle();
                if(frame==null || frame.Size!=ClientSize)
                {
                    if(frame!=null)frame.Dispose();if(surface!=null)surface.Dispose();
                    frame=new Bitmap(ClientSize.Width,ClientSize.Height,PixelFormat.Format32bppPArgb);surface=new LayeredWindowSurface(ClientSize);
                }
                using(var graphics=Graphics.FromImage(frame)){graphics.Clear(Color.Transparent);CanvasConstraintRenderer.Draw(graphics,layout,icons,scale);}
                surface.Present(Handle,frame,Location);
            }
            protected override void Dispose(bool disposing)
            {if(disposing){if(surface!=null){surface.Dispose();surface=null;}if(frame!=null){frame.Dispose();frame=null;}}base.Dispose(disposing);}
        }
        [StructLayout(LayoutKind.Sequential)] private struct NativePoint {public int X,Y;}
        [StructLayout(LayoutKind.Sequential)] private struct Rect {public int Left,Top,Right,Bottom;}
        private delegate bool EnumWindow(IntPtr window,IntPtr data);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent,EnumWindow callback,IntPtr data);
        [DllImport("user32.dll",CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr window,StringBuilder name,int count);
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window,out Rect bounds);
        [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr window,ref NativePoint point);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window,uint flags);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent,IntPtr child);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window,out uint process);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();
    }
}
