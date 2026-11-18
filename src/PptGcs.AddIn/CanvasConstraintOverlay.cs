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
    internal enum CanvasConstraintAction { Select, Edit, Enable, Disable, Delete }

    // The layered window is transparent to input. The PowerPoint thread hook
    // forwards badge clicks here, while the canvas handles all other input.
    internal sealed class CanvasConstraintOverlay : IDisposable
    {
        private readonly dynamic application;
        private readonly Func<ConstraintPaneModel> readModel;
        private readonly Func<string,int,Image> getIcon;
        private readonly Action<string,string> log;
        private readonly Action<CanvasConstraintAction,ConstraintPaneRelation> action;
        private readonly System.Windows.Forms.Timer singleClickTimer;
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
        private string highlightedRelationId;
        private IReadOnlyCollection<string> limitingRelationIds=Array.Empty<string>();
        private BadgePress press;
        private BadgeClick pendingClick;
        private bool disposed, interactive;
        private int renders, modelReads;
        public CanvasConstraintDisplay Display { get; private set; } = CanvasConstraintDisplay.Selection;
        public bool IsMenuOpen { get; private set; }

        public CanvasConstraintOverlay(object application, Func<ConstraintPaneModel> readModel,
            Func<string,int,Image> getIcon, Action<string,string> log,
            Action<CanvasConstraintAction,ConstraintPaneRelation> action)
        {
            this.application=application;this.readModel=readModel;this.getIcon=getIcon;this.log=log;this.action=action;
            singleClickTimer=new System.Windows.Forms.Timer {Interval=SystemInformation.DoubleClickTime};
            singleClickTimer.Tick+=(sender,args)=>CompleteSingleClick();
        }

        public void SetDisplay(CanvasConstraintDisplay value)
        {
            RequireThread();
            if(!Enum.IsDefined(typeof(CanvasConstraintDisplay),value))throw new ArgumentOutOfRangeException("value");
            Display=value;Invalidate();
            if(value==CanvasConstraintDisplay.Hidden)Hide("hidden-by-user");
        }

        public void Invalidate()
        { RequireThread();nextModelRefresh=DateTime.MinValue;lastSignature=null;interactive=false;CancelPendingClick(); }

        public void HideForContextChange()
        { RequireThread();if(IsMenuOpen)EndMenu();Hide("context-changed");Invalidate(); }

        public bool HandleMouse(int message, int x, int y, IntPtr targetWindow)
        {
            const int move=0x0200, leftDown=0x0201, leftUp=0x0202, leftDouble=0x0203;
            const int rightDown=0x0204, rightUp=0x0205, rightDouble=0x0206;
            var point=new Point(x,y);
            if(message==move && press!=null)
            {
                if(!Within(point,press.Click.Point,SystemInformation.DragSize))press.Cancelled=true;
                return false;
            }
            if(message==leftUp || message==rightUp)
            {
                if(press==null || press.Left!=(message==leftUp))return false;
                BadgePress completed=press;press=null;
                BadgeClick hit=FindClick(point,targetWindow);
                if(completed.Cancelled || !Within(point,completed.Click.Point,SystemInformation.DragSize) || !SameBadge(completed.Click,hit))
                {
                    if(completed.DoubleClick)CancelPendingClick();
                    return true;
                }
                if(!completed.Left)
                {
                    Defer(()=>ShowMenu(hit));
                    return true;
                }
                if(completed.DoubleClick)
                {
                    CancelPendingClick();
                    Defer(()=>Invoke(CanvasConstraintAction.Edit,hit));
                }
                else if(hit.HasParameters)
                {
                    pendingClick=hit;
                    singleClickTimer.Start();
                }
                else Defer(()=>Invoke(CanvasConstraintAction.Select,hit));
                return true;
            }
            bool left=message==leftDown || message==leftDouble;
            if(!left && message!=rightDown && message!=rightDouble)return false;
            // A new down for the same button proves an earlier release went
            // to another window or process. Never swallow this fresh click's up.
            if(press!=null && press.Left==left)press=null;
            BadgeClick click=press==null && Control.ModifierKeys==Keys.None ? FindClick(point,targetWindow) : null;
            if(click!=null && !IsBlankCanvas(click))click=null;
            bool doubleClick=left && SameBadge(pendingClick,click) &&
                Within(point,pendingClick.Point,SystemInformation.DoubleClickSize);
            CancelPendingClick();
            if(click==null)return false;
            press=new BadgePress {Click=click,Left=left,DoubleClick=doubleClick};
            return true;
        }

        private BadgeClick FindClick(Point screenPoint,IntPtr targetWindow)
        {
            if(disposed || !interactive || window==null || !window.Visible || layout==null ||
                IsMenuOpen || GetAncestor(GetForegroundWindow(),2)!=root ||
                (targetWindow!=canvas && !IsChild(canvas,targetWindow)))return null;
            var clientPoint=new Point(screenPoint.X-window.Left,screenPoint.Y-window.Top);
            CanvasConstraintBadge badge=layout.BadgeAt(clientPoint);
            var relation=badge==null ? null : model.Relations.FirstOrDefault(item=>item.Id==badge.RelationId);
            if(relation==null)return null;
            return new BadgeClick {Root=root,WindowIdentity=identity,PresentationIdentity=presentationIdentity,SlideId=slideId,
                RelationId=relation.Id,Point=screenPoint,HasParameters=relation.Value.HasValue || relation.Second.HasValue};
        }

        private static bool Within(Point point,Point center,Size tolerance)
        { return Math.Abs(point.X-center.X)*2<=tolerance.Width && Math.Abs(point.Y-center.Y)*2<=tolerance.Height; }

        private static bool SameBadge(BadgeClick first,BadgeClick second)
        {
            return first!=null && second!=null && first.Root==second.Root && first.WindowIdentity==second.WindowIdentity &&
                first.PresentationIdentity==second.PresentationIdentity && first.SlideId==second.SlideId && first.RelationId==second.RelationId;
        }

        private void CancelPendingClick()
        { singleClickTimer.Stop();pendingClick=null; }

        private void CompleteSingleClick()
        {
            BadgeClick target=pendingClick;
            CancelPendingClick();
            if(target!=null)Invoke(CanvasConstraintAction.Select,target);
        }

        private void Defer(Action callback)
        { if(window!=null && window.IsHandleCreated)window.BeginInvoke(callback); }

        private void ShowMenu(BadgeClick click)
        {
            ConstraintPaneRelation relation=Resolve(click);
            if(relation==null || IsMenuOpen)return;
            const uint toggle=1,remove=2;
            const uint rightButton=0x0002,nonotify=0x0080,returnCommand=0x0100;
            IntPtr menu=CreatePopupMenu();
            if(menu==IntPtr.Zero)return;
            uint command=0;
            try
            {
                if(!AppendMenu(menu,0,new UIntPtr(toggle),relation.Enabled?"停用约束":"启用约束") ||
                    !AppendMenu(menu,0,new UIntPtr(remove),"删除约束"))return;
                IsMenuOpen=true;
                command=TrackPopupMenuEx(menu,rightButton|nonotify|returnCommand,click.Point.X,click.Point.Y,click.Root,IntPtr.Zero);
            }
            finally{IsMenuOpen=false;DestroyMenu(menu);}
            if(command==toggle)Invoke(relation.Enabled?CanvasConstraintAction.Disable:CanvasConstraintAction.Enable,click);
            if(command==remove)Invoke(CanvasConstraintAction.Delete,click);
        }

        private ConstraintPaneRelation Resolve(BadgeClick target)
        {
            if(disposed || Display==CanvasConstraintDisplay.Hidden || GetAncestor(GetForegroundWindow(),2)!=target.Root)return null;
            try
            {
                dynamic active=application.ActiveWindow;
                if(!MatchesView((object)active,target))return null;
                return readModel().Relations.FirstOrDefault(item=>item.Id==target.RelationId);
            }
            catch(Exception error)
            { log("canvas-action-read-error",error.Message);return null; }
        }

        private bool IsBlankCanvas(BadgeClick target)
        {
            try
            {
                dynamic active=application.ActiveWindow;
                return MatchesView((object)active,target) && active.RangeFromPoint(target.Point.X,target.Point.Y)==null;
            }
            catch(Exception error)
            { log("canvas-action-hit-error",error.Message);return false; }
        }

        private static bool MatchesView(object windowObject,BadgeClick target)
        {
            if(windowObject==null || Identity(windowObject)!=target.WindowIdentity)return false;
            dynamic active=windowObject;
            return (int)active.ViewType==9 && (int)active.Selection.Type!=3 &&
                Identity((object)active.Presentation)==target.PresentationIdentity && (int)active.View.Slide.SlideID==target.SlideId;
        }

        private void Invoke(CanvasConstraintAction command,BadgeClick target)
        {
            ConstraintPaneRelation relation=Resolve(target);
            if(relation==null)return;
            if(command==CanvasConstraintAction.Select || command==CanvasConstraintAction.Edit)
            {
                highlightedRelationId=relation.Id;
                lastSignature=null;
            }
            action(command,relation);
        }

        public void CheckVisibility()
        {
            RequireThread();
            if(window!=null && window.Visible && (root==IntPtr.Zero || !IsWindow(root) || !IsWindowVisible(root) ||
                IsIconic(root) || GetAncestor(GetForegroundWindow(),2)!=root))Hide("inactive-window");
        }

        public void Tick(IReadOnlyList<PptGcs.Core.ShapeGeometry> dragGeometry, string driverId, bool radialOpen,
            IReadOnlyList<PptGcs.Core.LayoutConstraint> restrictions)
        {
            RequireThread();
            if(disposed)return;
            bool dragging=dragGeometry!=null;
            limitingRelationIds=restrictions?.Select(item=>item.Id).ToArray()??Array.Empty<string>();
            interactive=false;
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
                if(signature==lastSignature && window!=null && window.Visible)
                {interactive=!dragging;return;}
                layout=CanvasConstraintRenderer.Build(model,Display,projection,bounds.Size,highlightedRelationId,limitingRelationIds);
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
                interactive=!dragging;
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
            foreach(var item in model.Objects)
            {
                int nativeId;
                if(item.Id.StartsWith(prefix,StringComparison.Ordinal) && Int32.TryParse(item.Id.Substring(prefix.Length),out nativeId))byNativeId.Add(nativeId,item);
            }
            for(int index=1;index<=(int)slide.Shapes.Count;index++)
            {
                dynamic shape=slide.Shapes[index];ConstraintPaneObject item;
                if(!byNativeId.TryGetValue((int)shape.Id,out item))continue;
                if(item.Geometry==null)
                {
                    float width=(float)shape.Width,height=(float)shape.Height;
                    double angle=(double)shape.Rotation*Math.PI/180;
                    float rotatedWidth=(float)(Math.Abs(Math.Cos(angle))*width+Math.Abs(Math.Sin(angle))*height);
                    float rotatedHeight=(float)(Math.Abs(Math.Sin(angle))*width+Math.Abs(Math.Cos(angle))*height);
                    item.CanvasBounds=new RectangleF((float)shape.Left+(width-rotatedWidth)/2,(float)shape.Top+(height-rotatedHeight)/2,
                        rotatedWidth,rotatedHeight);
                }
                else if(members.Contains(item.Id))tracked.Add(new TrackedShape {Shape=(object)shape,Item=item});
            }
        }

        private string Signature()
        {
            var value=new StringBuilder().Append(Display).Append('|').Append(highlightedRelationId).Append('|')
                .Append(String.Join(",",limitingRelationIds)).Append('|').Append(viewport).Append('|')
                .Append(projection.OriginX.ToString("R",CultureInfo.InvariantCulture)).Append('|')
                .Append(projection.OriginY.ToString("R",CultureInfo.InvariantCulture)).Append('|')
                .Append(projection.ScaleX.ToString("R",CultureInfo.InvariantCulture)).Append('|')
                .Append(projection.ScaleY.ToString("R",CultureInfo.InvariantCulture)).Append('|').Append(projection.DpiScale);
            foreach(var item in model.Objects)
            {
                value.Append('|').Append(item.Id).Append(':').Append(item.Name).Append(':').Append(item.Selected);
                if(item.CanvasBounds.HasValue)value.Append(':').Append(item.CanvasBounds.Value);
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
                highlightedRelationId,menuVisible=IsMenuOpen,
                badges=layout==null?new object[0]:layout.Badges.Select(item=>(object)new {item.RelationId,item.Kind,item.Label,item.Highlighted,item.Limited,
                    bounds=new {item.Bounds.X,item.Bounds.Y,item.Bounds.Width,item.Bounds.Height}}).ToArray()};
        }

        // Exercises this router with recorded message kinds in an owned fixture.
        // WindowFromPoint is a read-only hit test. No messages are sent to Office.
        public object ReplayBadgeForTest(string relationId,string gesture)
        {
            RequireThread();
            if(layout==null || !interactive)throw new InvalidOperationException("The owned test canvas must show badges.");
            var badge=layout.Badges.Single(item=>item.RelationId==relationId);
            var point=new Point(window.Left+(int)(badge.Bounds.Left+badge.Bounds.Width/2),
                window.Top+(int)(badge.Bounds.Top+badge.Bounds.Height/2));
            if(gesture=="outside")point.X=window.Left+1;
            else if(gesture!="click" && gesture!="double-click" && gesture!="right-click" && gesture!="drag-away")
                throw new ArgumentException("Unknown canvas test gesture.","gesture");
            IntPtr target=WindowFromPoint(new NativePoint {X=point.X,Y=point.Y});
            var messages=new List<bool>();
            if(gesture=="right-click")
            {
                messages.Add(HandleMouse(0x0204,point.X,point.Y,target));
                messages.Add(HandleMouse(0x0205,point.X,point.Y,target));
            }
            else
            {
                messages.Add(HandleMouse(0x0201,point.X,point.Y,target));
                if(gesture=="drag-away")messages.Add(HandleMouse(0x0200,point.X+80,point.Y,target));
                messages.Add(HandleMouse(0x0202,point.X,point.Y,target));
                if(gesture=="double-click")
                {
                    messages.Add(HandleMouse(0x0201,point.X,point.Y,target));
                    messages.Add(HandleMouse(0x0202,point.X,point.Y,target));
                }
            }
            return new { replay=true,relationId,gesture,point=new {point.X,point.Y},target=target.ToInt64(),canvas=canvas.ToInt64(),handled=messages };
        }

        private void Hide(string reason)
        {if(window!=null && window.IsHandleCreated)window.Hide();interactive=false;CancelPendingClick();hiddenReason=reason;lastSignature=null;layout=null;}
        private void ClearContext()
        {if(IsMenuOpen)EndMenu();CancelPendingClick();highlightedRelationId=null;model=null;tracked.Clear();layout=null;canvas=IntPtr.Zero;identity=0;presentationIdentity=0;slideId=0;nextModelRefresh=DateTime.MinValue;lastSignature=null;}
        private void RequireThread()
        {if(Thread.CurrentThread.ManagedThreadId!=thread)throw new InvalidOperationException("Canvas overlay requires Office's UI thread.");}
        public void Dispose()
        {
            RequireThread();if(disposed)return;disposed=true;ClearContext();
            if(window!=null){window.Dispose();window=null;}
            singleClickTimer.Dispose();
        }
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
        private sealed class BadgeClick
        {
            public IntPtr Root;
            public long WindowIdentity,PresentationIdentity;
            public int SlideId;
            public string RelationId;
            public Point Point;
            public bool HasParameters;
        }
        private sealed class BadgePress {public BadgeClick Click;public bool Left,DoubleClick,Cancelled;}
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
        [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(NativePoint point);
        [DllImport("user32.dll")] private static extern IntPtr CreatePopupMenu();
        [DllImport("user32.dll",CharSet=CharSet.Unicode)] private static extern bool AppendMenu(IntPtr menu,uint flags,UIntPtr id,string text);
        [DllImport("user32.dll")] private static extern uint TrackPopupMenuEx(IntPtr menu,uint flags,int x,int y,IntPtr owner,IntPtr parameters);
        [DllImport("user32.dll")] private static extern bool DestroyMenu(IntPtr menu);
        [DllImport("user32.dll")] private static extern bool EndMenu();
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window,out uint process);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();
    }
}
