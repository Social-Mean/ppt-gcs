using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    [ComVisible(true), Guid("1B728F4B-5AF0-41FD-8589-86E44A7A5501"), ProgId("PptGcs.Connect"), ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class Connect : StandardOleMarshalObject, IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer
    {
        private dynamic app;
        private Timer timer;
        private ConstrainedDragController controller;
        private ShapeIdentityManager identities;
        private ConstraintCommands commands;
        private RadialMenuController radial;
        private OfficeIconProvider officeIcons;
        private ICTPFactory paneFactory;
        private readonly Dictionary<long,PaneInstance> panes = new Dictionary<long,PaneInstance>();
        private bool selectingFromPane;
        private bool stopping;
        private bool repairingIdentities;
        private Action<object> selectionChanged, presentationOpened, presentationClosed;
        // EApplication event interface and DISPIDs from Office's MSPPT.OLB.
        private static readonly Guid ApplicationEvents = new Guid("914934C2-5A91-11CF-8700-00AA0060263B");
        private bool busy;
        private bool refreshing;
        private DateTime nextRefresh;
        private string activeKey;
        private string lastError;
        private int constraintCount;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();
        private readonly string logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PptGcs", "layout-v2.log");

        public void OnConnection(object application, ext_ConnectMode mode, object addIn, ref Array custom)
        {
            Application.EnableVisualStyles();
            stopping=false;
            app = application;
            ((dynamic)addIn).Object = this;
            identities = new ShapeIdentityManager();
            officeIcons = new OfficeIconProvider(application, (id,error)=>Log("office-icon-error",json.Serialize(new { imageMso=id,error })));
            commands = new ConstraintCommands(application, NormalizeIdentities, Log);
            controller = new ConstrainedDragController(application, Log, NormalizeActiveIdentities);
            radial = new RadialMenuController(commands.BuildModel,commands.Preview,
                kind=>ExecuteShared(()=>commands.Add(kind,null,null)),
                more=> {var pane=EnsurePane();pane.Pane.Visible=true;if(more)pane.Control.ShowEditor();},
                ()=>(int)app.Presentations.Count==0?null:(object)app.ActiveWindow,NormalizeActiveIdentities,Log,officeIcons.Get);
            controller.AlternateMouseInput=radial.HandleMouse;
            controller.AlternateKeyboardInput=radial.HandleKeyboard;
            controller.AlternateInputActive=()=>radial!=null && radial.IsOpen;
            selectionChanged = selection => OnIdentityEvent(null);
            presentationOpened = presentation => OnIdentityEvent(presentation);
            presentationClosed = OnPresentationClosed;
            ComEventsHelper.Combine(application, ApplicationEvents, 2001, selectionChanged);
            ComEventsHelper.Combine(application, ApplicationEvents, 2021, presentationOpened);
            ComEventsHelper.Combine(application, ApplicationEvents, 2004, presentationClosed);
            for (int i = 1; i <= (int)app.Presentations.Count; i++) OnIdentityEvent((object)app.Presentations[i]);
            timer = new Timer { Interval = 25 };
            timer.Tick += Tick;
            timer.Start();
            Log("connected", json.Serialize(new { version = "0.5.2", powerpoint = (string)app.Version, processId = Process.GetCurrentProcess().Id }));
        }
        public void OnDisconnection(ext_DisconnectMode mode, ref Array custom) { Stop(); }
        public void OnStartupComplete(ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { Stop(); }
        private void Stop()
        {
            if (timer == null && controller == null) return;
            if(stopping)return;
            stopping=true;
            if (timer != null) { timer.Stop(); timer.Dispose(); timer = null; }
            if(controller!=null)
            {
                controller.AlternateMouseInput=null;controller.AlternateKeyboardInput=null;controller.AlternateInputActive=null;
                controller.Dispose();controller=null;
            }
            if(radial!=null){radial.Dispose();radial=null;}
            foreach(var pane in panes.Values) pane.Dispose();
            panes.Clear(); paneFactory = null;
            if(officeIcons!=null){officeIcons.Dispose();officeIcons=null;}
            if (app != null)
            {
                if (selectionChanged != null) ComEventsHelper.Remove((object)app, ApplicationEvents, 2001, selectionChanged);
                if (presentationOpened != null) ComEventsHelper.Remove((object)app, ApplicationEvents, 2021, presentationOpened);
                if (presentationClosed != null) ComEventsHelper.Remove((object)app, ApplicationEvents, 2004, presentationClosed);
            }
            selectionChanged = null; presentationOpened = null;presentationClosed=null;
            Log("disconnected", "{}");commands=null;identities=null;app = null;
        }

        public string GetCustomUI(string ribbonId) { return RibbonDefinition.CreateXml(); }

        public void CreateConstraint(IRibbonControl control)
        { RunCommand(()=>BeginConstraintCore(control.Tag)); }

        public void BeginConstraint(string kind)
        { ExecuteShared(()=>BeginConstraintCore(kind)); }

        private void BeginConstraintCore(string kindName)
        {
            ConstraintKind kind=RibbonDefinition.ParseKind(kindName);
            if(RibbonDefinition.RequiresParameterEditor(kind))
            {
                var pane=EnsurePane();
                pane.Pane.Visible=true;
                pane.Control.ShowEditor(kind.ToString());
            }
            else commands.Add(kind.ToString(),null,null);
        }
        public void UnlinkSelected(object control)
        { RunCommand(() => commands.DeleteForSelection()); }

        public void CreateDemo(object control)
        {
            RunCommand(() => {
                dynamic presentation = app.Presentations.Add();
                presentation.Tags.Add("PPTGCS_TEST","V3");
                presentation.PageSetup.SlideWidth = 960f; presentation.PageSetup.SlideHeight = 540f;
                var kinds = new[] { ConstraintKind.SameY, ConstraintKind.SameX, ConstraintKind.SameWidth,
                    ConstraintKind.SameHeight, ConstraintKind.EqualHorizontalGap, ConstraintKind.EqualVerticalGap };
                string[] titles = { "顶端对齐", "左对齐", "同宽 · 拉伸联动", "同高 · 拉伸联动", "水平等间距 · 按边缘距离", "垂直等间距 · 按边缘距离" };
                string[] hints = { "从内部上下拖动 A，再上下拖动 B。", "从内部左右拖动 A，再左右拖动 B。",
                    "先选中 A，拖动右边中点缩放；再对 B 做同样操作。", "先选中 A，拖动底边中点缩放；再对 B 做同样操作。",
                    "从内部左右拖动 A，再左右拖动 C。", "从内部上下拖动 A，再上下拖动 C。" };
                for (int index = 0; index < kinds.Length; index++)
                {
                    ConstraintKind kind = kinds[index];
                    dynamic slide = presentation.Slides.Add(index+1, 12);
                    slide.Name = kind.ToString(); slide.Tags.Add("PPTGCS_CASE", kind.ToString());
                    AddText(slide, (index+1) + "  " + titles[index], 28, 18, 900, 38, 24);
                    AddText(slide, hints[index], 28, 57, 900, 30, 15);
                    var shapes = new List<object>();
                    if (kind == ConstraintKind.SameY)
                    {
                        shapes.Add(Rectangle(slide, kind, "A", 180,190,170,100));
                        shapes.Add(Rectangle(slide, kind, "B", 580,260,170,100));
                    }
                    else if (kind == ConstraintKind.SameX)
                    {
                        shapes.Add(Rectangle(slide, kind, "A", 200,150,150,90));
                        shapes.Add(Rectangle(slide, kind, "B", 600,350,170,90));
                    }
                    else if (kind == ConstraintKind.SameWidth)
                    {
                        shapes.Add(Rectangle(slide, kind, "A", 180,200,160,100));
                        shapes.Add(Rectangle(slide, kind, "B", 580,200,220,140));
                    }
                    else if (kind == ConstraintKind.SameHeight)
                    {
                        shapes.Add(Rectangle(slide, kind, "A", 180,200,180,100));
                        shapes.Add(Rectangle(slide, kind, "B", 580,200,220,160));
                    }
                    else if (kind == ConstraintKind.EqualHorizontalGap)
                    {
                        shapes.Add(Rectangle(slide, kind, "A", 120,220,120,90));
                        shapes.Add(Rectangle(slide, kind, "B", 380,220,160,90));
                        shapes.Add(Rectangle(slide, kind, "C", 700,220,100,90));
                    }
                    else
                    {
                        shapes.Add(Rectangle(slide, kind, "A", 420,100,140,60));
                        shapes.Add(Rectangle(slide, kind, "B", 420,245,180,80));
                        shapes.Add(Rectangle(slide, kind, "C", 420,420,120,55));
                    }
                    app.ActiveWindow.View.GotoSlide(index+1);
                    slide.Shapes.Range(shapes.Select(shape => (object)((dynamic)shape).Name).ToArray()).Select();
                    commands.Add(kind.ToString(), null, null);
                }
                app.ActiveWindow.View.GotoSlide(1);
                activeKey = null;
                Log("demo-created", json.Serialize(new { slides = 6, processId = Process.GetCurrentProcess().Id }));
            });
        }

        private static void AddText(dynamic slide, string text, float left, float top, float width, float height, float size)
        {
            dynamic label = slide.Shapes.AddTextbox(1, left, top, width, height);
            label.TextFrame.TextRange.Text = text; label.TextFrame.TextRange.Font.Size = size;
            label.TextFrame.TextRange.Font.Color.RGB = 0x403020;
        }
        private static object Rectangle(dynamic slide, ConstraintKind kind, string letter, float left, float top, float width, float height)
        {
            dynamic shape = slide.Shapes.AddShape(1,left,top,width,height);
            shape.Name = "GCS_" + kind + "_" + letter;
            shape.LockAspectRatio = 0; shape.TextFrame.AutoSize = 0;
            shape.Fill.ForeColor.RGB = letter == "B" ? 0x8B6542 : 0xD88928;
            shape.TextFrame.TextRange.Text = letter;
            shape.TextFrame.TextRange.Font.Size = 26f;
            return shape;
        }

        public string GetSnapshot()
        {
            dynamic presentation = app.ActivePresentation;
            ConstraintDocument document = PowerPointLayout.ReadDocument((object)presentation);
            var slides = new List<object>();
            for (int i=1; i <= (int)presentation.Slides.Count; i++)
            {
                dynamic slide = presentation.Slides[i];
                int slideId = (int)slide.SlideID;
                var geometry = new List<object>();
                for (int j=1; j <= (int)slide.Shapes.Count; j++)
                {
                    dynamic shape = slide.Shapes[j];
                    string id = PowerPointLayout.ReadId((object)shape);
                    if (id == null) continue;
                    geometry.Add(new { Id=id, Name=(string)shape.Name, Left=(double)shape.Left, Top=(double)shape.Top, Width=(double)shape.Width, Height=(double)shape.Height });
                }
                var constraints = document.Constraints.Where(c => c.SlideId == slideId).Select(c => new { c.Id,c.SlideId,Kind=c.Kind.ToString(),c.ShapeIds,c.Value,c.SecondaryValue,c.Enabled }).ToList();
                slides.Add(new { slideId, name=(string)slide.Name, kind=(string)slide.Tags["PPTGCS_CASE"], geometry, constraints });
            }
            return json.Serialize(new { presentation=(string)presentation.FullName, storage="presentation-tags", constraintXml=document.ToXml(), slides });
        }
        public void ShowStatus(object control) { MessageBox.Show(GetStatus(),"PPT GCS"); }
        public string GetStatus() { return "几何约束 v0.5.2\r\n已载入约束：" + constraintCount + "\r\n状态：" + (lastError ?? "运行中") + "\r\n日志：" + logPath; }
        public string GetDiagnostics() { return controller.InspectTargets(); }
        public void Reload() { activeKey = null; }

        public string GetConstraintXml() { return PowerPointLayout.ReadDocument((object)app.ActivePresentation).ToXml(); }
        public string GetPaneState() { return json.Serialize(commands.BuildModel()); }
        public string GetRadialState() { return radial.GetState(); }
        public string GetOfficeIconDiagnostics() { return json.Serialize(officeIcons.Diagnostics()); }
        public object GetCentersCoincidentImage(IRibbonControl control) { return officeIcons.GetCentersCoincidentPicture(); }
        public string AddConstraintCommand(string kind,string parameters)
        {
            string id=null;var values=ReadParameters(parameters);
            ExecuteShared(()=>id=commands.Add(kind,values.Item1,values.Item2));return id;
        }
        public void UpdateConstraintCommand(string id,string parameters)
        { var values=ReadParameters(parameters);ExecuteShared(()=>commands.Update(id,values.Item1,values.Item2)); }
        public void SetConstraintEnabled(string id,bool enabled) { ExecuteShared(()=>commands.SetEnabled(id,enabled)); }
        public void DeleteConstraintCommand(string id) { ExecuteShared(()=>commands.Delete(id)); }
        public void SelectConstraintObjects(string idsJson)
        { commands.SelectShapes(json.Deserialize<List<string>>(idsJson));RefreshPane(); }

        private Tuple<double?,double?> ReadParameters(string parameters)
        {
            var values=String.IsNullOrWhiteSpace(parameters)?new Dictionary<string,double>():json.Deserialize<Dictionary<string,double>>(parameters);
            return Tuple.Create(values.ContainsKey("value")?(double?)values["value"]:null,values.ContainsKey("second")?(double?)values["second"]:null);
        }

        public void CTPFactoryAvailable(ICTPFactory factory) { paneFactory=factory; }
        public void ShowConstraintPane(object control) { RunCommand(()=>EnsurePane().Pane.Visible=true); }
        public string OpenConstraintPane()
        { var pane=EnsurePane();pane.Pane.Visible=true;return GetPaneState(); }
        public string GetPaneDiagnostics()
        {
            long handle=WindowIdentity((object)app.ActiveWindow);PaneInstance pane;
            if(!panes.TryGetValue(handle,out pane))return json.Serialize(new {factory=paneFactory!=null,window=handle,created=false});
            return json.Serialize(new {factory=paneFactory!=null,window=handle,created=true,visible=pane.Pane.Visible,title=pane.Pane.Title,editorKind=pane.Control.EditorKind,
                dock=pane.Pane.DockPosition.ToString(),width=pane.Pane.Width,host=(long)((dynamic)pane.Pane.ContentControl).HostHandle,control=pane.Control.Handle.ToInt64()});
        }

        private PaneInstance EnsurePane()
        {
            if(paneFactory==null) throw new InvalidOperationException("Office 尚未提供任务窗格接口。");
            dynamic window=app.ActiveWindow;
            long handle=WindowIdentity((object)window);
            PaneInstance instance;
            if(!panes.TryGetValue(handle,out instance))
            {
                CustomTaskPane pane=paneFactory.CreateCTP("PptGcs.ConstraintPaneHost","约束",(object)window);
                var ui=new ConstraintPaneControl();
                instance=new PaneInstance {Pane=pane,Control=ui,Window=(object)window,Presentation=(object)window.Presentation};
                ui.Bind(new ConstraintPaneActions {
                    Add=(kind,value,second)=>InPaneWindow(handle,()=>ExecuteShared(()=>commands.Add(kind,value,second))),
                    Update=(id,value,second)=>InPaneWindow(handle,()=>ExecuteShared(()=>commands.Update(id,value,second))),
                    SetEnabled=(id,enabled)=>InPaneWindow(handle,()=>ExecuteShared(()=>commands.SetEnabled(id,enabled))),
                    Delete=id=>InPaneWindow(handle,()=>ExecuteShared(()=>commands.Delete(id))),
                    SelectShapes=ids=>InPaneWindow(handle,()=> { if(selectingFromPane)return;selectingFromPane=true;try{commands.SelectShapes(ids);}finally{selectingFromPane=false;}RefreshPane(); }),
                    Refresh=()=>InPaneWindow(handle,RefreshPane)
                });
                try
                {
                    ((dynamic)pane.ContentControl).AttachWindow(ui.Handle.ToInt64());
                    ui.SetIcons(officeIcons.Get);
                    pane.DockPosition=MsoCTPDockPosition.msoCTPDockPositionRight;
                    pane.Width=380;
                    panes.Add(handle,instance);
                }
                catch {instance.Dispose();throw;}
            }
            instance.Control.Render(commands.BuildModel());
            return instance;
        }

        private void RefreshPane()
        {
            if(app==null || (int)app.Presentations.Count==0)return;
            try
            {
                long handle=WindowIdentity((object)app.ActiveWindow);
                PaneInstance pane;
                if(panes.TryGetValue(handle,out pane)) pane.Control.Render(commands.BuildModel());
            }
            catch(Exception exception) {lastError=exception.Message;}
        }

        private void ExecuteShared(Action command)
        {
            if(busy || (controller!=null && controller.IsDragging) || (radial!=null && radial.IsOpen)) throw new InvalidOperationException("请先结束当前操作。");
            busy=true;
            try {command();activeKey=null;lastError=null;}
            finally {busy=false;RefreshPane();}
        }

        private void InPaneWindow(long owner,Action action)
        {
            if(stopping || app==null || (int)app.Presentations.Count==0 || WindowIdentity((object)app.ActiveWindow)!=owner)
                throw new InvalidOperationException("请先激活此窗格所属的演示文稿窗口。");
            action();
        }

        private static long WindowIdentity(object window)
        {
            IntPtr identity=Marshal.GetIUnknownForObject(window);
            try {return identity.ToInt64();} finally {Marshal.Release(identity);}
        }

        private void OnPresentationClosed(object presentation)
        {
            if(stopping)return;
            long key=WindowIdentity(presentation);
            foreach(var item in panes.Where(p=>WindowIdentity(p.Value.Presentation)==key).ToList())
            {item.Value.Dispose();panes.Remove(item.Key);}
        }

        private void PrunePanes()
        {
            if(panes.Count==0)return;
            var active=new HashSet<long>();
            for(int i=1;i<=(int)app.Windows.Count;i++)active.Add(WindowIdentity((object)app.Windows[i]));
            foreach(var item in panes.Where(p=>!active.Contains(p.Key)).ToList())
            {item.Value.Dispose();panes.Remove(item.Key);}
        }

        private sealed class PaneInstance : IDisposable
        {
            public CustomTaskPane Pane;public ConstraintPaneControl Control;public object Window,Presentation;
            private bool disposed;
            public void Dispose()
            {
                if(disposed)return;disposed=true;
                try {((dynamic)Pane.ContentControl).DetachWindow();} catch(COMException) { }
                catch(InvalidComObjectException) { }
                if(Control!=null) Control.Dispose();
                try {if(Pane!=null)Pane.Delete();} catch(COMException) { }
                catch(InvalidComObjectException) { }
                Window=null;Presentation=null;Pane=null;Control=null;
            }
        }

        private void NormalizeActiveIdentities()
        {
            if (app != null && (int)app.Presentations.Count > 0)
                NormalizeIdentities((object)app.ActivePresentation);
        }

        private void NormalizeIdentities(object presentation)
        {
            if (repairingIdentities) return;
            repairingIdentities = true;
            try
            {
                int changed = identities.Normalize(presentation);
                if (changed > 0)
                {
                    activeKey = null;
                    Log("shape-identities-updated", json.Serialize(new {
                        presentation = (string)((dynamic)presentation).FullName, changed
                    }));
                }
            }
            finally { repairingIdentities = false; }
        }

        private void OnIdentityEvent(object presentation)
        {
            if (stopping || busy || repairingIdentities || app == null || (controller != null && controller.IsDragging)) return;
            try
            {
                if (presentation == null) NormalizeActiveIdentities();
                else NormalizeIdentities(presentation);
                if(!selectingFromPane)RefreshPane();
            }
            catch (Exception exception)
            {
                lastError = exception.Message;
                Log("identity-error", json.Serialize(new { error = exception.ToString() }));
            }
        }

        // Integration harness only: this is programmatic geometry replay, never
        // evidence of real mouse input. Require the explicit test document marker.
        public string ReplayGestureForTest(string driverId, string mode, string offsetsJson)
        {
            if (busy || controller.IsDragging || radial.IsOpen) throw new InvalidOperationException("The add-in is busy.");
            dynamic presentation = app.ActivePresentation;
            if ((string)presentation.Tags["PPTGCS_TEST"] != "V3")
                throw new InvalidOperationException("Replay requires an explicitly marked V3 test fixture.");
            double[][] offsets = json.Deserialize<double[][]>(offsetsJson);
            busy = true;
            try
            {
                controller.ReplayGestureForTest(driverId, mode, offsets);
                return GetSnapshot();
            }
            finally { busy = false; }
        }

        private void Tick(object sender, EventArgs args)
        {
            if (busy || refreshing || app == null || stopping) return;
            refreshing = true;
            try
            {
                controller.Tick();
                radial.Tick();
                if (controller.IsDragging || radial.IsOpen || (activeKey != null && DateTime.UtcNow < nextRefresh)) return;
                nextRefresh = DateTime.UtcNow.AddMilliseconds(200);
                PrunePanes();
                if ((int)app.Presentations.Count == 0) { activeKey=null;constraintCount=0;return; }
                dynamic window = app.ActiveWindow;
                if (window == null || (int)window.ViewType != 9) return;
                // Idle status only needs metadata. A queued native paste may
                // temporarily contain copied tags until SelectionChange arrives;
                // shape validation belongs to the synchronized edit entry points.
                dynamic presentation = window.Presentation;
                int slideId = (int)window.View.Slide.SlideID;
                ConstraintDocument document = PowerPointLayout.ReadDocument((object)presentation);
                string key = (string)presentation.FullName + "/" + slideId + "/" + document.ToXml();
                constraintCount = document.Constraints.Count;
                if (key != activeKey)
                {
                    activeKey=key;
                    Log("document-loaded",json.Serialize(new { presentation=(string)presentation.FullName,slideId,constraints=constraintCount }));
                }
                lastError=null;
                RefreshPane();
            }
            catch (COMException) { }
            catch (Exception ex) { if (lastError!=ex.Message) Log("host-error",json.Serialize(new { error=ex.Message }));lastError=ex.Message; }
            finally { refreshing=false; }
        }
        private void RunCommand(Action action)
        {
            if (busy || (controller != null && controller.IsDragging) || (radial!=null && radial.IsOpen)) return;
            busy=true;
            try { action(); }
            catch (Exception ex) { Log("command-error",json.Serialize(new { error=ex.ToString() }));MessageBox.Show(ex.Message,"PPT GCS"); }
            finally { busy=false;activeKey=null;RefreshPane(); }
        }
        private void Log(string type,string data)
        {
            try { Directory.CreateDirectory(Path.GetDirectoryName(logPath));File.AppendAllText(logPath,DateTime.UtcNow.ToString("O")+" "+type+" "+data+Environment.NewLine); }
            catch(IOException) { }
        }
    }
}
