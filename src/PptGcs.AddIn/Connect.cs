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
        private static readonly string AddInVersion = ReadAddInVersion();

        private static string ReadAddInVersion()
        {
            using (var reader = new StreamReader(typeof(Connect).Assembly.GetManifestResourceStream("PptGcs.Version")))
                return reader.ReadToEnd().Trim();
        }

        private dynamic app;
        private Timer timer;
        private ConstrainedDragController controller;
        private KeyboardNudgeController keyboard;
        private NativeFieldController nativeFields;
        private ShapeIdentityManager identities;
        private ConstraintCommands commands;
        private ParameterCommands parameters;
        private RadialMenuController radial;
        private OfficeIconProvider officeIcons;
        private CanvasConstraintOverlay canvasOverlay;
        private SelectionPaneLayerOrder selectionPaneOrder;
        private NativeMarqueeLockFilter marqueeLocks;
        private IRibbonUI ribbonUi;
        private ICTPFactory paneFactory;
        private readonly Dictionary<long,PaneInstance> panes = new Dictionary<long,PaneInstance>();
        private bool selectingFromPane;
        private bool stopping;
        private bool repairingIdentities;
        private Action<object> selectionChanged, presentationOpened, presentationClosed;
        private delegate void BeforeSaveHandler(object presentation, ref bool cancel);
        private BeforeSaveHandler presentationSaving;
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
            try { Initialize(application,addIn); }
            catch
            {
                try { Stop(); }
                catch(Exception cleanupError) { Log("connection-cleanup-error",json.Serialize(new {error=cleanupError.ToString()})); }
                throw;
            }
        }

        private void Initialize(object application,object addIn)
        {
            Application.EnableVisualStyles();
            stopping=false;
            app = application;
            ((dynamic)addIn).Object = this;
            identities = new ShapeIdentityManager();
            officeIcons = new OfficeIconProvider(application, (id,error)=>Log("office-icon-error",json.Serialize(new { imageMso=id,error })));
            commands = new ConstraintCommands(application, NormalizeIdentities, Log);
            parameters = new ParameterCommands(application);
            nativeFields = new NativeFieldController(application, Log, NormalizeActiveIdentities);
            keyboard = new KeyboardNudgeController(application, Log, NormalizeActiveIdentities);
            keyboard.Finished += OnKeyboardFinished;
            controller = new ConstrainedDragController(application, Log, NormalizeActiveIdentities);
            radial = new RadialMenuController(commands.BuildModel,commands.Preview,
                kind=>ExecuteInteractive(()=>commands.Add(kind,null,null)),
                more=> {var pane=EnsurePane();pane.Pane.Visible=true;if(more)pane.Control.ShowEditor();},
                ()=>(int)app.Presentations.Count==0?null:(object)app.ActiveWindow,NormalizeActiveIdentities,Log,officeIcons.Get);
            canvasOverlay=new CanvasConstraintOverlay(application,commands.BuildModel,officeIcons.Get,Log,RunCanvasAction,
                commands.FindSuggestions,RunSuggestedConstraint);
            selectionPaneOrder = new SelectionPaneLayerOrder(application, ActiveSlideLayers, (type, data) => Log(type, json.Serialize(data)));
            marqueeLocks = new NativeMarqueeLockFilter(application, (type, data) => Log(type, json.Serialize(data)));
            controller.AlternateMouseInput=RouteCanvasMouse;
            controller.AlternateKeyboardInput=RouteCanvasKeyboard;
            controller.AlternateInputActive=()=>radial.IsOpen || keyboard.IsActive || nativeFields.IsApplying;
            controller.ObserveNativeCanvasDown=(x,y,target)=>{ if(!busy && !radial.IsOpen && !nativeFields.IsApplying) marqueeLocks.Begin(x,y,target); };
            controller.CanvasPointerStarting=canvasOverlay.BeginCanvasPointer;
            controller.CanvasGestureFinished=canvasOverlay.CompleteSuggestionEdit;
            selectionChanged = selection => OnIdentityEvent(null);
            presentationOpened = presentation => OnIdentityEvent(presentation);
            presentationClosed = OnPresentationClosed;
            presentationSaving = OnPresentationSaving;
            ComEventsHelper.Combine(application, ApplicationEvents, 2001, selectionChanged);
            ComEventsHelper.Combine(application, ApplicationEvents, 2021, presentationOpened);
            ComEventsHelper.Combine(application, ApplicationEvents, 2004, presentationClosed);
            ComEventsHelper.Combine(application, ApplicationEvents, 2018, presentationSaving);
            for (int i = 1; i <= (int)app.Presentations.Count; i++) OnIdentityEvent((object)app.Presentations[i]);
            timer = new Timer { Interval = 25 };
            timer.Tick += Tick;
            timer.Start();
            Log("connected", json.Serialize(new { version = AddInVersion, powerpoint = (string)app.Version, processId = Process.GetCurrentProcess().Id }));
        }
        public void OnDisconnection(ext_DisconnectMode mode, ref Array custom) { Stop(); }
        public void OnStartupComplete(ref Array custom) { }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { Stop(); }
        private void Stop()
        {
            if(stopping)return;
            stopping=true;
            if (timer != null)
            {
                Cleanup("timer-stop",()=>timer.Stop());
                Cleanup("timer-dispose",()=>timer.Dispose());
                timer=null;
            }
            if(controller!=null)
            {
                controller.AlternateMouseInput=null;controller.AlternateKeyboardInput=null;controller.AlternateInputActive=null;
                controller.ObserveNativeCanvasDown=null;
                controller.CanvasPointerStarting=null;controller.CanvasGestureFinished=null;
                Cleanup("pointer",()=>controller.Dispose());controller=null;
            }
            if(keyboard!=null){keyboard.Finished-=OnKeyboardFinished;Cleanup("keyboard",()=>keyboard.Dispose());keyboard=null;}
            if(nativeFields!=null){Cleanup("native-fields",()=>nativeFields.Dispose());nativeFields=null;}
            if(radial!=null){Cleanup("radial",()=>radial.Dispose());radial=null;}
            selectionPaneOrder = null;
            marqueeLocks = null;
            if(canvasOverlay!=null){Cleanup("canvas",()=>canvasOverlay.Dispose());canvasOverlay=null;}
            ribbonUi=null;
            foreach(var pane in panes.Values.ToArray()) Cleanup("pane",()=>pane.Dispose());
            panes.Clear(); paneFactory = null;
            if(officeIcons!=null){Cleanup("icons",()=>officeIcons.Dispose());officeIcons=null;}
            if (app != null)
            {
                if (selectionChanged != null) Cleanup("selection-event",()=>ComEventsHelper.Remove((object)app, ApplicationEvents, 2001, selectionChanged));
                if (presentationOpened != null) Cleanup("open-event",()=>ComEventsHelper.Remove((object)app, ApplicationEvents, 2021, presentationOpened));
                if (presentationClosed != null) Cleanup("close-event",()=>ComEventsHelper.Remove((object)app, ApplicationEvents, 2004, presentationClosed));
                if (presentationSaving != null) Cleanup("save-event",()=>ComEventsHelper.Remove((object)app, ApplicationEvents, 2018, presentationSaving));
            }
            selectionChanged = null; presentationOpened = null;presentationClosed=null;
            presentationSaving = null;
            Cleanup("disconnect-log",()=>Log("disconnected", "{}"));commands=null;parameters=null;identities=null;app = null;
        }

        private void Cleanup(string resource,Action release)
        {
            try {release();}
            catch(Exception error)
            {
                // A failed resource must not prevent later native hooks and
                // event subscriptions from being released.
                try {Log("cleanup-error",json.Serialize(new {resource,error=error.ToString()}));} catch { }
            }
        }

        public string GetCustomUI(string ribbonId) { return RibbonDefinition.CreateXml(); }
        public void OnRibbonLoad(IRibbonUI ribbon) { ribbonUi=ribbon; }
        public bool GetCanvasConstraintDisplay(IRibbonControl control)
        {return (canvasOverlay==null?CanvasConstraintDisplay.Selection:canvasOverlay.Display)==ParseCanvasDisplay(control.Tag);}
        public void ChangeCanvasConstraintDisplay(IRibbonControl control,bool pressed) { SetCanvasDisplay(control.Tag); }
        public void SetCanvasDisplay(string display)
        {
            if(canvasOverlay==null)throw new InvalidOperationException("画布标记尚未初始化。");
            canvasOverlay.SetDisplay(ParseCanvasDisplay(display));
            if(ribbonUi!=null)ribbonUi.Invalidate();
        }
        private static CanvasConstraintDisplay ParseCanvasDisplay(string display)
        {
            CanvasConstraintDisplay value;
            if(!Enum.TryParse(display,false,out value) || !Enum.IsDefined(typeof(CanvasConstraintDisplay),value) || value.ToString()!=display)
                throw new ArgumentException("未知画布标记显示范围。","display");
            return value;
        }
        public string GetCanvasConstraintState() { return json.Serialize(canvasOverlay.GetState()); }

        public string ReplayCanvasBadgeForTest(string relationId,string gesture)
        {
            if(stopping || app==null || busy || controller.IsDragging || keyboard.IsActive || nativeFields.IsApplying || radial.IsOpen ||
                (int)app.Presentations.Count==0 || (string)app.ActivePresentation.Tags["PPTGCS_TEST"]!="V3")
                throw new InvalidOperationException("Canvas badge replay requires an idle V3 fixture.");
            return json.Serialize(canvasOverlay.ReplayBadgeForTest(relationId,gesture));
        }

        private void RunCanvasAction(CanvasConstraintAction action,ConstraintPaneRelation relation)
        {
            RunCommand(()=>{
                switch(action)
                {
                    case CanvasConstraintAction.Enable: commands.SetEnabled(relation.Id,true);break;
                    case CanvasConstraintAction.Disable: commands.SetEnabled(relation.Id,false);break;
                    case CanvasConstraintAction.Delete: commands.Delete(relation.Id);break;
                    case CanvasConstraintAction.Select:
                    case CanvasConstraintAction.Edit:
                        commands.SelectShapes(relation.MemberIds.ToList());
                        if(action==CanvasConstraintAction.Edit)
                        {
                            PaneInstance pane=EnsurePane();
                            pane.Pane.Visible=true;
                            pane.Control.ActivateRelation(relation.Id,true);
                        }
                        else
                        {
                            PaneInstance pane;
                            if(panes.TryGetValue(WindowIdentity((object)app.ActiveWindow),out pane) && pane.Control!=null)
                            {
                                pane.Control.Render(commands.BuildPaneModel());
                                pane.Control.ActivateRelation(relation.Id,false);
                            }
                        }
                        break;
                }
            });
        }

        public void CreateConstraint(IRibbonControl control)
        { RunCommand(()=>BeginConstraintCore(control.Tag)); }

        private bool RouteCanvasMouse(int kind,int x,int y,IntPtr target)
        {
            if(!busy)selectionPaneOrder.ObserveMouse(kind,target);
            marqueeLocks.ObserveMouse(kind,x,y);
            if(canvasOverlay.IsMenuOpen)return false;
            if(keyboard.HandleMouseInput(kind) || radial.HandleMouse(kind,x,y,target))return true;
            if(busy || nativeFields.IsApplying)return false;
            if(canvasOverlay.HandleMouse(kind,x,y,target))return true;
            canvasOverlay.ObserveEditPointer(kind,target);
            return false;
        }

        private bool RouteCanvasKeyboard(int key,bool down)
        {
            if(down)selectionPaneOrder.Cancel();
            if(down && key==0x1B)marqueeLocks.Cancel();
            if(canvasOverlay.IsMenuOpen)return false;
            if(nativeFields.IsApplying)return true;
            if(radial.IsOpen)return radial.HandleKeyboard(key,down);
            canvasOverlay.ObserveEditKey(key,down);
            return nativeFields.HandleKeyboard(key,down) || keyboard.HandleKeyboard(key,down);
        }

        private void RunSuggestedConstraint(CanvasConstraintOffer offer)
        {
            try
            {
                ExecuteShared(()=>{
                    string id=commands.AcceptSuggestion(offer);
                    if(id!=null)Log("suggestion-accepted",json.Serialize(new {id,kind=offer.Kind.ToString(),members=offer.MemberIds}));
                });
            }
            catch(Exception error) {Log("suggestion-not-applied",json.Serialize(new {message=error.Message}));}
        }

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
                var constraints = document.Constraints.Where(c => c.SlideId == slideId).Select(c => new { c.Id,c.SlideId,Kind=c.Kind.ToString(),c.ShapeIds,c.Value,c.SecondaryValue,c.Enabled,c.ParameterId }).ToList();
                slides.Add(new { slideId, name=(string)slide.Name, kind=(string)slide.Tags["PPTGCS_CASE"], geometry, constraints });
            }
            return json.Serialize(new { presentation=(string)presentation.FullName, storage="presentation-tags", constraintXml=document.ToXml(), parameters=document.Parameters, slides });
        }
        public void ShowStatus(object control) { MessageBox.Show(GetStatus(),"PPT GCS"); }
        public string GetStatus() { return "几何约束 v" + AddInVersion + "\r\n已载入约束：" + constraintCount + "\r\n状态：" + (lastError ?? "运行中") + "\r\n日志：" + logPath; }
        public string GetDiagnostics() { return controller.InspectTargets(); }
        public void Reload() { activeKey = null; }

        public string GetConstraintXml() { return PowerPointLayout.ReadDocument((object)app.ActivePresentation).ToXml(); }
        public string GetPaneState() { return json.Serialize(commands.BuildPaneModel()); }
        public string GetRadialState() { return radial.GetState(); }
        public string GetNativeFieldState() { return nativeFields.GetState(); }
        public string GetOfficeIconDiagnostics() { return json.Serialize(officeIcons.Diagnostics()); }
        public string AddConstraintCommand(string kind,string parameters)
        {
            string id=null;var values=ReadParameters(parameters);
            ExecuteShared(()=>id=commands.Add(kind,values.value,values.second,values.containerId,values.parameterId));return id;
        }
        public string AddContainerConstraintCommand(string kind,double value,string containerId)
        {
            string id=null;
            ExecuteShared(()=>id=commands.Add(kind,value,null,containerId));
            return id;
        }
        public void ResizePageCommand(double width,double height)
        { ExecuteShared(()=>PageResize.Apply((object)app,width,height)); }
        public string CopyConstrainedComponent(int targetSlideId,bool shareParameters)
        {
            IReadOnlyList<string> ids=null;
            ExecuteShared(()=>ids=ComponentCopyCommand.Apply((object)app,targetSlideId,shareParameters));
            return json.Serialize(ids);
        }
        public string AddNamedParameterCommand(string name,double value)
        {
            string id=null;
            ExecuteShared(()=>id=parameters.Add(name,value));
            return id;
        }
        public void UpdateNamedParameterCommand(string id,string name,double value)
        { ExecuteShared(()=>parameters.Update(id,name,value)); }
        public void DeleteNamedParameterCommand(string id)
        { ExecuteShared(()=>parameters.Delete(id)); }
        public void UpdateConstraintCommand(string id,string parameters)
        { var values=ReadParameters(parameters);ExecuteShared(()=>commands.Update(id,values.value,values.second,values.parameterId)); }
        public void SetConstraintEnabled(string id,bool enabled) { ExecuteShared(()=>commands.SetEnabled(id,enabled)); }
        public void DeleteConstraintCommand(string id) { ExecuteShared(()=>commands.Delete(id)); }
        public void SelectConstraintObjects(string idsJson)
        { commands.SelectShapes(json.Deserialize<List<string>>(idsJson));RefreshPane(); }

        private sealed class ConstraintInput
        {
            public double? value { get; set; }
            public double? second { get; set; }
            public string containerId { get; set; }
            public string parameterId { get; set; }
        }

        private ConstraintInput ReadParameters(string parameters)
        {
            return String.IsNullOrWhiteSpace(parameters)?new ConstraintInput():json.Deserialize<ConstraintInput>(parameters)??new ConstraintInput();
        }

        public void CTPFactoryAvailable(ICTPFactory factory) { paneFactory=factory; }
        public void ShowConstraintPane(object control) { RunCommand(()=>EnsurePane().Pane.Visible=true); }
        public void ShowLayerPane(object control)
        {
            RunCommand(() =>
            {
                var pane = EnsureLayerPane();
                pane.LayerPane.Visible = true;
                RenderLayers(pane);
            });
        }

        public void OnLayerArrange(IRibbonControl control, ref bool cancelDefault)
        {
            cancelDefault = false;
            if (stopping || busy || controller == null || controller.IsDragging || keyboard.IsActive ||
                nativeFields.IsApplying) return;
            try
            {
                var slide = ActiveSlideLayers();
                if (slide == null || slide.Stack.Length < 2) return;
                string id = control.Id;
                // If an edit fails, the requested native command must not
                // subsequently move the selection across layer boundaries.
                cancelDefault = true;
                bool handled = false;
                ExecuteShared(() => handled = slide.Arrange((object)app, id));
                cancelDefault = handled;
            }
            catch (Exception error)
            {
                Log("layer-arrange-error", json.Serialize(new { error = error.ToString() }));
                MessageBox.Show(error.Message, "PPT GCS");
            }
        }
        public void ShowPageResize(object control)
        {
            RunCommand(()=>{
                var pane=EnsurePane();
                dynamic page=app.ActivePresentation.PageSetup;
                pane.Pane.Visible=true;
                pane.Control.ShowPageResize((double)page.SlideWidth,(double)page.SlideHeight);
            });
        }
        public void ShowComponentCopy(object control)
        {
            RunCommand(()=>{
                var pane=EnsurePane();
                dynamic window=app.ActiveWindow;
                dynamic presentation=window.Presentation;
                int current=(int)window.View.Slide.SlideID;
                var slides=new List<ConstraintPaneSlide>();
                for(int index=1;index<=(int)presentation.Slides.Count;index++)
                {
                    dynamic slide=presentation.Slides[index];
                    int id=(int)slide.SlideID;
                    slides.Add(new ConstraintPaneSlide {Id=id,Label="第 "+index+" 页 · "+(string)slide.Name+(id==current?"（当前页）":"")});
                }
                pane.Pane.Visible=true;
                pane.Control.ShowCopyEditor(slides,current);
            });
        }
        public void ShowNamedParameters(object control)
        {
            RunCommand(()=>{
                var pane=EnsurePane();
                pane.Pane.Visible=true;
                pane.Control.ShowNamedParameters();
            });
        }
        public string OpenConstraintPane()
        { var pane=EnsurePane();pane.Pane.Visible=true;return GetPaneState(); }
        public string GetPaneDiagnostics()
        {
            long handle=WindowIdentity((object)app.ActiveWindow);PaneInstance pane;
            if(!panes.TryGetValue(handle,out pane) || pane.Pane==null)return json.Serialize(new {factory=paneFactory!=null,window=handle,created=false});
            return json.Serialize(new {factory=paneFactory!=null,window=handle,created=true,visible=pane.Pane.Visible,title=pane.Pane.Title,editorKind=pane.Control.EditorKind,
                selectedRelationId=pane.Control.SelectedRelationId,hasConflict=pane.Control.HasConflict,dock=pane.Pane.DockPosition.ToString(),width=pane.Pane.Width,
                host=(long)((dynamic)pane.Pane.ContentControl).HostHandle,control=pane.Control.Handle.ToInt64()});
        }

        private PaneInstance WindowPanes()
        {
            if(paneFactory==null) throw new InvalidOperationException("Office 尚未提供任务窗格接口。");
            dynamic window=app.ActiveWindow;
            long handle=WindowIdentity((object)window);
            PaneInstance instance;
            if (!panes.TryGetValue(handle, out instance))
            {
                instance = new PaneInstance { Window = (object)window, Presentation = (object)window.Presentation };
                panes.Add(handle, instance);
            }
            return instance;
        }

        private PaneInstance EnsurePane()
        {
            var instance = WindowPanes();
            if (instance.Pane == null)
            {
                long handle = WindowIdentity(instance.Window);
                var ui = new ConstraintPaneControl();
                ui.Bind(new ConstraintPaneActions {
                    Add=(kind,value,second,containerId,parameterId)=>InPaneWindow(handle,()=>ExecuteInteractive(()=>commands.Add(kind,value,second,containerId,parameterId))),
                    Update=(id,value,second,parameterId)=>InPaneWindow(handle,()=>ExecuteInteractive(()=>commands.Update(id,value,second,parameterId))),
                    ResizePage=(width,height)=>InPaneWindow(handle,()=>ResizePageCommand(width,height)),
                    CopyComponent=(slide,share)=>InPaneWindow(handle,()=>CopyConstrainedComponent(slide,share)),
                    SaveParameter=(id,name,value)=>{
                        string saved=id;
                        InPaneWindow(handle,()=>{
                            if(id==null)saved=AddNamedParameterCommand(name,value);
                            else UpdateNamedParameterCommand(id,name,value);
                        });
                        return saved;
                    },
                    DeleteParameter=id=>InPaneWindow(handle,()=>DeleteNamedParameterCommand(id)),
                    SetEnabled=(id,enabled)=>InPaneWindow(handle,()=>ExecuteInteractive(()=>commands.SetEnabled(id,enabled))),
                    Delete=id=>InPaneWindow(handle,()=>ExecuteShared(()=>commands.Delete(id))),
                    SelectShapes=ids=>InPaneWindow(handle,()=> { if(selectingFromPane)return;selectingFromPane=true;try{commands.SelectShapes(ids);}finally{selectingFromPane=false;}RefreshPane(); }),
                    Refresh=()=>InPaneWindow(handle,RefreshPane),
                    ResolveConflict=(draft,disabled)=>InPaneWindow(handle,()=>ExecuteInteractive(()=>commands.ResolveConflict(draft,disabled)))
                });
                ui.SetIcons(officeIcons.Get);
                instance.Pane = AttachPane(instance.Window, "约束", ui, 380);
                instance.Control = ui;
            }
            instance.Control.Render(commands.BuildPaneModel());
            return instance;
        }

        private PaneInstance EnsureLayerPane()
        {
            var owner = WindowPanes();
            if (owner.LayerPane != null) return owner;
            long handle = WindowIdentity(owner.Window);
            var ui = new LayerPaneControl();
            ui.ActivateRow = row => InPaneWindow(handle, () =>
            {
                var slide = RequireLayerSlide();
                if (!MatchesLayerRow(slide, row)) return;
                owner.ActiveLayers[slide.SlideId] = row.LayerId;
                if (row.ShapeId.HasValue) slide.SelectObject(row.LayerId, row.ShapeId.Value);
                RenderLayers(owner);
            });
            ui.ToggleState = (row, visibility) => InPaneWindow(handle, () => ExecuteShared(() =>
            {
                var slide = RequireLayerSlide();
                if (MatchesLayerRow(slide, row))
                    slide.ToggleObjectState((object)app, row.LayerId, row.ShapeId, visibility);
            }));
            ui.AddLayer = name => InPaneWindow(handle, () => ExecuteShared(() =>
            {
                var slide = RequireLayerSlide();
                owner.ActiveLayers[slide.SlideId] = slide.Add((object)app, name);
            }));
            ui.RenameLayer = name => InPaneWindow(handle, () => ExecuteShared(() =>
            {
                var slide = RequireLayerSlide();
                slide.Rename((object)app, ActiveLayer(slide), name);
            }));
            ui.RemoveLayer = id => InPaneWindow(handle, () => ExecuteShared(() => RequireLayerSlide().Remove((object)app, id)));
            ui.ReorderLayer = (id, index) => InPaneWindow(handle, () => ExecuteShared(() =>
            {
                var slide = RequireLayerSlide();
                if (!slide.Stack.Any(layer => layer.Id == id)) return;
                owner.ActiveLayers[slide.SlideId] = id;
                slide.Reorder((object)app, id, index);
            }));
            ui.DropObject = (source, target, groupId, after) => InPaneWindow(handle, () => ExecuteShared(() =>
            {
                var slide = RequireLayerSlide();
                if (!MatchesLayerRow(slide, source) || !MatchesLayerRow(slide, target)) return;
                if (groupId.HasValue)
                    slide.ReorderGroupMember((object)app, source.LayerId, groupId.Value, source.ShapeId.Value, target.ShapeId.Value, after);
                else if (slide.MoveObject((object)app, source.LayerId, source.ShapeId.Value, target.LayerId, target.ShapeId, after))
                    owner.ActiveLayers[slide.SlideId] = target.LayerId;
            }));
            ui.AssignSelection = () => InPaneWindow(handle, () => ExecuteShared(() =>
            {
                var slide = RequireLayerSlide();
                slide.Assign((object)app, ActiveLayer(slide));
            }));
            owner.LayerPane = AttachPane(owner.Window, "图层", ui, 320);
            owner.Layers = ui;
            return owner;
        }

        private CustomTaskPane AttachPane(object window, string title, Control ui, int width)
        {
            CustomTaskPane pane = null;
            try
            {
                pane = paneFactory.CreateCTP("PptGcs.ConstraintPaneHost", title, window);
                ((dynamic)pane.ContentControl).AttachWindow(ui.Handle.ToInt64());
                pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                pane.Width = width;
                return pane;
            }
            catch { PaneInstance.DisposePane(pane, ui); throw; }
        }

        private static bool MatchesLayerRow(SlideLayers slide, LayerPaneRow row)
        { return slide.SlideId == row.SlideId && slide.Stack.Any(layer => layer.Id == row.LayerId); }

        private SlideLayers ActiveSlideLayers()
        { return app == null || (int)app.Presentations.Count == 0 ? null : SlideLayers.Read((object)app.ActiveWindow); }

        private SlideLayers RequireLayerSlide()
        { return ActiveSlideLayers() ?? throw new InvalidOperationException("请切换到普通幻灯片编辑视图。"); }

        private string ActiveLayer(SlideLayers slide)
        {
            PaneInstance pane;
            string id;
            return panes.TryGetValue(WindowIdentity(slide.Window), out pane) && pane.ActiveLayers.TryGetValue(slide.SlideId, out id) &&
                slide.Stack.Any(layer => layer.Id == id) ? id : slide.Stack[0].Id;
        }

        private void RenderLayers(PaneInstance pane)
        {
            var slide = ActiveSlideLayers();
            pane.Layers.Render(slide, slide == null ? null : ActiveLayer(slide));
        }

        private void RefreshPane()
        {
            if(app==null || (int)app.Presentations.Count==0)return;
            try
            {
                long handle=WindowIdentity((object)app.ActiveWindow);
                PaneInstance pane;
                if (panes.TryGetValue(handle, out pane))
                {
                    if (pane.Pane != null && pane.Pane.Visible) pane.Control.Render(commands.BuildPaneModel());
                    if (pane.LayerPane != null && pane.LayerPane.Visible) RenderLayers(pane);
                }
            }
            catch(Exception exception) {lastError=exception.Message;}
        }

        private void ExecuteShared(Action command)
        {
            if(busy || (nativeFields!=null && nativeFields.IsApplying) || (controller!=null && controller.IsDragging) || (keyboard!=null && keyboard.IsActive) || (radial!=null && radial.IsOpen)) throw new InvalidOperationException("请先结束当前操作。");
            busy=true;
            try {command();activeKey=null;lastError=null;}
            finally {busy=false;if(canvasOverlay!=null)canvasOverlay.Invalidate();RefreshPane();}
        }

        private void ExecuteInteractive(Action command)
        {
            try {ExecuteShared(command);}
            catch(ConstraintCommands.DraftConflict conflict) when(conflict.Choices.Any(choice=>choice.CanDisable))
            {
                var pane=EnsurePane();
                pane.Pane.Visible=true;
                pane.Control.ShowConflict(conflict);
                throw;
            }
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
            if(canvasOverlay!=null)canvasOverlay.HideForContextChange();
            long key=WindowIdentity(presentation);
            foreach(var item in panes.Where(p=>WindowIdentity(p.Value.Presentation)==key).ToList())
            {item.Value.Dispose();panes.Remove(item.Key);}
        }

        private void OnPresentationSaving(object presentation, ref bool cancel)
        {
            if (stopping || busy || cancel) return;
            busy = true;
            try { SlideLayers.EnsurePresentationOrder(presentation); }
            catch (Exception error) { Log("layer-save-order-error", json.Serialize(new { error = error.ToString() })); }
            finally { busy = false; }
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
            public CustomTaskPane LayerPane;
            public LayerPaneControl Layers;
            public readonly Dictionary<int, string> ActiveLayers = new Dictionary<int, string>();
            private bool disposed;
            public void Dispose()
            {
                if(disposed)return;disposed=true;
                try { DisposePane(Pane, Control); }
                finally { DisposePane(LayerPane, Layers); }
            }

            public static void DisposePane(CustomTaskPane pane, Control view)
            {
                try {if(pane!=null)((dynamic)pane.ContentControl).DetachWindow();} catch(COMException) { }
                catch(InvalidComObjectException) { }
                if(view!=null) view.Dispose();
                try {if(pane!=null)pane.Delete();} catch(COMException) { }
                catch(InvalidComObjectException) { }
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
            if (stopping || busy || repairingIdentities || app == null || (nativeFields != null && nativeFields.IsApplying) || (controller != null && controller.IsDragging) || (keyboard != null && keyboard.IsActive)) return;
            try
            {
                if (presentation == null) marqueeLocks.Finish(true);
                if (presentation == null) NormalizeActiveIdentities();
                else NormalizeIdentities(presentation);
                if(canvasOverlay!=null)canvasOverlay.Invalidate();
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
        {return ReplayForTest(()=>controller.ReplayGestureForTest(driverId,mode,json.Deserialize<double[][]>(offsetsJson)));}

        public string ReplaySelectionMoveForTest(string shapeNamesJson, string offsetsJson)
        {return ReplayForTest(()=>controller.ReplaySelectionMoveForTest(json.Deserialize<string[]>(shapeNamesJson),json.Deserialize<double[][]>(offsetsJson)));}

        public string ReplayKeyboardNudgeForTest(int key, int repeats, bool control)
        {return ReplayForTest(()=>keyboard.ReplayForTest(key,repeats,control));}

        public string ReplayGroupChildGestureForTest(string shapePathJson, string mode, string offsetsJson)
        {return ReplayForTest(()=>controller.ReplayGroupChildGestureForTest(json.Deserialize<string[]>(shapePathJson),mode,json.Deserialize<double[][]>(offsetsJson)));}

        public string ReplayGroupChildrenMoveForTest(string shapePathsJson, string offsetsJson)
        {return ReplayForTest(()=>controller.ReplayGroupChildrenMoveForTest(json.Deserialize<string[][]>(shapePathsJson),json.Deserialize<double[][]>(offsetsJson)));}

        private string ReplayForTest(Action replay)
        {
            if (stopping || app==null || busy || controller.IsDragging || keyboard.IsActive || nativeFields.IsApplying || radial.IsOpen)
                throw new InvalidOperationException("The add-in is busy.");
            if ((int)app.Presentations.Count==0 || (string)app.ActivePresentation.Tags["PPTGCS_TEST"] != "V3")
                throw new InvalidOperationException("Replay requires an explicitly marked V3 test fixture.");
            busy=true;
            try {replay();return GetSnapshot();}
            finally { busy=false; }
        }

        private void OnKeyboardFinished()
        {
            activeKey=null;
            if(canvasOverlay!=null)canvasOverlay.Invalidate();
        }

        private void Tick(object sender, EventArgs args)
        {
            if(canvasOverlay!=null)canvasOverlay.CheckVisibility();
            if (busy || refreshing || app == null || stopping || nativeFields.IsApplying || canvasOverlay.IsMenuOpen) return;
            refreshing = true;
            try
            {
                nativeFields.Tick();
                keyboard.Tick();
                controller.Tick();
                radial.Tick();
                marqueeLocks.Finish();
                busy = true;
                try { selectionPaneOrder.Tick(); }
                finally { busy = false; }
                canvasOverlay.Tick(controller.ActiveGeometry ?? keyboard.ActiveGeometry,
                    controller.ActiveDriverId ?? keyboard.ActiveDriverId,radial.IsOpen,controller.ActiveRestrictions);
                if (controller.IsDragging || keyboard.IsActive || radial.IsOpen || (activeKey != null && DateTime.UtcNow < nextRefresh)) return;
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
            if (busy || (nativeFields!=null && nativeFields.IsApplying) || (controller != null && controller.IsDragging) || (keyboard!=null && keyboard.IsActive) || (radial!=null && radial.IsOpen)) return;
            try { ExecuteInteractive(action); }
            catch (ConstraintCommands.DraftConflict conflict) when(conflict.Choices.Any(choice=>choice.CanDisable))
            { Log("command-conflict",json.Serialize(new { proposed=conflict.Attempted, relations=conflict.Choices })); }
            catch (Exception ex) { Log("command-error",json.Serialize(new { error=ex.ToString() }));MessageBox.Show(ex.Message,"PPT GCS"); }
        }
        private void Log(string type,string data)
        {
            try { Directory.CreateDirectory(Path.GetDirectoryName(logPath));File.AppendAllText(logPath,DateTime.UtcNow.ToString("O")+" "+type+" "+data+Environment.NewLine); }
            catch(IOException) { }
        }
    }
}
