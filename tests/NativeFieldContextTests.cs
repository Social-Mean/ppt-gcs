using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Web.Script.Serialization;
using PptGcs.AddIn;

public static class NativeFieldContextTests
{
    private static int count;
    public static int Main()
    {
        try
        {
            Equal(NativeFieldContextState.Matched,NativeFieldContextPolicy.Observe(()=>true));
            Equal(NativeFieldContextState.Different,NativeFieldContextPolicy.Observe(()=>false));
            Equal(NativeFieldContextState.Unavailable,NativeFieldContextPolicy.Observe(()=>{throw new COMException("Office busy",unchecked((int)0x8001010A));}));
            Equal(NativeFieldContextState.Unavailable,NativeFieldContextPolicy.Observe(()=>{throw new InvalidOperationException("View temporarily unavailable");}));
            Check(NativeFieldContextPolicy.CanCommit(NativeFieldContextState.Matched),"Matched commit rejected");
            Check(!NativeFieldContextPolicy.CanCommit(NativeFieldContextState.Different),"Different context claimed");
            bool rejected=false;try{NativeFieldContextPolicy.CanCommit(NativeFieldContextState.Unavailable);}catch(InvalidOperationException){rejected=true;}
            Check(rejected,"Unavailable context was treated as native pass-through");
            Check(NativeFieldContextPolicy.RetainCachedField(NativeFieldContextState.Unavailable),"Temporary COM failure discarded known field");
            Check(!NativeFieldContextPolicy.RetainCachedField(NativeFieldContextState.Different),"Different context retained");
            Check(NativeFieldContextPolicy.CanVerifyScope(NativeFieldContextState.Matched,true),"Known matched scope not verifiable");
            Check(!NativeFieldContextPolicy.CanVerifyScope(NativeFieldContextState.Matched,false),"Unknown scope counted as empty valid scope");
            Check(!NativeFieldContextPolicy.CanVerifyScope(NativeFieldContextState.Unavailable,true),"Unavailable context permitted verification");
            Check(!NativeFieldContextPolicy.CanVerifyScope(NativeFieldContextState.Different,true),"Foreign context permitted verification");
            bool captureRejected=false;
            try{NativeFieldContextPolicy.CaptureIdentified<object>(()=>{throw new COMException("Identified field owner getter failed");});}
            catch(NativeFieldContextUnavailableException error){captureRejected=error.InnerException is COMException;}
            Check(captureRejected,"Known field capture failure lost its identified-field rejection status");

            var unavailable=new Fixture();unavailable.App.Unavailable=true;
            Check(!(bool)unavailable.Call("Commit",unavailable.Field,"test"),"Real commit handler let native input bypass a failed getter");
            Check(unavailable.Normalizations==0&&unavailable.Window.LayoutReads==0,"Unavailable commit read or normalized a document");
            Set(unavailable.Field,"Handle",new IntPtr(12345));Set(unavailable.Controller,"active",unavailable.Field);
            Check(Object.ReferenceEquals(unavailable.Call("GetField",new IntPtr(12345),false),unavailable.Field),"Failed getter lost known HWND ownership");
            var newCapture=new Fixture();newCapture.App.WindowUnavailable=true;bool newRejected=false;
            try{newCapture.Call("CreateIdentifiedField",IntPtr.Zero,NativeFieldProperty.Width,null,"5.29 cm");}
            catch(NativeFieldContextUnavailableException error){newRejected=error.InnerException is COMException;}
            Check(newRejected&&Get(newCapture.Controller,"active")==null,"New identified field owner failure leaked native input or cached an empty context");

            var unknownScope=new Fixture();
            unknownScope.Call("VerifyFocusLoss",unknownScope.Field);
            Check(unknownScope.Window.LayoutReads==0,"Unknown focus scope read a layout anyway");
            Check(unknownScope.Last["residualPt"]==null&&!(bool)unknownScope.Last["ScopeKnown"],"Unknown focus scope reported zero residual");

            var foreign=new Fixture();foreign.App.CurrentPresentation=new object();Set(foreign.Field,"ScopeKnown",true);
            foreign.Call("VerifyFocusLoss",foreign.Field);
            Check(foreign.Window.LayoutReads==0&&foreign.Last["residualPt"]==null,"Different document was read or reported valid");

            var failedRead=new Fixture();Set(failedRead.Field,"ScopeKnown",true);Set(failedRead.Field,"RelatedIds",new string[0]);
            failedRead.Call("VerifyFocusLoss",failedRead.Field);
            Check(failedRead.Window.LayoutReads==1&&failedRead.Last["residualPt"]==null,"Failed captured-window read reported zero residual");

            // A message-only STATIC belongs solely to this test process. It is
            // never shown, receives no input, and is unrelated to PowerPoint.
            IntPtr handle=CreateWindowEx(0,"STATIC","",0,0,0,0,0,new IntPtr(-3),IntPtr.Zero,IntPtr.Zero,IntPtr.Zero);
            if(handle==IntPtr.Zero)throw new Exception("Cannot create owned message-only test window.");
            try
            {
                var guarded=new Fixture();Set(guarded.Field,"Handle",handle);Set(guarded.Controller,"active",guarded.Field);
                bool sawGuard=false;guarded.App.BeforePresentation=()=>{sawGuard=(bool)Get(guarded.Controller,"tracingDeactivation");if(!sawGuard)throw new Exception("COM called before reentry guard");};
                guarded.Call("TraceDeactivation",Message(handle));
                Check(sawGuard&&!(bool)Get(guarded.Controller,"tracingDeactivation"),"Deactivation guard was late or not cleared");
                var writer=new Fixture();Set(writer.Field,"Handle",handle);Set(writer.Controller,"active",writer.Field);Set(writer.Controller,"applying",true);
                writer.Call("TraceDeactivation",Message(handle));
                Check(writer.App.PresentationReads==0&&writer.Window.LayoutReads==0,"Writer reentry read COM/layout instead of Win32 scalars");

                var mouse=new Fixture();Set(mouse.Field,"Handle",handle);Set(mouse.Controller,"active",mouse.Field);
                Set(mouse.Controller,"fieldMouseWindow",handle);mouse.App.Unavailable=true;
                Check(Object.ReferenceEquals(mouse.Call("FocusedMouseField",handle),mouse.Field),"Known focused mouse field was re-identified through COM");
                Check(mouse.App.PresentationReads==0,"Mouse cache lookup pumped a COM getter before the writer guard");
                mouse.Call("TraceFieldMouse",3,0x201);
                Check(mouse.LastKind=="native-field-mouse-observed"&&(int)mouse.Last["code"]==3&&(bool)mouse.Last["readOnly"],"Known-field preview diagnostics were missing");
                int logs=mouse.LogCount;for(int repeat=0;repeat<12;repeat++)mouse.Call("TraceFieldMouse",3,0x201);
                Check(mouse.LogCount-logs==3,"Repeated queue peeks were not bounded to four scalar diagnostics");
                var releases=(HashSet<int>)Get(mouse.Controller,"blockedMouseUps");object input=Mouse(handle);
                object allowed=Click(mouse.Controller,handle,false,false);
                Check(!(bool)mouse.Call("CachedMouseDown",3,0x201,input),"Successful preview result was not reused");
                Check(!(bool)mouse.Call("CachedMouseDown",3,0x201,input),"Repeated preview changed the successful result");
                Check(!(bool)mouse.Call("CachedMouseDown",0,0x201,input),"Removed DOWN did not reuse the preview result");
                Check(mouse.App.PresentationReads==0&&mouse.Window.LayoutReads==0,"A cached preview/removal re-entered COM or solved again");
                Check(!(bool)mouse.Call("MouseRelease",3,0x202)&&Get(mouse.Controller,"mouseClick")!=null,"Successful preview UP cleared its lifecycle early");
                Check(!(bool)mouse.Call("CachedMouseDown",0,0x201,input)&&Object.ReferenceEquals(Get(mouse.Controller,"mouseClick"),allowed),"A delayed removed DOWN discarded its earlier preview result");
                Check(!(bool)mouse.Call("MouseRelease",0,0x202)&&Get(mouse.Controller,"mouseClick")==null,"Actual UP did not finish a successful click");
                Check(mouse.Call("CachedMouseDown",3,0x201,input)==null,"The next identical-position dirty click could not start");

                object rejectedClick=Click(mouse.Controller,handle,false,false);mouse.Call("CompleteMouseClick",rejectedClick,true);
                Check((bool)mouse.Call("CachedMouseDown",3,0x201,input)&&(bool)mouse.Call("CachedMouseDown",0,0x201,input),"Rejected DOWN was not consistently consumed in preview/removal");
                Check((bool)mouse.Call("MouseRelease",3,0x202)&&(bool)mouse.Call("MouseRelease",3,0x202)&&releases.Contains(0x202),"Repeated preview UP stole the rejected release latch");
                Check((bool)mouse.Call("MouseRelease",0,0x202)&&releases.Count==0,"Actual rejected UP did not clear the pair");
                object previewOnly=Click(mouse.Controller,handle,false,false);mouse.Call("CompleteMouseClick",previewOnly,true);
                mouse.Call("MouseRelease",3,0x202);
                Check(mouse.Call("CachedMouseDown",3,0x201,input)==null&&releases.Count==0,"Missing HC_ACTION UP starved the next physical down");

                object nested=Click(mouse.Controller,handle,false,true);
                Check((bool)mouse.Call("CachedMouseDown",3,0x201,input),"Nested duplicate input escaped during the first preview commit");
                Check((bool)mouse.Call("CompleteMouseClick",nested,false),"Outer DOWN was allowed after its nested duplicate was consumed");
                Check((bool)mouse.Call("MouseRelease",0,0x202)&&releases.Count==0,"Nested duplicate did not finish with a consistent release");

                IntPtr mouseData=Marshal.AllocHGlobal(Marshal.SizeOf(input));
                try
                {
                    Marshal.StructureToPtr(input,mouseData,false);
                    Check(!(bool)mouse.Call("TryHandleFieldMouse",-1,0x201,mouseData),"Negative hook code was intercepted");
                    Check(!(bool)mouse.Call("TryHandleFieldMouse",0,0x201,mouseData),"A field was used after native focus had already left it");
                    Check(mouse.App.PresentationReads==0&&mouse.Window.LayoutReads==0,"Late mouse input attempted a post-focus geometry write");
                    Set(mouse.Controller,"fieldMouseWindow",handle);Set(mouse.Controller,"applying",true);
                    Check((bool)mouse.Call("TryHandleFieldMouse",3,0xA1,mouseData)&&releases.Contains(0xA2),"A reentrant nonclient preview escaped a field write");
                    Set(mouse.Controller,"applying",false);
                    Check((bool)mouse.Call("TryHandleFieldMouse",3,0xA2,mouseData)&&releases.Contains(0xA2),"Nonclient preview removed its release latch");
                    Check((bool)mouse.Call("TryHandleFieldMouse",0,0xA2,mouseData)&&releases.Count==0,"Nonclient removal did not finish its release pair");
                }
                finally{Marshal.FreeHGlobal(mouseData);}
                Set(mouse.Controller,"fieldMouseBusy",true);Check(mouse.Controller.IsApplying,"Host reentry guard did not see field mouse capture");
                Set(mouse.Controller,"fieldMouseBusy",false);Check(!mouse.Controller.IsApplying,"The field mouse guard outlived capture");

                var guardedCommit=new Fixture();Set(guardedCommit.Field,"Handle",handle);Set(guardedCommit.Controller,"active",guardedCommit.Field);
                Check(Object.ReferenceEquals(guardedCommit.Call("FocusedMouseField",handle),guardedCommit.Field),"Cached mouse field was lost");
                bool writerGuard=false;guardedCommit.App.Unavailable=true;
                guardedCommit.App.BeforePresentation=()=>writerGuard=(bool)Get(guardedCommit.Controller,"applying");
                Check(!(bool)guardedCommit.Call("Commit",guardedCommit.Field,"pointer-before-focus-loss")&&writerGuard,"Mouse COM context validation preceded the writer guard");
            }
            finally{DestroyWindow(handle);}
            Console.WriteLine("Native field context tests passed: "+count);return 0;
        }
        catch(Exception error){Console.Error.WriteLine(error);return 1;}
    }
    private sealed class Fixture
    {
        public readonly NativeFieldController Controller;
        public readonly object Field;
        public readonly FakeApplication App;
        public readonly FakeWindow Window=new FakeWindow();
        public Dictionary<string,object> Last;
        public string LastKind;public int LogCount;
        public int Normalizations;
        public Fixture()
        {
            object document=new object();App=new FakeApplication {CurrentPresentation=document,CurrentWindow=Window};
            Controller=(NativeFieldController)FormatterServices.GetUninitializedObject(typeof(NativeFieldController));
            var serializer=new JavaScriptSerializer();Set(Controller,"application",App);Set(Controller,"json",serializer);
            Set(Controller,"log",(Action<string,string>)((kind,body)=>{LastKind=kind;LogCount++;Last=serializer.Deserialize<Dictionary<string,object>>(body);}));
            Set(Controller,"normalizeIdentities",(Action)(()=>Normalizations++));
            Set(Controller,"thread",GetCurrentThreadId());Set(Controller,"blockedMouseUps",new HashSet<int>());
            Type type=typeof(NativeFieldController).GetNestedType("Field",BindingFlags.NonPublic);Field=Activator.CreateInstance(type,true);
            Set(Field,"Presentation",document);Set(Field,"Window",Window);Set(Field,"SlideId",256);Set(Field,"InitialText","");
        }
        public object Call(string method,params object[] args)
        {try{return typeof(NativeFieldController).GetMethod(method,BindingFlags.Instance|BindingFlags.NonPublic).Invoke(Controller,args);}catch(TargetInvocationException error){throw error.InnerException;}}
    }
    public sealed class FakeApplication
    {
        public bool Unavailable,WindowUnavailable;public object CurrentPresentation;public FakeWindow CurrentWindow;public Action BeforePresentation;public int PresentationReads;
        public FakeWindow ActiveWindow {get{if(WindowUnavailable)throw new COMException("ActiveWindow getter unavailable");return CurrentWindow;}}
        public FakePresentations Presentations {get{return new FakePresentations();}}
        public object ActivePresentation {get{PresentationReads++;if(BeforePresentation!=null)BeforePresentation();if(Unavailable)throw new COMException("Office busy");return CurrentPresentation;}}
    }
    public sealed class FakePresentations {public int Count {get{return 1;}}}
    public sealed class FakeWindow
    {
        public int LayoutReads;public int ViewType {get{return 9;}}public FakeView View {get{return new FakeView();}}
        public object Presentation {get{LayoutReads++;throw new InvalidOperationException("Layout read unavailable");}}
    }
    public sealed class FakeView {public FakeSlide Slide {get{return new FakeSlide();}}}
    public sealed class FakeSlide {public int SlideID {get{return 256;}}}
    private static object Mouse(IntPtr handle)
    {
        object value=Activator.CreateInstance(typeof(NativeFieldController).GetNestedType("MouseMessage",BindingFlags.NonPublic));
        Set(value,"Window",handle);return value;
    }
    private static object Click(NativeFieldController controller,IntPtr handle,bool consume,bool inProgress)
    {
        object click=Activator.CreateInstance(typeof(NativeFieldController).GetNestedType("MouseClick",BindingFlags.NonPublic),true);
        Set(click,"Owner",handle);Set(click,"Target",handle);Set(click,"Kind",0x201);Set(click,"Up",0x202);
        Set(click,"Consume",consume);Set(click,"InProgress",inProgress);Set(controller,"mouseClick",click);return click;
    }
    private static object Message(IntPtr handle)
    {
        object value=Activator.CreateInstance(typeof(NativeFieldController).GetNestedType("CallMessage",BindingFlags.NonPublic));
        Set(value,"Window",handle);Set(value,"Message",(uint)0x86);return value;
    }
    private static void Set(object target,string name,object value)
    {target.GetType().GetField(name,BindingFlags.Instance|BindingFlags.Public|BindingFlags.NonPublic).SetValue(target,value);}
    private static object Get(object target,string name)
    {return target.GetType().GetField(name,BindingFlags.Instance|BindingFlags.Public|BindingFlags.NonPublic).GetValue(target);}
    private static void Equal(NativeFieldContextState expected,NativeFieldContextState actual){Check(expected==actual,"Wrong context state");}
    private static void Check(bool test,string message){if(!test)throw new Exception(message);count++;}
    [DllImport("user32.dll",CharSet=CharSet.Unicode)]private static extern IntPtr CreateWindowEx(uint extended,string className,string title,uint style,int x,int y,int width,int height,IntPtr parent,IntPtr menu,IntPtr instance,IntPtr parameter);
    [DllImport("user32.dll")]private static extern bool DestroyWindow(IntPtr handle);
    [DllImport("kernel32.dll")]private static extern uint GetCurrentThreadId();
}
