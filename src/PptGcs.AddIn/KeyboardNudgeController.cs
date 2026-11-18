using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    // Like an owned pointer drag, one held arrow stays inside its originating
    // DOWN callback. Every repeat and follower write shares one native Undo.
    public sealed class KeyboardNudgeController : IDisposable
    {
        private readonly dynamic application;
        private readonly Action<string,string> log;
        private readonly Action normalizeIdentities;
        private readonly uint thread;
        private readonly JavaScriptSerializer json=new JavaScriptSerializer {MaxJsonLength=4*1024*1024};
        private readonly Queue<KeyInput> queued=new Queue<KeyInput>();
        private readonly HashSet<int> ownedKeys=new HashSet<int>();
        private readonly HashSet<int> blockedMouseUps=new HashSet<int>();
        private Nudge state;
        private bool loopRunning,applying,draining,finishing,disposed;
        private string pendingFinish;
        private object lastReport;

        public bool IsActive {get{return state!=null || loopRunning || applying || finishing;}}
        public IReadOnlyList<ShapeGeometry> ActiveGeometry {get{return state==null?null:state.LastFrame.Geometry;}}
        public string ActiveDriverId {get{return state==null?null:state.GroupChild!=null?state.GroupChild.Plan.GroupIds[0]:state.Plan.DriverIds[0];}}
        public event Action Finished;

        public KeyboardNudgeController(object application,Action<string,string> log,Action normalizeIdentities)
        {
            if(application==null)throw new ArgumentNullException("application");
            if(log==null)throw new ArgumentNullException("log");
            if(normalizeIdentities==null)throw new ArgumentNullException("normalizeIdentities");
            this.application=application;this.log=log;this.normalizeIdentities=normalizeIdentities;
            thread=GetCurrentThreadId();
        }

        public bool HandleKeyboard(int key,bool down)
        {
            if(disposed || (state!=null && state.Test))return false;
            bool owned=IsActive;
            try
            {
                RequireThread();
                if(!down)
                {
                    if(!ownedKeys.Remove(key))return false;
                    queued.Enqueue(new KeyInput {Key=key,Down=false});
                    if(!applying)Drain();
                    return true;
                }
                // Capture the modifiers for this dequeued arrow before any
                // Office call can pump later keyboard messages.
                bool shift=QueuedKeyDown(0x10);
                bool altOrWindows=QueuedKeyDown(0x12)||QueuedKeyDown(0x5B)||QueuedKeyDown(0x5C);
                bool control=QueuedKeyDown(0x11);
                if(!KeyboardNudgePolicy.IsArrow(key) || shift || altOrWindows || !HasCanvasFocus())
                {
                    if(IsActive)
                    {
                        if(applying){pendingFinish="native-input";return true;}
                        Drain();Finish(key==0x1B,"native-input");
                    }
                    ownedKeys.Remove(key);
                    return false;
                }
                if(state==null && !loopRunning && !TryStart(key,false))return false;
                owned=true;
                queued.Enqueue(new KeyInput {Key=key,Down=true,Control=control});
                ownedKeys.Add(key);
                if(!loopRunning)RunLoop();
                return true;
            }
            catch(GroupChildCaptureException error)
            {
                SafeLog("keyboard-nudge-capture-rejected",new {key,message=error.Message,
                    reason=error.InnerException==null ? null : error.InnerException.Message});
                if(state!=null)Finish(true,"capture-rejected");
                // Capture raises this type only for a recognized constrained
                // shape selection after the canvas/type/modifier gates passed.
                // Consume its paired UP as well; no geometry or Undo was begun.
                if(down&&KeyboardNudgePolicy.IsArrow(key)){ownedKeys.Add(key);return true;}
                return owned;
            }
            catch(Exception error)
            {
                SafeLog("keyboard-nudge-error",new {error=error.GetType().Name,message=error.Message});
                if(state!=null)Finish(true,"input-error");
                return owned;
            }
        }

        // Call before the pointer/radial router. Only a mouse event arriving
        // reentrantly during a COM write is consumed; its matching UP is paired.
        public bool HandleMouseInput(int kind)
        {
            if(blockedMouseUps.Remove(kind))return true;
            bool begins=kind==0x0201||kind==0x0203||kind==0x0204||kind==0x0206||kind==0x0207||kind==0x0209||kind==0x020A||kind==0x020E;
            if(!begins || state==null)return false;
            if(applying)
            {
                pendingFinish="mouse-input";
                if(kind==0x0201||kind==0x0203)blockedMouseUps.Add(0x0202);
                if(kind==0x0204||kind==0x0206)blockedMouseUps.Add(0x0205);
                if(kind==0x0207||kind==0x0209)blockedMouseUps.Add(0x0208);
                return true;
            }
            Drain();Finish(false,"mouse-input");return false;
        }

        public void Tick()
        {
            RequireThread();
            if(state==null || applying || finishing)return;
            if(!HasNudgeContext(state))Finish(false,"context-changed");
            // A physical release can precede the last queued repeat. Close the
            // Undo entry when Drain receives the matching key-up in queue order.
        }

        public void EndInput(string reason)
        {
            RequireThread();
            if(applying){pendingFinish=reason??"external-input";return;}
            Drain();Finish(false,reason??"external-input");
        }

        public string ReplayForTest(int key,int repeats,bool control)
        {
            RequireThread();
            if(disposed || IsActive)throw new InvalidOperationException("A keyboard gesture is already active or disposed.");
            if(!KeyboardNudgePolicy.IsArrow(key)||repeats<1||repeats>1000)throw new ArgumentOutOfRangeException("key/repeats");
            if((int)application.Presentations.Count==0 ||
                (string)application.ActivePresentation.Tags["PPTGCS_TEST"]!="V3")
                throw new InvalidOperationException("Keyboard replay requires an explicitly marked V3 fixture.");
            if(!TryStart(key,true))throw new InvalidOperationException("Select a constrained shape or members of a constrained root group in the fixture.");
            try
            {
                Nudge captured=state;
                for(int index=0;index<repeats;index++)
                {
                    queued.Enqueue(new KeyInput {Key=key,Down=true,Control=control});
                    Drain();
                    // The fixture has no physical key held down here. A tick
                    // between queued repeats must preserve this transaction.
                    Tick();
                    if(state!=captured)throw new InvalidOperationException("Keyboard history closed before the queued key-up.");
                }
                queued.Enqueue(new KeyInput {Key=key,Down=false});
                Drain();
                if(state!=null)throw new InvalidOperationException("Keyboard history did not close on the queued key-up.");
                return GetState();
            }
            finally{queued.Clear();if(state!=null)Finish(true,"test-interrupted");}
        }

        public string GetState()
        {
            return json.Serialize(state==null ? new {active=false,last=lastReport} :
                (object)new {active=true,id=state.Id,key=state.Key,frames=state.Frames,dx=state.Dx,dy=state.Dy,
                    selectedNativeIds=state.NativeIds,geometry=state.LastFrame.Geometry,extras=state.LastFrame.Extras});
        }

        private bool TryStart(int key,bool test)
        {
            if((int)application.Presentations.Count==0)return false;
            dynamic window=application.ActiveWindow;
            if(window==null || (int)window.ViewType!=9)return false;
            if(test && (int)window.ActivePane.ViewType!=1)
            {
                for(int index=1;index<=(int)window.Panes.Count;index++)
                    if((int)window.Panes[index].ViewType==1){window.Panes[index].Activate();break;}
                if((int)window.ActivePane.ViewType!=1)throw new InvalidOperationException("The fixture slide pane could not be activated.");
            }
            dynamic selection=window.Selection;
            bool activeWindow=test;
            bool canvasFocus=test || CanvasFocus(out activeWindow);
            if(!KeyboardNudgePolicy.CanOwn(key,(int)selection.Type,false,false,canvasFocus,activeWindow))return false;
            normalizeIdentities();
            var context=PowerPointLayout.Read((object)application);
            if(context.Constraints.Count==0)return false;
            bool childSelection=(bool)selection.HasChildShapeRange;
            dynamic range=childSelection ? selection.ChildShapeRange : selection.ShapeRange;
            var selected=Enumerable.Range(1,(int)range.Count).Select(index=>(object)range[index]).ToArray();
            if(selected.Length==0)return false;
            double scaleX=KeyboardNudgePolicy.PixelsPerPoint((double)window.PointsToScreenPixelsX(0f),(double)window.PointsToScreenPixelsX(10000f));
            double scaleY=KeyboardNudgePolicy.PixelsPerPoint((double)window.PointsToScreenPixelsY(0f),(double)window.PointsToScreenPixelsY(10000f));
            if(childSelection)
            {
                var edit=GroupChildNativeEdit.Capture(context,selected,GeometryEditMode.Move,false);
                if(edit==null)return false;
                state=new Nudge {Id=Guid.NewGuid().ToString("D"),Context=context,Key=key,
                    NativeIds=selected.Select(shape=>(int)((dynamic)shape).Id).ToArray(),GroupChild=edit,
                    LastFrame=edit.LastFrame,ScaleX=scaleX,ScaleY=scaleY,Test=test};
                SafeStateLog("keyboard-nudge-start",new {id=state.Id,key,selectedNativeIds=state.NativeIds,
                    driverIds=edit.Plan.GroupIds,before=edit.Plan.Initial.Geometry,extras=edit.Plan.Initial.Extras,scaleX,scaleY,test,groupChild=true});
                return true;
            }
            if(selected.Any(shape=>(int)((dynamic)shape).Child!=0))return false;
            var byNative=context.Shapes.ToDictionary(pair=>(int)((dynamic)pair.Value).Id,pair=>pair.Key);
            var driverIds=new List<string>();var extras=new Dictionary<int,object>();var positions=new List<SelectionMovePosition>();
            var ids=new List<int>();
            foreach(object item in selected)
            {
                int nativeId=(int)((dynamic)item).Id;
                if(ids.Contains(nativeId))throw new InvalidOperationException("The selection contains duplicate native shape IDs.");
                ids.Add(nativeId);string driver;
                if(byNative.TryGetValue(nativeId,out driver))driverIds.Add(driver);
                else{extras.Add(nativeId,item);positions.Add(ReadPosition(item));}
            }
            if(driverIds.Count==0)return false;
            var plan=new SelectionMovePlan(context.Geometry,context.Constraints,driverIds,positions);
            state=new Nudge {Id=Guid.NewGuid().ToString("D"),Context=context,Key=key,NativeIds=ids.ToArray(),
                Extras=extras,Plan=plan,LastFrame=plan.Initial,ScaleX=scaleX,ScaleY=scaleY,Test=test};
            SafeStateLog("keyboard-nudge-start",new {id=state.Id,key,selectedNativeIds=state.NativeIds,
                driverIds=plan.DriverIds,before=plan.Initial.Geometry,extras=plan.Initial.Extras,scaleX,scaleY,test});
            return true;
        }

        private void RunLoop()
        {
            loopRunning=true;
            try
            {
                while(state!=null || queued.Count>0)
                {
                    Drain();
                    if(state==null && queued.Count==0)break;
                    System.Windows.Forms.Application.DoEvents();
                    Drain();Tick();
                    if(state==null && queued.Count==0)break;
                    uint result=MsgWaitForMultipleObjectsEx(0,IntPtr.Zero,15,0x04FF,0x0004);
                    if(result==UInt32.MaxValue)throw new InvalidOperationException("PowerPoint keyboard input wait failed.");
                }
            }
            finally
            {
                loopRunning=false;queued.Clear();
                if(state!=null)Finish(false,"loop-ended");
            }
        }

        private void Drain()
        {
            if(draining || applying || finishing)return;
            draining=true;
            try
            {
                while(queued.Count>0)
                {
                    var input=queued.Dequeue();
                    if(!input.Down){if(state!=null && input.Key==state.Key)Finish(false,"key-up");continue;}
                    if(state!=null && input.Key!=state.Key)Finish(false,"direction-changed");
                    if(state==null && !TryStart(input.Key,false)){queued.Clear();ownedKeys.Clear();break;}
                    if(!HasNudgeContext(state))
                    {Finish(false,"context-changed");queued.Clear();ownedKeys.Clear();break;}
                    ApplyStep(input.Control);
                    if(pendingFinish!=null){string reason=pendingFinish;pendingFinish=null;Finish(false,reason);queued.Clear();ownedKeys.Clear();break;}
                }
            }
            finally{draining=false;}
        }

        private void ApplyStep(bool control)
        {
            Nudge current=state;
            if(current==null)return;
            dynamic presentation=current.Context.Presentation;
            bool snap=(int)presentation.SnapToGrid!=0;
            double grid=(double)presentation.GridDistance;
            var delta=KeyboardNudgePolicy.Delta(current.Key,
                KeyboardNudgePolicy.StepPoints(snap,grid,control,current.ScaleX),
                KeyboardNudgePolicy.StepPoints(snap,grid,control,current.ScaleY));
            double dx=current.Dx+delta.Item1,dy=current.Dy+delta.Item2;
            SelectionMoveFrame requested=null;
            IReadOnlyList<SelectionMovePosition> children=null;
            try
            {
                if(current.GroupChild!=null)
                {
                    children=current.GroupChild.Plan.Target(dx,dy);
                    current.GroupChild.ThrowIfProvablyImpossible(children);
                }
                else requested=current.Plan.Solve(dx,dy);
            }
            catch(ConstraintConflictException error)
            {
                if(current.Rejection!=error.Message)SafeStateLog("keyboard-nudge-rejected",new {id=current.Id,key=current.Key,message=error.Message,dx,dy});
                current.Rejection=error.Message;return;
            }
            if(current.GroupChild!=null ? GroupChildEditPlan.SameChildren(current.LastFrame.Extras,children) : current.LastFrame.Matches(requested))return;
            applying=true;
            try
            {
                if(!current.UndoStarted){application.StartNewUndoEntry();current.UndoStarted=true;}
                current.LastFrame=current.GroupChild!=null ? current.GroupChild.Apply(children) :
                    current.Plan.ApplyAtomic(current.LastFrame,requested,frame=>WriteFrame(current,frame),()=>ReadFrame(current));
                current.Dx=dx;current.Dy=dy;current.Frames++;current.Rejection=null;
                SafeStateLog("keyboard-nudge-frame",new {id=current.Id,key=current.Key,frame=current.Frames,dx,dy,control,snapToGrid=snap,gridDistance=grid,
                    residualPt=LayoutSolver.MaxResidual(current.LastFrame.Geometry,
                        LayoutSolver.ConnectedConstraints(current.Context.Constraints,DriverIds(current))),geometry=current.LastFrame.Geometry,
                    extras=current.LastFrame.Extras,groupChild=current.GroupChild!=null});
            }
            catch(SelectionMoveRejectedException error)
            {
                current.Rejection=error.Message;
                SafeStateLog("keyboard-nudge-write-rejected",new {id=current.Id,message=error.Message});
            }
            catch(GroupChildEditRejectedException error)
            {
                current.Rejection=error.InnerException==null ? error.Message : error.InnerException.Message;
                SafeStateLog("keyboard-nudge-write-rejected",new {id=current.Id,message=current.Rejection,groupChild=true});
            }
            finally{applying=false;}
        }

        private void Finish(bool cancel,string reason)
        {
            if(state==null || finishing)return;
            Nudge ended=state;state=null;finishing=true;
            try
            {
                if(cancel && ended.UndoStarted)
                {
                    WriteFrame(ended,Initial(ended));
                    if(!ReadFrame(ended).Matches(Initial(ended)))throw new InvalidOperationException("Cancelled keyboard edit did not restore its complete selection.");
                    ended.LastFrame=Initial(ended);
                }
                if(ended.GroupChild!=null && ended.UndoStarted && SameContextAndSelection(ended))ended.GroupChild.SelectExact();
                if(ended.UndoStarted)application.StartNewUndoEntry();
                lastReport=new {id=ended.Id,key=ended.Key,reason,cancelled=cancel,frames=ended.Frames,undoStarted=ended.UndoStarted,
                    dx=ended.Dx,dy=ended.Dy,selectedNativeIds=ended.NativeIds,before=Initial(ended).Geometry,beforeExtras=Initial(ended).Extras,
                    after=ended.LastFrame.Geometry,afterExtras=ended.LastFrame.Extras,test=ended.Test,groupChild=ended.GroupChild!=null};
                SafeLog(ended.Test ? "test-keyboard-nudge-end" : "keyboard-nudge-end",lastReport);
            }
            catch(Exception error)
            {
                SafeLog("keyboard-nudge-end-error",new {id=ended.Id,reason,message=error.Message});
                try{if(ended.UndoStarted)application.StartNewUndoEntry();}catch{}
                if(ended.Test)throw;
            }
            finally
            {
                finishing=false;
                try{if(Finished!=null)Finished();}catch(Exception error){SafeLog("keyboard-nudge-notify-error",new {message=error.Message});}
            }
        }

        private bool SameContextAndSelection(Nudge current)
        {
            if(current==null)return false;
            if((int)application.Presentations.Count==0)return false;
            dynamic window=application.ActiveWindow;
            if(window==null || !SameComObject((object)window,current.Context.Window) || (int)window.ViewType!=9 ||
                (int)window.View.Slide.SlideID!=current.Context.SlideId)return false;
            dynamic selection=window.Selection;
            if((int)selection.Type!=2 || (bool)selection.HasChildShapeRange!=(current.GroupChild!=null))return false;
            dynamic range=current.GroupChild!=null ? selection.ChildShapeRange : selection.ShapeRange;
            var ids=Enumerable.Range(1,(int)range.Count).Select(index=>(int)range[index].Id).ToArray();
            return ids.Length==current.NativeIds.Length && new HashSet<int>(ids).SetEquals(current.NativeIds);
        }

        private bool HasCanvasFocus()
        {bool activeWindow;return CanvasFocus(out activeWindow)&&activeWindow;}

        private bool HasNudgeContext(Nudge current)
        {return SameContextAndSelection(current) && (current.Test || HasCanvasFocus());}

        private bool CanvasFocus(out bool activeWindow)
        {
            activeWindow=false;
            IntPtr focus=GetFocus();uint process;
            if(focus==IntPtr.Zero || GetWindowThreadProcessId(focus,out process)!=thread || process!=GetCurrentProcessId())return false;
            var name=new StringBuilder(64);GetClassName(focus,name,name.Capacity);
            activeWindow=GetAncestor(focus,2)==GetAncestor(GetForegroundWindow(),2);
            return name.ToString()=="mdiClass";
        }

        private static SelectionMovePosition ReadPosition(object shapeObject)
        {
            dynamic shape=shapeObject;
            return new SelectionMovePosition((int)shape.Id,(double)shape.Left,(double)shape.Top,(double)shape.Width,(double)shape.Height);
        }
        private static SelectionMoveFrame ReadFrame(Nudge current)
        {return current.GroupChild!=null ? current.GroupChild.ReadFrame() : new SelectionMoveFrame(PowerPointLayout.Snapshot(current.Context),current.Extras.Values.Select(ReadPosition));}
        private static void WriteFrame(Nudge current,SelectionMoveFrame frame)
        {
            if(current.GroupChild!=null){current.GroupChild.Restore(frame);return;}
            PowerPointLayout.Apply(current.Context,frame.Geometry);
            foreach(var position in frame.Extras)
            {
                dynamic shape=current.Extras[position.NativeId];
                if(Math.Abs((double)shape.Left-position.Left)>.0001)shape.Left=(float)position.Left;
                if(Math.Abs((double)shape.Top-position.Top)>.0001)shape.Top=(float)position.Top;
            }
        }
        private static SelectionMoveFrame Initial(Nudge current)
        {return current.GroupChild!=null ? current.GroupChild.Plan.Initial : current.Plan.Initial;}
        private static IEnumerable<string> DriverIds(Nudge current)
        {return current.GroupChild!=null ? (IEnumerable<string>)current.GroupChild.Plan.GroupIds : current.Plan.DriverIds;}
        private static bool SameComObject(object first,object second)
        {
            IntPtr left=Marshal.GetIUnknownForObject(first);
            try{IntPtr right=Marshal.GetIUnknownForObject(second);try{return left==right;}finally{Marshal.Release(right);}}
            finally{Marshal.Release(left);}
        }
        private void SafeStateLog(string kind,object value){SafeLog(state!=null&&state.Test ? "test-"+kind : kind,value);}
        private void SafeLog(string kind,object value){try{log(kind,json.Serialize(value));}catch{}}
        private void RequireThread(){if(GetCurrentThreadId()!=thread)throw new InvalidOperationException("Keyboard edits must run on PowerPoint's UI thread.");}
        private static bool QueuedKeyDown(int key){return (GetKeyState(key)&0x8000)!=0;}
        public void Dispose(){RequireThread();if(disposed)return;disposed=true;queued.Clear();Finish(false,"dispose");ownedKeys.Clear();blockedMouseUps.Clear();}

        private sealed class KeyInput {public int Key;public bool Down,Control;}
        private sealed class Nudge
        {
            public string Id,Rejection;
            public LayoutContext Context;
            public int Key,Frames;
            public int[] NativeIds;
            public Dictionary<int,object> Extras;
            public SelectionMovePlan Plan;
            public GroupChildNativeEdit GroupChild;
            public SelectionMoveFrame LastFrame;
            public double ScaleX,ScaleY,Dx,Dy;
            public bool UndoStarted,Test;
        }
        [DllImport("user32.dll")]private static extern IntPtr GetFocus();
        [DllImport("user32.dll")]private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]private static extern IntPtr GetAncestor(IntPtr window,uint flags);
        [DllImport("user32.dll")]private static extern uint GetWindowThreadProcessId(IntPtr window,out uint process);
        [DllImport("user32.dll",CharSet=CharSet.Unicode)]private static extern int GetClassName(IntPtr window,StringBuilder name,int count);
        [DllImport("user32.dll")]private static extern short GetKeyState(int key);
        [DllImport("user32.dll",SetLastError=true)]private static extern uint MsgWaitForMultipleObjectsEx(uint count,IntPtr handles,uint timeout,uint mask,uint flags);
        [DllImport("kernel32.dll")]private static extern uint GetCurrentThreadId();
        [DllImport("kernel32.dll")]private static extern uint GetCurrentProcessId();
    }
}
