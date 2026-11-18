using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    // Own a complete pointer gesture on PowerPoint's UI thread. Native drag/resize
    // previews only commit geometry on release, so the plugin applies real shapes.
    public sealed class ConstrainedDragController : IDisposable
    {
        private const int WhMouse = 7, WhKeyboard = 2, HcAction = 0;
        private const int WmMouseMove = 0x0200, WmLeftDown = 0x0201, WmLeftUp = 0x0202, WmLeftDoubleClick = 0x0203;
        private const uint GaRoot = 2;
        private const double HandleRadius = 8.0;
        private readonly dynamic application;
        private readonly Action<string, string> log;
        private readonly Action normalizeIdentities;
        private readonly uint uiThreadId;
        private readonly MouseHook hookDelegate;
        private readonly KeyboardHook keyboardDelegate;
        private readonly JavaScriptSerializer serializer = new JavaScriptSerializer { MaxJsonLength = 4 * 1024 * 1024 };
        private IntPtr hook, keyboardHook;
        private Gesture gesture;
        private bool disposed, inCallback, updating, finishing, pendingLeftUp, gestureLoopRunning;
        private Point? queuedMotion, queuedRelease;

        public bool IsDragging { get { return gesture != null; } }
        internal IReadOnlyList<ShapeGeometry> ActiveGeometry { get { return gesture==null?null:gesture.LastGeometry; } }
        internal string ActiveDriverId { get { return gesture==null?null:gesture.Driver.Id; } }
        internal IReadOnlyList<LayoutConstraint> ActiveRestrictions { get { return gesture?.Restrictions; } }
        public Func<int,int,int,IntPtr,bool> AlternateMouseInput { get; set; }
        public Action<int,int,IntPtr> ObserveNativeCanvasDown { get; set; }
        public Func<int,bool,bool> AlternateKeyboardInput { get; set; }
        public Func<bool> AlternateInputActive { get; set; }
        internal Action<int,int,IntPtr> CanvasPointerStarting { get; set; }
        internal Action<bool> CanvasGestureFinished { get; set; }

        public ConstrainedDragController(object application, Action<string, string> log, Action normalizeIdentities)
        {
            if (application == null) throw new ArgumentNullException("application");
            if (log == null) throw new ArgumentNullException("log");
            if (normalizeIdentities == null) throw new ArgumentNullException("normalizeIdentities");
            this.application = application;
            this.log = log;
            this.normalizeIdentities = normalizeIdentities;
            uiThreadId = GetCurrentThreadId();
            hookDelegate = OnMouse;
            hook = SetWindowsHookEx(WhMouse, hookDelegate, IntPtr.Zero, uiThreadId);
            if (hook == IntPtr.Zero)
                throw new InvalidOperationException("Cannot install PowerPoint mouse hook: " + Marshal.GetLastWin32Error());
            keyboardDelegate = OnKeyboard;
            keyboardHook = SetKeyboardHookEx(WhKeyboard, keyboardDelegate, IntPtr.Zero, uiThreadId);
            if (keyboardHook == IntPtr.Zero)
            {
                int error = Marshal.GetLastWin32Error();
                UnhookWindowsHookEx(hook); hook = IntPtr.Zero;
                throw new InvalidOperationException("Cannot install PowerPoint gesture keyboard guard: " + error);
            }
            SafeLog("gesture-hook-installed", new { thread = uiThreadId, process = GetCurrentProcessId() });
        }

        // Inspect real coordinates, visible shape hits and all eight resize handle
        // locations without synthesizing input or changing the presentation.
        public string InspectTargets()
        {
            RequireUiThread();
            IntPtr root = GetAncestor(GetForegroundWindow(), GaRoot);
            uint process;
            uint thread = GetWindowThreadProcessId(root, out process);
            try
            {
                LayoutContext context = PowerPointLayout.Read((object)application);
                ScreenMap map = ScreenMap.From(context.Window);
                dynamic window = context.Window;
                int selectedId = SelectedShapeId(window);
                var shapes = new List<object>();
                foreach (ShapeGeometry geometry in context.Geometry)
                {
                    dynamic shape = context.Shapes[geometry.Id];
                    var center = map.Point(geometry.Left + geometry.Width / 2, geometry.Top + geometry.Height / 2);
                    IntPtr hitWindow = WindowFromPoint(center);
                    int? rangeId = null;
                    string rangeError = null;
                    try
                    {
                        dynamic hit = window.RangeFromPoint(center.X, center.Y);
                        if (hit != null) rangeId = (int)hit.Id;
                    }
                    catch (Exception exception) { rangeError = exception.Message; }
                    shapes.Add(new
                    {
                        id = geometry.Id, name = (string)shape.Name, nativeId = (int)shape.Id,
                        geometry, selected = selectedId == (int)shape.Id,
                        center = new { x = center.X, y = center.Y },
                        hitWindow = hitWindow.ToInt64(), hitClass = WindowClass(hitWindow),
                        canvas = IsCanvasWindow(hitWindow), rangeId, rangeError,
                        handles = Handles(geometry, map).Select(DescribeHandle).ToArray()
                    });
                }
                return serializer.Serialize(new
                {
                    ownerProcess = GetCurrentProcessId(), ownerThread = uiThreadId,
                    currentThread = GetCurrentThreadId(), foregroundRoot = root.ToInt64(),
                    foregroundClass = WindowClass(root), foregroundProcess = process, foregroundThread = thread,
                    presentation = (string)((dynamic)context.Presentation).FullName,
                    slideId = context.SlideId, map.OriginX, map.OriginY, map.ScaleX, map.ScaleY,
                    constraints = ConstraintData(context.Constraints), shapes
                });
            }
            catch (Exception exception)
            {
                return serializer.Serialize(new
                {
                    ownerProcess = GetCurrentProcessId(), ownerThread = uiThreadId,
                    foregroundRoot = root.ToInt64(), foregroundProcess = process, foregroundThread = thread,
                    error = exception.GetType().Name + ": " + exception.Message
                });
            }
        }

        // Fixture-only integration entry point, exposed through a validating
        // Connect wrapper. Offsets are screen pixels; the final point is release.
        // This exercises the same writer/Undo path without synthesizing input.
        public void ReplayGestureForTest(string driverId, string mode, double[][] offsets)
        {
            RequireUiThread();
            if (disposed || gesture != null || gestureLoopRunning || (AlternateInputActive!=null && AlternateInputActive()))
                throw new InvalidOperationException("A gesture is already active or the controller is disposed.");
            GestureMode parsedMode;
            if (!Enum.TryParse(mode, false, out parsedMode) || !Enum.IsDefined(typeof(GestureMode), parsedMode))
                throw new ArgumentException("Unknown gesture mode.", "mode");
            if (offsets == null || offsets.Length < 2 || offsets.Any(p => p == null || p.Length != 2 ||
                p.Any(v => Double.IsNaN(v) || Double.IsInfinity(v))))
                throw new ArgumentException("Supply motion offsets followed by a release offset.", "offsets");
            normalizeIdentities();
            LayoutContext context = PowerPointLayout.Read((object)application);
            driverId = ShapeIdentity.Normalize(driverId);
            ShapeGeometry driver = context.Geometry.Single(g => g.Id == driverId);
            if (!context.Constraints.Any(c => c.ShapeIds.Contains(driverId)))
                throw new InvalidOperationException("The test driver has no constraints.");
            dynamic window = context.Window;
            ActivateSlidePane(window);
            ScreenMap map = ScreenMap.From(context.Window);
            Point start = parsedMode == GestureMode.Move ? map.Point(driver.Left+driver.Width/2,driver.Top+driver.Height/2) :
                HandlePoint(Handles(driver,map).Single(h => h.Mode == parsedMode));
            ((dynamic)context.Shapes[driverId]).Select();
            StartGesture(context, driver, parsedMode, map, start, IntPtr.Zero, IntPtr.Zero, 96, 8, true);
            ReplayOffsetsForTest(offsets);
        }

        public void ReplaySelectionMoveForTest(string[] nativeShapeNames,double[][] offsets)
        {
            RequireUiThread();
            if(disposed || gesture!=null || gestureLoopRunning || (AlternateInputActive!=null && AlternateInputActive()))
                throw new InvalidOperationException("A gesture is already active or the controller is disposed.");
            if(nativeShapeNames==null || nativeShapeNames.Length<2 || nativeShapeNames.Any(String.IsNullOrWhiteSpace) ||
                nativeShapeNames.Distinct(StringComparer.Ordinal).Count()!=nativeShapeNames.Length)
                throw new ArgumentException("Supply distinct native names for at least two selected shapes.","nativeShapeNames");
            if(offsets==null || offsets.Length<2 || offsets.Any(p=>p==null || p.Length!=2 || p.Any(v=>Double.IsNaN(v)||Double.IsInfinity(v))))
                throw new ArgumentException("Supply motion offsets followed by a release offset.","offsets");
            normalizeIdentities();
            LayoutContext context=PowerPointLayout.Read((object)application);
            dynamic window=context.Window;
            ActivateSlidePane(window);
            var shapes=nativeShapeNames.Select(name=>(object)((dynamic)context.Slide).Shapes.Item(name)).ToArray();
            SelectNativeShapes(shapes);
            var selection=CaptureSelectionMove(context,shapes);
            if(selection==null)throw new InvalidOperationException("The selection has no constrained member.");
            ScreenMap map=ScreenMap.From(context.Window);
            dynamic range=window.Selection.ShapeRange;
            Point start=map.Point((double)range.Left+(double)range.Width/2,(double)range.Top+(double)range.Height/2);
            ShapeGeometry driver=context.Geometry.Single(shape=>shape.Id==selection.Plan.DriverIds[0]);
            StartGesture(context,driver,GestureMode.Move,map,start,IntPtr.Zero,IntPtr.Zero,96,8,true,selection);
            ReplayOffsetsForTest(offsets);
        }

        public void ReplayGroupChildGestureForTest(string[] nativeShapePath,string mode,double[][] offsets)
        {
            ReplayGroupChildrenForTest(new[]{nativeShapePath},mode,offsets);
        }

        public void ReplayGroupChildrenMoveForTest(string[][] nativeShapePaths,double[][] offsets)
        {
            ReplayGroupChildrenForTest(nativeShapePaths,"Move",offsets);
        }

        private void ReplayGroupChildrenForTest(string[][] nativeShapePaths,string mode,double[][] offsets)
        {
            RequireUiThread();
            if(disposed || gesture!=null || gestureLoopRunning || (AlternateInputActive!=null && AlternateInputActive()))
                throw new InvalidOperationException("A gesture is already active or the controller is disposed.");
            if(nativeShapePaths==null || nativeShapePaths.Length==0 || nativeShapePaths.Any(path=>
                path==null || path.Length<2 || path.Any(String.IsNullOrWhiteSpace)))
                throw new ArgumentException("Supply the top-level group and each exact native child path.","nativeShapePaths");
            GestureMode parsedMode;
            if(!Enum.TryParse(mode,false,out parsedMode)||!Enum.IsDefined(typeof(GestureMode),parsedMode))
                throw new ArgumentException("Unknown gesture mode.","mode");
            if(offsets==null || offsets.Length<2 || offsets.Any(p=>p==null || p.Length!=2 || p.Any(v=>Double.IsNaN(v)||Double.IsInfinity(v))))
                throw new ArgumentException("Supply motion offsets followed by a release offset.","offsets");
            normalizeIdentities();
            LayoutContext context=PowerPointLayout.Read((object)application);
            dynamic window=context.Window;
            ActivateSlidePane(window);
            object[] children=nativeShapePaths.Select(path=>{
                dynamic child=((dynamic)context.Slide).Shapes.Item(path[0]);
                foreach(string name in path.Skip(1))child=child.GroupItems.Item(name);
                return (object)child;
            }).ToArray();
            var capture=CaptureGroupChildren(context,children,parsedMode);
            if(capture==null)throw new InvalidOperationException("The native child does not belong to a constrained top-level group.");
            if(parsedMode!=GestureMode.Move && (children.Length!=1 || !ShapeCapabilities.IsSupported(children[0])))
                throw new InvalidOperationException("This native member's resize handles are not supported by the controlled child path.");
            capture.Edit.SelectExact();
            ScreenMap map=ScreenMap.From(context.Window);
            var bounds=ReadNativePosition(children[0]);
            Point start=parsedMode==GestureMode.Move ? map.Point(bounds.Left+bounds.Width/2,bounds.Top+bounds.Height/2) :
                HandlePoint(Handles(new ShapeGeometry(capture.Edit.Plan.GroupIds[0],bounds.Left,bounds.Top,bounds.Width,bounds.Height),map).Single(handle=>handle.Mode==parsedMode));
            ShapeGeometry driver=context.Geometry.Single(shape=>shape.Id==capture.Edit.Plan.GroupIds[0]);
            StartGesture(context,driver,parsedMode,map,start,IntPtr.Zero,IntPtr.Zero,96,8,true,child:capture);
            ReplayOffsetsForTest(offsets);
        }

        private void ReplayOffsetsForTest(double[][] offsets)
        {
            Point start=gesture.StartMouse;
            gestureLoopRunning = true;
            try
            {
                foreach (double[] offset in offsets.Take(offsets.Length-1))
                {
                    QueueOwnedPointer(WmMouseMove, new Point { X=start.X+(int)Math.Round(offset[0]), Y=start.Y+(int)Math.Round(offset[1]) });
                    DrainQueuedPointer();
                }
                double[] release = offsets[offsets.Length-1];
                QueueOwnedPointer(WmLeftUp, new Point { X=start.X+(int)Math.Round(release[0]), Y=start.Y+(int)Math.Round(release[1]) });
                DrainQueuedPointer();
            }
            finally
            {
                gestureLoopRunning = false;
                queuedMotion = null; queuedRelease = null;
                if (gesture != null) Finish(true,"test-interrupted");
            }
        }

        public void Tick()
        {
            RequireUiThread();
            if (disposed || gesture == null || updating || finishing || gestureLoopRunning) return;
            try
            {
                CheckCancellation();
            }
            catch (Exception exception)
            {
                SafeLog("gesture-error", new { error = exception.GetType().Name, message = exception.Message });
                Finish(true, "update-error");
            }
        }

        private static object DescribeHandle(Handle handle)
        {
            var point = new Point { X=(int)Math.Round(handle.X), Y=(int)Math.Round(handle.Y) };
            IntPtr window = WindowFromPoint(point);
            return new { mode=handle.Mode.ToString(), x=handle.X, y=handle.Y,
                hitClass=WindowClass(window), canvas=IsCanvasWindow(window) };
        }

        private IntPtr OnMouse(int code, IntPtr message, IntPtr data)
        {
            // HC_NOREMOVE can repeat the same event. Only act on its removal.
            if (code != HcAction || disposed || (gesture != null && gesture.IsTest))
                return CallNextHookEx(hook, code, message, data);
            int kind = message.ToInt32();
            if (inCallback)
            {
                // COM writes and the owned message loop may pump more input.
                // Never forward our nested move/up to PowerPoint: doing so both
                // loses the final cursor sample and lets native Undo end early.
                if (gesture != null)
                {
                    var nested = (MouseHookData)Marshal.PtrToStructure(data, typeof(MouseHookData));
                    if (QueueOwnedPointer(kind, nested.Position)) return new IntPtr(1);
                }
                else if(AlternateInputActive!=null && AlternateInputActive() && AlternateMouseInput!=null)
                {
                    var nested=(MouseHookData)Marshal.PtrToStructure(data,typeof(MouseHookData));
                    if(AlternateMouseInput(kind,nested.Position.X,nested.Position.Y,nested.Window))return new IntPtr(1);
                }
                return CallNextHookEx(hook, code, message, data);
            }
            bool owned = gesture != null || pendingLeftUp;
            inCallback = true;
            try
            {
                var mouse = (MouseHookData)Marshal.PtrToStructure(data, typeof(MouseHookData));
                if (gesture != null)
                {
                    if (QueueOwnedPointer(kind, mouse.Position)) return new IntPtr(1);
                }
                else if (AlternateMouseInput != null && AlternateMouseInput(kind,mouse.Position.X,mouse.Position.Y,mouse.Window))
                {
                    return new IntPtr(1);
                }
                else if (kind == WmLeftUp && pendingLeftUp)
                {
                    // A cancelled gesture already consumed the matching down.
                    pendingLeftUp = false;
                    return new IntPtr(1);
                }
                else if ((kind == WmLeftDown || kind == WmLeftDoubleClick) && !finishing)
                {
                    // A fresh down also proves a release after a lost capture;
                    // do not consume this new native click's up as the old one.
                    pendingLeftUp = false;
                    if (TryBegin(mouse, kind))
                    {
                        owned = true;
                        RunGestureLoop();
                        return new IntPtr(1);
                    }
                    if (kind == WmLeftDown)
                        ObserveNativeCanvasDown?.Invoke(mouse.Position.X, mouse.Position.Y, mouse.Window);
                }
            }
            catch (Exception exception)
            {
                SafeLog("gesture-hook-error", new { error = exception.GetType().Name, message = exception.Message });
                if (gesture != null) Finish(true, "hook-error");
                if (owned) return new IntPtr(1);
            }
            finally { inCallback = false; }
            return CallNextHookEx(hook, code, message, data);
        }

        private bool QueueOwnedPointer(int kind, Point point)
        {
            if (gesture == null) return false;
            // Preserve crossing the threshold before motion is coalesced. A
            // fast excursion and return to the start is still a drag, not a
            // group-member click, even if Update has not run in between.
            if((kind==WmMouseMove || kind==WmLeftUp) && !gesture.DragThresholdPassed &&
                GroupPointerPolicy.HasDragDistance(point.X-gesture.StartMouse.X,point.Y-gesture.StartMouse.Y,gesture.DragThreshold))
                gesture.DragThresholdPassed=true;
            if (kind == WmMouseMove)
            {
                if (!queuedRelease.HasValue) queuedMotion = point;
                return true;
            }
            if (kind == WmLeftUp)
            {
                if (!queuedRelease.HasValue) queuedRelease = point;
                pendingLeftUp = false;
                // A group click has no geometry/Undo transaction to finish.
                // Commit it before DoEvents can dispatch the next double-click
                // into this old gesture. An UP that crosses the threshold still
                // follows the regular drag release path, even without a MOVE.
                if((gesture.DeferSelection || gesture.SelectionMove!=null || gesture.GroupChild!=null) && !gesture.DragThresholdPassed && !updating && !finishing &&
                    !GroupPointerPolicy.HasDragDistance(point.X-gesture.StartMouse.X,point.Y-gesture.StartMouse.Y,gesture.DragThreshold))
                    DrainQueuedPointer();
                return true;
            }
            // Mouse capture makes these messages part of this PowerPoint gesture;
            // do not open a context menu or scroll/zoom the canvas mid-transaction.
            return kind >= 0x0200 && kind <= 0x020E;
        }

        private IntPtr OnKeyboard(int code, IntPtr key, IntPtr flags)
        {
            if (code == HcAction && !disposed && gesture == null && AlternateKeyboardInput != null &&
                AlternateKeyboardInput(key.ToInt32(),(flags.ToInt64() & 0x80000000L)==0)) return new IntPtr(1);
            if (code != HcAction || disposed || gesture == null || !gestureLoopRunning || gesture.IsTest)
                return CallNextHookEx(keyboardHook, code, key, flags);
            if (key.ToInt32() == 0x1B && (flags.ToInt64() & 0x80000000L) == 0 && !queuedRelease.HasValue)
                gesture.CancelRequested = true;
            // App-local keyboard input cannot invoke Undo/Delete/Ribbon commands
            // inside the gesture's Undo boundary. System task switching is outside
            // this thread hook and is handled by the focus-loss cancellation path.
            return new IntPtr(1);
        }

        private void RunGestureLoop()
        {
            gestureLoopRunning = true;
            try
            {
                // All geometry writes, including the exact release endpoint,
                // stay inside the originating DOWN callback. This avoids a
                // separate native PowerPoint Undo item for a release correction.
                while (gesture != null && !disposed)
                {
                    System.Windows.Forms.Application.DoEvents();
                    if (gesture == null) break;
                    DrainQueuedPointer();
                    if (gesture == null) break;
                    // A queued up is committed before focus/capture polling.
                    CheckCancellation();
                    if (gesture == null) break;
                    uint result = MsgWaitForMultipleObjectsEx(0, IntPtr.Zero, 15, 0x04FF, 0x0004);
                    if (result == UInt32.MaxValue)
                        throw new InvalidOperationException("PowerPoint gesture input wait failed: " + Marshal.GetLastWin32Error());
                }
            }
            finally
            {
                gestureLoopRunning = false;
                queuedMotion = null;
                queuedRelease = null;
                if (gesture != null) Finish(true, "loop-interrupted");
            }
        }

        private void DrainQueuedPointer()
        {
            if (gesture == null) return;
            if (gesture.CancelRequested) { Finish(true, "escape"); return; }
            Point? motion = queuedMotion;
            queuedMotion = null;
            if (motion.HasValue) Update(motion.Value, !gesture.IsTest && KeyDown(1), "move");
            // A COM write above may have queued additional motion and the up.
            // The up's exact position subsumes coalesced intermediate samples.
            if (gesture != null && queuedRelease.HasValue)
            {
                Point release = queuedRelease.Value;
                queuedMotion = null;
                Update(release, false, "up");
                CheckRelease(release);
                Finish(false, "mouse-up");
                queuedRelease = null;
            }
        }

        private void CheckCancellation()
        {
            if (gesture == null) return;
            if (gesture.CancelRequested || KeyDown(0x1B)) { Finish(true, "escape"); return; }
            if (GetAncestor(GetForegroundWindow(), GaRoot) != gesture.RootWindow)
            { Finish(true, "focus-lost"); return; }
            if (GetCapture() != gesture.CaptureWindow)
            { Finish(true, "capture-lost"); return; }
            if (!KeyDown(1) && !queuedRelease.HasValue)
            {
                if (!gesture.PhysicalReleaseSince.HasValue) gesture.PhysicalReleaseSince = DateTime.UtcNow;
                else if ((DateTime.UtcNow-gesture.PhysicalReleaseSince.Value).TotalMilliseconds >= 500)
                {
                    SafeGestureLog(gesture,"gesture-input-error",new { gestureId=gesture.Id, reason="release-message-lost" });
                    Finish(true,"release-message-lost");
                }
            }
            else gesture.PhysicalReleaseSince = null;
        }

        private bool TryBegin(MouseHookData mouse, int messageKind)
        {
            if (messageKind == WmLeftDoubleClick) return false;
            if ((int)application.Presentations.Count == 0) return false;
            dynamic window = application.ActiveWindow;
            if (window == null || (int)window.ViewType != 9) return false;
            bool childSelection=(bool)window.Selection.HasChildShapeRange;
            IntPtr root = GetAncestor(mouse.Window, GaRoot);
            uint process;
            if (root == IntPtr.Zero || root != GetAncestor(GetForegroundWindow(), GaRoot) ||
                GetWindowThreadProcessId(mouse.Window, out process) != uiThreadId ||
                process != GetCurrentProcessId()) return false;
            if (!IsCanvasWindow(mouse.Window))
            { LogRoute(mouse, messageKind, "native-noncanvas"); return false; }
            if (GetCapture() != IntPtr.Zero)
            { LogRoute(mouse, messageKind, "native-capture"); return false; }
            if (KeyDown(0x11) || KeyDown(0x12) || KeyDown(0x10))
            { LogRoute(mouse, messageKind, "native-modifier"); return false; }

            // A thumbnail click leaves ActivePane on the thumbnail pane. Its
            // PointsToScreenPixels methods are invalid until the slide pane is
            // activated. Only the actual mdiClass input surface can reach here;
            // do not change panes for thumbnail, notes, Ribbon or task-pane input.
            int previousPane = ActivateSlidePane(window);
            if(messageKind==WmLeftDown)CanvasPointerStarting?.Invoke(mouse.Position.X,mouse.Position.Y,mouse.Window);

            // Resolve the active document at mouse-down; never carry a cached
            // slide ID or COM object into a different presentation window.
            normalizeIdentities();
            LayoutContext context = PowerPointLayout.Read((object)application);
            if (context.Constraints.Count == 0)
            { LogRoute(mouse, messageKind, "native-unconstrained", ReadSelection(window)); return false; }
            ScreenMap map = ScreenMap.From(context.Window);
            uint dpi = GetDpiForWindow(mouse.Window);
            if (dpi == 0) throw new InvalidOperationException("PowerPoint canvas DPI is unavailable.");
            double handleRadius = HandleRadius * dpi / 96.0;
            if(childSelection)
            {
                bool? childRoute=TryBeginGroupChild(context,map,mouse,messageKind,root,dpi,handleRadius,previousPane);
                if(childRoute.HasValue)return childRoute.Value;
            }
            var constrained = new HashSet<string>(context.Constraints.SelectMany(c => c.ShapeIds), StringComparer.Ordinal);
            var byId = context.Geometry.ToDictionary(item => item.Id, StringComparer.Ordinal);
            // A real unselected-object hit starts its own normal gesture;
            // selected-member hits must never collapse the current selection.
            if(SelectedShapeCount(window)>1)
            {
                bool? selectionRoute=TryBeginSelectionMove(context,map,mouse,messageKind,root,dpi,handleRadius,previousPane);
                if(selectionRoute.HasValue)return selectionRoute.Value;
            }
            ShapeGeometry driver = null;
            GestureMode mode = GestureMode.Move;
            object closest = null;
            object rawHit = null;
            ShapeGeometry selectedBorder = null;

            int selectedId = childSelection ? -1 : SelectedShapeId(window);
            if (selectedId >= 0)
            {
                foreach (var item in context.Shapes)
                {
                    dynamic shape = item.Value;
                    if ((int)shape.Id != selectedId || !constrained.Contains(item.Key) || !CanDrag((object)shape)) continue;
                    closest = ClosestHandle(byId[item.Key], map, mouse.Position);
                    GestureMode? handle = HitHandle(byId[item.Key], map, mouse.Position, handleRadius);
                    if (handle.HasValue) { driver = byId[item.Key]; mode = handle.Value; }
                    else if (TextPointerPolicy.HitMoveBorder((int)shape.Type, (int)window.Selection.Type == 3,
                        ScreenBounds(byId[item.Key], map), new System.Drawing.PointF(mouse.Position.X,mouse.Position.Y), 6 * dpi / 96f) ||
                        GroupPointerPolicy.HitMoveBorder((int)shape.Type,true,ScreenBounds(byId[item.Key],map),
                            new System.Drawing.PointF(mouse.Position.X,mouse.Position.Y),6 * dpi / 96f))
                        selectedBorder = byId[item.Key];
                    break;
                }
            }
            if (driver == null)
            {
                dynamic hit = window.RangeFromPoint(mouse.Position.X, mouse.Position.Y);
                rawHit = (object)hit;
                if (hit != null) hit = ShapeCapabilities.TopLevel((object)hit);
                int? hitId = hit == null ? (int?)null : (int)hit.Id;
                if (selectedBorder != null && TextPointerPolicy.CanUseBorderHit(true,
                    hitId.HasValue ? (bool?)(hitId.Value == selectedId) : null)) driver = selectedBorder;
                if (driver == null && !hitId.HasValue)
                { LogRoute(mouse, messageKind, "native-no-hit", ReadSelection(window), handleRadius, closest, context.Geometry, previousPane); return false; }
                if (driver == null) foreach (var item in context.Shapes)
                {
                    dynamic shape = item.Value;
                    if ((int)shape.Id == hitId && constrained.Contains(item.Key) && CanDrag((object)shape))
                    {
                        ShapeGeometry geometry = byId[item.Key];
                        var bounds = ScreenBounds(geometry,map);
                        var point = new System.Drawing.PointF(mouse.Position.X,mouse.Position.Y);
                        bool isSelected=(int)shape.Id==selectedId;
                        bool editingText = isSelected && (int)window.Selection.Type == 3;
                        bool border = TextPointerPolicy.HitMoveBorder((int)shape.Type,editingText,bounds,point,6 * dpi / 96f) ||
                            GroupPointerPolicy.HitMoveBorder((int)shape.Type,isSelected,bounds,point,6 * dpi / 96f);
                        if (!GroupPointerPolicy.AcceptNativeHit((int)shape.Type,border || HitBody(geometry,map,mouse.Position))) continue;
                        if (TextPointerPolicy.UseNativeTextInput((int)shape.Type, editingText, bounds, point, 6 * dpi / 96f))
                        {
                            LogRoute(mouse, messageKind, "native-text-interior", ReadSelection(window));
                            return false;
                        }
                        driver = geometry; break;
                    }
                }
            }
            if (driver == null)
            { LogRoute(mouse, messageKind, "native-no-driver", ReadSelection(window), handleRadius, closest, context.Geometry, previousPane); return false; }
            LogRoute(mouse, messageKind, "owned-" + mode, ReadSelection(window), handleRadius, closest, context.Geometry, previousPane);
            dynamic selected = context.Shapes[driver.Id];
            bool groupMove=(int)selected.Type==6 && mode==GestureMode.Move;
            // Keep the exact native member hit. A no-drag click can enter the
            // group without ever giving PowerPoint a second, unsolved drag path.
            object clickChild=null;
            if(groupMove && rawHit!=null && (int)((dynamic)rawHit).Child!=0 &&
                (int)((dynamic)ShapeCapabilities.TopLevel(rawHit)).Id==(int)selected.Id)
                clickChild=rawHit;
            if(!groupMove)selected.Select();
            SetFocus(mouse.Window);
            SetCapture(mouse.Window);
            if (GetCapture() != mouse.Window)
            { LogRoute(mouse, messageKind, "capture-failed", ReadSelection(window)); return false; }
            StartGesture(context, driver, mode, map, mouse.Position, root, mouse.Window, dpi, handleRadius, false);
            gesture.DeferSelection=groupMove;
            gesture.ClickChild=clickChild;
            gesture.WasSelectedGroup=groupMove && selectedId==(int)selected.Id;
            if(groupMove)SafeGestureLog(gesture,"group-pointer-candidate",new {gestureId=gesture.Id,
                wasSelectedGroup=gesture.WasSelectedGroup,rawNativeId=rawHit==null ? -1 : (int)((dynamic)rawHit).Id,
                groupNativeId=(int)selected.Id,exactChild=clickChild!=null});
            return true;
        }

        // null means a hit outside the currently edited group; the regular
        // top-level route may handle it. false is an intentional native route.
        private bool? TryBeginGroupChild(LayoutContext context,ScreenMap map,MouseHookData mouse,int messageKind,
            IntPtr root,uint dpi,double handleRadius,int previousPane)
        {
            dynamic window=context.Window;
            dynamic selection=window.Selection;
            dynamic range=selection.ChildShapeRange;
            if((int)range.Count>1)
                return TryBeginGroupChildrenMove(context,map,mouse,messageKind,root,dpi,handleRadius,previousPane,(object)range);
            if((int)range.Count!=1)return false;
            dynamic selectedChild=range[1];
            dynamic selectedGroup=ShapeCapabilities.TopLevel((object)selectedChild);
            var point=new System.Drawing.PointF(mouse.Position.X,mouse.Position.Y);
            var selectedBounds=ScreenBounds(ReadNativePosition((object)selectedChild),map);
            object child=null;
            GestureMode mode=GestureMode.Move;
            if(MultiSelectionPointerPolicy.HitHandle(selectedBounds,point,handleRadius))
            {
                if(!ShapeCapabilities.IsSupported((object)selectedChild))
                {LogRoute(mouse,messageKind,"native-child-endpoint-handle",ReadSelection(window));return false;}
                var box=ReadNativePosition((object)selectedChild);
                string groupId=FindGroupId(context,(object)selectedGroup);
                if(groupId==null)return false;
                GestureMode? handle=HitHandle(new ShapeGeometry(groupId,box.Left,box.Top,box.Width,box.Height),map,mouse.Position,handleRadius);
                if(!handle.HasValue)return false;
                child=(object)selectedChild;mode=handle.Value;
            }
            else
            {
                dynamic hit=window.RangeFromPoint(mouse.Position.X,mouse.Position.Y);
                if(hit==null)
                {
                    if(!MultiSelectionPointerPolicy.CanMove(selectedBounds,point,handleRadius,6*dpi/96f,null))return null;
                    child=(object)selectedChild;
                }
                else if((int)hit.Child!=0)
                {
                    dynamic hitGroup=ShapeCapabilities.TopLevel((object)hit);
                    if((int)hitGroup.Id!=(int)selectedGroup.Id)return null;
                    var bounds=ScreenBounds(ReadNativePosition((object)hit),map);
                    bool editingText=(int)hit.Id==(int)selectedChild.Id && (int)selection.Type==3;
                    if(TextPointerPolicy.UseNativeTextInput((int)hit.Type,editingText,bounds,point,6*dpi/96f))
                    {LogRoute(mouse,messageKind,"native-child-text-interior",ReadSelection(window));return false;}
                    child=(object)hit;
                }
                else if((int)hit.Id==(int)selectedGroup.Id)
                {LogRoute(mouse,messageKind,"native-parent-in-child-edit",ReadSelection(window));return false;}
                else return null;
            }
            return BeginGroupChildren(context,map,mouse,messageKind,root,dpi,handleRadius,previousPane,new[]{child},mode,true);
        }

        private bool? TryBeginGroupChildrenMove(LayoutContext context,ScreenMap map,MouseHookData mouse,int messageKind,
            IntPtr root,uint dpi,double handleRadius,int previousPane,object rangeObject)
        {
            dynamic range=rangeObject;dynamic window=context.Window;
            object[] selected=Enumerable.Range(1,(int)range.Count).Select(index=>(object)range[index]).ToArray();
            var bounds=new System.Drawing.RectangleF((float)(map.OriginX+(double)range.Left*map.ScaleX),
                (float)(map.OriginY+(double)range.Top*map.ScaleY),(float)((double)range.Width*map.ScaleX),(float)((double)range.Height*map.ScaleY));
            var point=new System.Drawing.PointF(mouse.Position.X,mouse.Position.Y);
            if(MultiSelectionPointerPolicy.HitHandle(bounds,point,handleRadius))
            {LogRoute(mouse,messageKind,"native-multi-child-resize-unsupported",ReadSelection(window));return false;}
            dynamic hit=window.RangeFromPoint(mouse.Position.X,mouse.Position.Y);
            var selectedIds=new HashSet<int>(selected.Select(shape=>(int)((dynamic)shape).Id));
            bool? selectedHit=hit==null ? (bool?)null : selectedIds.Contains((int)hit.Id);
            bool? route=MultiSelectionPointerPolicy.RouteMove(bounds,point,handleRadius,6*dpi/96f,selectedHit);
            if(!route.HasValue)
            {
                if((int)hit.Child==0)
                {
                    var roots=new HashSet<int>(selected.Select(shape=>(int)((dynamic)ShapeCapabilities.TopLevel(shape)).Id));
                    if(roots.Contains((int)hit.Id))
                    {LogRoute(mouse,messageKind,"native-parent-in-child-edit",ReadSelection(window));return false;}
                    return null;
                }
                int hitRoot=(int)((dynamic)ShapeCapabilities.TopLevel((object)hit)).Id;
                if(!selected.Any(shape=>(int)((dynamic)ShapeCapabilities.TopLevel(shape)).Id==hitRoot))return null;
                var hitBounds=ScreenBounds(ReadNativePosition((object)hit),map);
                if(TextPointerPolicy.UseNativeTextInput((int)hit.Type,false,hitBounds,point,6*dpi/96f))
                {LogRoute(mouse,messageKind,"native-child-text-interior",ReadSelection(window));return false;}
                return BeginGroupChildren(context,map,mouse,messageKind,root,dpi,handleRadius,previousPane,new[]{(object)hit},GestureMode.Move,true);
            }
            if(!route.Value)
            {LogRoute(mouse,messageKind,"native-multi-child-selection-miss",ReadSelection(window));return false;}
            return BeginGroupChildren(context,map,mouse,messageKind,root,dpi,handleRadius,previousPane,selected,GestureMode.Move,false);
        }

        private bool BeginGroupChildren(LayoutContext context,ScreenMap map,MouseHookData mouse,int messageKind,
            IntPtr root,uint dpi,double handleRadius,int previousPane,object[] children,GestureMode mode,bool pendingSelection)
        {
            dynamic window=context.Window;
            GroupChildCapture capture;
            try{capture=CaptureGroupChildren(context,children,mode);}
            catch(GroupChildCaptureException error)
            {
                // This entry point is reached only after the hit policy chose
                // a real related child or the selected members' move outline.
                // A failed capture must not leak that DOWN into a native move.
                SafeLog("gesture-child-capture-rejected",new {message=error.Message,
                    reason=error.InnerException==null ? null : error.InnerException.Message});
                pendingLeftUp=true;
                return true;
            }
            if(capture==null)
            {LogRoute(mouse,messageKind,"native-unconstrained-child",ReadSelection(window));return false;}
            capture.PendingSelection=pendingSelection;
            ShapeGeometry driver=context.Geometry.Single(shape=>shape.Id==capture.Edit.Plan.GroupIds[0]);
            LogRoute(mouse,messageKind,"owned-child-"+mode,ReadSelection(window),handleRadius,null,context.Geometry,previousPane);
            SetFocus(mouse.Window);SetCapture(mouse.Window);
            if(GetCapture()!=mouse.Window)
            {
                LogRoute(mouse,messageKind,"child-capture-failed",ReadSelection(window));
                pendingLeftUp=true;
                return true;
            }
            StartGesture(context,driver,mode,map,mouse.Position,root,mouse.Window,dpi,handleRadius,false,child:capture);
            return true;
        }

        private static string FindGroupId(LayoutContext context,object groupObject)
        {
            int nativeId=(int)((dynamic)groupObject).Id;
            return context.Shapes.Where(pair=>(int)((dynamic)pair.Value).Id==nativeId).Select(pair=>pair.Key).FirstOrDefault();
        }

        private static GroupChildCapture CaptureGroupChildren(LayoutContext context,object[] children,GestureMode mode)
        {
            var edit=GroupChildNativeEdit.Capture(context,children,EditMode(mode),
                mode!=GestureMode.Move&&children.Length==1&&(int)((dynamic)children[0]).LockAspectRatio!=0);
            return edit==null ? null : new GroupChildCapture {Edit=edit,PendingSelection=true};
        }

        private void UpdateGroupChild(Gesture state,Point point,bool mouseDown,string pointerMessage)
        {
            var child=state.GroupChild;
            double dx=(point.X-state.StartMouse.X)/state.Map.ScaleX,dy=(point.Y-state.StartMouse.Y)/state.Map.ScaleY;
            IReadOnlyList<SelectionMovePosition> requested;
            try
            {
                requested=child.Edit.Plan.Target(dx,dy);
                child.Edit.ThrowIfProvablyImpossible(requested);
            }
            catch(ConstraintConflictException error)
            {
                SetRestrictions(state,error.Relations);
                if(state.LastRejection!=error.Message)SafeGestureLog(state,"constraint-rejected",new {gestureId=state.Id,driverId=state.Driver.Id,
                    mode=state.Mode.ToString(),groupChild=true,childNativeIds=child.Edit.ChildRoots.Keys,message=error.Message});
                state.LastRejection=error.Message;return;
            }
            if(GroupChildEditPlan.SameChildren(requested,child.Edit.LastFrame.Extras)){SetRestrictions(state,null);state.LastRejection=null;return;}
            if(child.PendingSelection){child.Edit.SelectExact();child.PendingSelection=false;}
            if(!state.UndoStarted){application.StartNewUndoEntry();state.UndoStarted=true;}
            SelectionMoveFrame actual;
            try
            {
                actual=child.Edit.Apply(requested);
            }
            catch(GroupChildEditRejectedException error)
            {
                SetRestrictions(state,(error.InnerException as ConstraintConflictException)?.Relations);
                string rejection=error.InnerException==null ? error.Message : error.InnerException.Message;
                if(state.LastRejection!=rejection)SafeGestureLog(state,error.InnerException is ConstraintConflictException ? "constraint-rejected" : "gesture-write-rejected",
                    new {gestureId=state.Id,driverId=state.Driver.Id,mode=state.Mode.ToString(),groupChild=true,childNativeIds=child.Edit.ChildRoots.Keys,
                        message=rejection,geometry=child.Edit.LastFrame.Geometry,childGeometry=child.Edit.LastFrame.Extras});
                state.LastRejection=rejection;
                return;
            }
            SetRestrictions(state,null);state.LastRejection=null;state.LastGeometry=actual.Geometry.ToList();state.Frames++;
            SafeGestureLog(state,"gesture-frame",new {gestureId=state.Id,driverId=state.Driver.Id,mode=state.Mode.ToString(),groupChild=true,
                childNativeIds=child.Edit.ChildRoots.Keys,childGeometry=actual.Extras,frame=state.Frames,mouseDown,input="owned-gesture-loop",pointerMessage,
                residualPt=LayoutSolver.MaxResidual(actual.Geometry,state.Constraints),cursor=new {x=point.X,y=point.Y},geometry=actual.Geometry});
        }

        private bool? TryBeginSelectionMove(LayoutContext context,ScreenMap map,MouseHookData mouse,int messageKind,
            IntPtr root,uint dpi,double handleRadius,int previousPane)
        {
            dynamic window=context.Window;
            if((int)window.Selection.Type!=2)return false;
            dynamic range=window.Selection.ShapeRange;
            var selected=Enumerable.Range(1,(int)range.Count).Select(index=>(object)range[index]).ToArray();
            if(selected.Any(ShapeCapabilities.IsLocked))return false;
            if(selected.Any(shape=>(int)((dynamic)shape).Child!=0))
            {LogRoute(mouse,messageKind,"native-child-selection",ReadSelection(window));return false;}
            var point=new System.Drawing.PointF(mouse.Position.X,mouse.Position.Y);
            var selectionBounds=new System.Drawing.RectangleF((float)(map.OriginX+(double)range.Left*map.ScaleX),
                (float)(map.OriginY+(double)range.Top*map.ScaleY),(float)((double)range.Width*map.ScaleX),(float)((double)range.Height*map.ScaleY));
            if(MultiSelectionPointerPolicy.HitHandle(selectionBounds,point,handleRadius) ||
                selected.Any(shape=>MultiSelectionPointerPolicy.HitHandle(ScreenBounds(ReadNativePosition(shape),map),point,handleRadius)))
            {LogRoute(mouse,messageKind,"native-selection-handle",ReadSelection(window));return false;}
            dynamic hit=window.RangeFromPoint(mouse.Position.X,mouse.Position.Y);
            if(hit!=null)hit=ShapeCapabilities.TopLevel((object)hit);
            var selectedIds=selected.Select(shape=>(int)((dynamic)shape).Id).ToArray();
            bool? selectedHit=hit==null ? (bool?)null : selectedIds.Contains((int)hit.Id);
            bool? route=MultiSelectionPointerPolicy.RouteMove(selectionBounds,point,handleRadius,6*dpi/96f,selectedHit);
            if(!route.HasValue)return null;
            if(!route.Value)
            {LogRoute(mouse,messageKind,"native-selection-miss",ReadSelection(window));return false;}
            var capture=CaptureSelectionMove(context,selected);
            if(capture==null)
            {LogRoute(mouse,messageKind,"native-unconstrained-selection",ReadSelection(window));return false;}
            ShapeGeometry driver=context.Geometry.Single(shape=>shape.Id==capture.Plan.DriverIds[0]);
            LogRoute(mouse,messageKind,"owned-selection-Move",ReadSelection(window),handleRadius,null,context.Geometry,previousPane);
            SetFocus(mouse.Window);SetCapture(mouse.Window);
            if(GetCapture()!=mouse.Window)
            {LogRoute(mouse,messageKind,"capture-failed",ReadSelection(window));return false;}
            StartGesture(context,driver,GestureMode.Move,map,mouse.Position,root,mouse.Window,dpi,handleRadius,false,capture);
            return true;
        }

        private static SelectionMoveCapture CaptureSelectionMove(LayoutContext context,object[] selected)
        {
            var byNative=context.Shapes.ToDictionary(pair=>(int)((dynamic)pair.Value).Id,pair=>pair.Key);
            var driverIds=new List<string>();
            var extraShapes=new Dictionary<int,object>();
            var positions=new List<SelectionMovePosition>();
            var ids=new List<int>();
            foreach(object item in selected)
            {
                dynamic shape=item;
                int nativeId=(int)shape.Id;
                if(ids.Contains(nativeId))throw new InvalidOperationException("The selection contains a repeated native shape.");
                ids.Add(nativeId);
                string constrainedId;
                if(byNative.TryGetValue(nativeId,out constrainedId))driverIds.Add(constrainedId);
                else {extraShapes.Add(nativeId,item);positions.Add(ReadNativePosition(item));}
            }
            if(driverIds.Count==0)return null;
            var plan=new SelectionMovePlan(context.Geometry,context.Constraints,driverIds,positions);
            return new SelectionMoveCapture {Shapes=selected,NativeIds=ids.ToArray(),ExtraShapes=extraShapes,Plan=plan,LastFrame=plan.Initial};
        }

        private static SelectionMovePosition ReadNativePosition(object shapeObject)
        {
            dynamic shape=shapeObject;
            return new SelectionMovePosition((int)shape.Id,(double)shape.Left,(double)shape.Top,(double)shape.Width,(double)shape.Height);
        }

        private static void SelectNativeShapes(object[] shapes)
        {
            for(int index=0;index<shapes.Length;index++)((dynamic)shapes[index]).Select(index==0 ? -1 : 0);
        }

        private static SelectionMoveFrame ReadSelectionFrame(Gesture state)
        {
            return new SelectionMoveFrame(PowerPointLayout.Snapshot(state.Context),
                state.SelectionMove.ExtraShapes.Values.Select(ReadNativePosition));
        }

        private static void WriteSelectionFrame(Gesture state,SelectionMoveFrame frame)
        {
            PowerPointLayout.Apply(state.Context,frame.Geometry);
            foreach(var position in frame.Extras)
            {
                dynamic shape=state.SelectionMove.ExtraShapes[position.NativeId];
                if(Math.Abs((double)shape.Left-position.Left)>0.0001)shape.Left=(float)position.Left;
                if(Math.Abs((double)shape.Top-position.Top)>0.0001)shape.Top=(float)position.Top;
            }
        }

        private void UpdateSelectionMove(Gesture state,Point point,bool mouseDown,string pointerMessage)
        {
            var selection=state.SelectionMove;
            double dx=(point.X-state.StartMouse.X)/state.Map.ScaleX,dy=(point.Y-state.StartMouse.Y)/state.Map.ScaleY;
            SelectionMoveFrame requested;
            try {requested=selection.Plan.Solve(dx,dy);}
            catch(ConstraintConflictException error)
            {
                SetRestrictions(state,LayoutDiagnostics.ExplainEdit(state.Before,state.Constraints,selection.Plan.RequestedEdits(dx,dy),state.Restrictions));
                if(state.LastRejection!=error.Message)SafeGestureLog(state,"constraint-rejected",new {
                    gestureId=state.Id,driverId=state.Driver.Id,mode=state.Mode.ToString(),selectionMove=true,message=error.Message,dx,dy,selectedDriverIds=selection.Plan.DriverIds});
                state.LastRejection=error.Message;
                return;
            }
            SetRestrictions(state,null);state.LastRejection=null;
            if(selection.LastFrame.Matches(requested))return;
            if(!state.UndoStarted){application.StartNewUndoEntry();state.UndoStarted=true;}
            SelectionMoveFrame actual;
            try {actual=selection.Plan.ApplyAtomic(selection.LastFrame,requested,
                frame=>WriteSelectionFrame(state,frame),()=>ReadSelectionFrame(state));}
            catch(SelectionMoveRejectedException error)
            {
                state.LastRejection=error.InnerException==null ? error.Message : error.InnerException.Message;
                SafeGestureLog(state,"gesture-write-rejected",new {gestureId=state.Id,mode=state.Mode.ToString(),selectionMove=true,message=state.LastRejection,
                    geometry=selection.LastFrame.Geometry,extras=selection.LastFrame.Extras});
                return;
            }
            selection.LastFrame=actual;
            state.LastGeometry=actual.Geometry.ToList();state.Frames++;
            SafeGestureLog(state,"gesture-frame",new {gestureId=state.Id,driverId=state.Driver.Id,mode=state.Mode.ToString(),selectionMove=true,
                frame=state.Frames,mouseDown,input="owned-gesture-loop",pointerMessage,
                residualPt=LayoutSolver.MaxResidual(actual.Geometry,state.Constraints),cursor=new {x=point.X,y=point.Y},
                geometry=actual.Geometry,extras=actual.Extras,selectedNativeIds=selection.NativeIds});
        }

        private void StartGesture(LayoutContext context, ShapeGeometry driver, GestureMode mode, ScreenMap map,
            Point start, IntPtr root, IntPtr captureWindow, uint dpi, double handleRadius, bool test,
            SelectionMoveCapture selection=null,GroupChildCapture child=null)
        {
            queuedMotion = null;
            queuedRelease = null;
            gesture = new Gesture
            {
                Id = Guid.NewGuid().ToString("D"), Context = context, Driver = driver,
                Mode = mode, StartMouse = start, LastMouse = start,
                Map = map, Before = context.Geometry.ToList(), LastGeometry = context.Geometry.ToList(),
                RootWindow = root, CaptureWindow = captureWindow, IsTest = test,
                DragThreshold = 4.0 * dpi / 96.0, DragThresholdPassed = mode != GestureMode.Move,
                SelectionMove=selection,GroupChild=child,
                Constraints=LayoutSolver.ConnectedConstraints(context.Constraints,
                    selection!=null ? (IEnumerable<string>)selection.Plan.DriverIds : child!=null ? (IEnumerable<string>)child.Edit.Plan.GroupIds : new[]{driver.Id})
            };
            pendingLeftUp = !test;
            SafeGestureLog(gesture, "gesture-start", new
            {
                gestureId = gesture.Id, driverId = driver.Id, mode = mode.ToString(), slideId = context.SlideId,
                cursor = new { x = start.X, y = start.Y },
                scale = new { x = map.ScaleX, y = map.ScaleY }, dpi, handleRadius, before = gesture.Before,
                dragThreshold = gesture.DragThreshold, test,
                selectedNativeIds=selection==null ? null : selection.NativeIds,
                selectedDriverIds=selection==null ? null : selection.Plan.DriverIds,
                extras=selection==null ? null : selection.LastFrame.Extras,
                childNativeIds=child==null ? null : child.Edit.ChildRoots.Keys.ToArray(),
                childBefore=child==null ? null : child.Edit.LastFrame.Extras,
                constraints = ConstraintData(gesture.Constraints)
            });
        }

        private void Update(Point point, bool mouseDown, string pointerMessage)
        {
            if (gesture == null || updating || finishing) return;
            Gesture state = gesture;
            if (point.X == state.LastMouse.X && point.Y == state.LastMouse.Y) return;
            state.LastMouse = point;
            if (!state.DragThresholdPassed)
            {
                if (!GroupPointerPolicy.HasDragDistance(point.X-state.StartMouse.X,point.Y-state.StartMouse.Y,state.DragThreshold)) return;
                state.DragThresholdPassed = true;
            }
            updating = true;
            try
            {
                if(state.GroupChild!=null)
                {
                    UpdateGroupChild(state,point,mouseDown,pointerMessage);
                    return;
                }
                if(state.SelectionMove!=null)
                {
                    UpdateSelectionMove(state,point,mouseDown,pointerMessage);
                    return;
                }
                if(state.DeferSelection)
                {
                    ((dynamic)state.Context.Shapes[state.Driver.Id]).Select();
                    state.DeferSelection=false;
                }
                double dx = (point.X - state.StartMouse.X) / state.Map.ScaleX;
                double dy = (point.Y - state.StartMouse.Y) / state.Map.ScaleY;
                ShapeGeometry target = ComputeTarget(state.Driver, state.Mode, dx, dy);
                var edits = new[] { new GeometryEdit(target, EditMode(state.Mode)) };
                IReadOnlyList<ShapeGeometry> result;
                try { result = LayoutSolver.Solve(state.Before, state.Constraints, edits); }
                catch (ConstraintConflictException exception)
                {
                    SetRestrictions(state,LayoutDiagnostics.ExplainEdit(state.Before,state.Constraints,edits,state.Restrictions));
                    if (state.LastRejection != exception.Message)
                        SafeGestureLog(state, "constraint-rejected", new
                        {
                            gestureId = state.Id, driverId = state.Driver.Id, mode = state.Mode.ToString(),
                            message = exception.Message, requested = target
                        });
                    state.LastRejection = exception.Message;
                    return;
                }
                SetRestrictions(state,LayoutSolver.NeedsUpdate(target,result.Single(shape=>shape.Id==target.Id))
                    ? LayoutDiagnostics.ExplainEdit(state.Before,state.Constraints,edits,state.Restrictions) : null);
                state.LastRejection = null;
                if (SameGeometry(state.LastGeometry, result)) return;
                if (!state.UndoStarted)
                {
                    application.StartNewUndoEntry();
                    state.UndoStarted = true;
                }
                PowerPointLayout.Apply(state.Context, result);
                List<ShapeGeometry> actual = PowerPointLayout.Snapshot(state.Context);
                if (!SameGeometry(actual, result))
                    throw new InvalidOperationException("PowerPoint geometry readback differs from the solved layout.");
                double residual = LayoutSolver.MaxResidual(actual, state.Constraints);
                if (residual > LayoutSolver.PositionTolerance)
                    throw new InvalidOperationException("PowerPoint geometry violates a constraint after writeback.");
                state.LastGeometry = actual;
                state.Frames++;
                SafeGestureLog(state, "gesture-frame", new
                {
                    gestureId = state.Id, driverId = state.Driver.Id, mode = state.Mode.ToString(),
                    frame = state.Frames, mouseDown, input = "owned-gesture-loop", pointerMessage, residualPt = residual,
                    cursor = new { x = point.X, y = point.Y }, geometry = actual
                });
            }
            finally { updating = false; }
        }

        private void SetRestrictions(Gesture state,IReadOnlyList<LayoutConstraint> relations)
        {
            var current=relations??Array.Empty<LayoutConstraint>();
            if(state.Restrictions.SequenceEqual(current))return;
            state.Restrictions=current;
            SafeGestureLog(state,"gesture-restrictions",new {gestureId=state.Id,driverId=state.Driver.Id,
                slideId=state.Context.SlideId,relations=ConstraintData(current)});
        }

        private void CheckRelease(Point point)
        {
            Gesture state = gesture;
            if (state == null) return;
            if((state.SelectionMove!=null || state.GroupChild!=null) && state.LastRejection!=null)
            {
                SafeGestureLog(state,"gesture-release-rejected",new {gestureId=state.Id,driverId=state.Driver.Id,
                    mode=state.Mode.ToString(),reason=state.LastRejection,frames=state.Frames,geometry=state.LastGeometry,
                    extras=state.SelectionMove!=null ? state.SelectionMove.LastFrame.Extras : state.GroupChild.Edit.LastFrame.Extras});
                return;
            }
            // The originating DOWN loop has flushed the exact release point.
            // Verify it from the actual readback, including clamps and axes that
            // are irrelevant to an edge handle, before closing the gesture.
            bool geometryMatched = false;
            string error = null;
            try
            {
                // A click is selection-only, even if the document was already
                // out of constraint before DOWN. It must not solve or add Undo.
                if(state.GroupChild!=null)
                {
                    double dx=(point.X-state.StartMouse.X)/state.Map.ScaleX;
                    double dy=(point.Y-state.StartMouse.Y)/state.Map.ScaleY;
                    geometryMatched=GroupChildEditPlan.SameChildren(state.GroupChild.Edit.LastFrame.Extras,
                        state.DragThresholdPassed ? state.GroupChild.Edit.Plan.Target(dx,dy) : state.GroupChild.Edit.Plan.Initial.Extras);
                }
                else if(state.SelectionMove!=null)
                {
                    double dx=(point.X-state.StartMouse.X)/state.Map.ScaleX;
                    double dy=(point.Y-state.StartMouse.Y)/state.Map.ScaleY;
                    var expected=state.DragThresholdPassed ? state.SelectionMove.Plan.Solve(dx,dy) : state.SelectionMove.Plan.Initial;
                    geometryMatched=state.SelectionMove.LastFrame.Matches(expected);
                }
                else if(!state.DragThresholdPassed)
                {
                    geometryMatched=SameGeometry(state.LastGeometry,state.Before);
                }
                else
                {
                    double dx = (point.X - state.StartMouse.X) / state.Map.ScaleX;
                    double dy = (point.Y - state.StartMouse.Y) / state.Map.ScaleY;
                    ShapeGeometry target = ComputeTarget(state.Driver, state.Mode, dx, dy);
                    var expected = LayoutSolver.Solve(state.Before, state.Constraints, state.Driver.Id, target, EditMode(state.Mode));
                    geometryMatched = SameGeometry(state.LastGeometry, expected);
                }
            }
            catch (ConstraintConflictException exception)
            {
                SafeGestureLog(state,"gesture-release-rejected",new {gestureId=state.Id,driverId=state.Driver.Id,
                    mode=state.Mode.ToString(),reason=exception.Message,frames=state.Frames,geometry=state.LastGeometry});
                return;
            }
            catch (Exception exception) { error = exception.GetType().Name + ": " + exception.Message; }
            SafeGestureLog(state, geometryMatched ? "gesture-release" : "gesture-release-mismatch-error", new
            {
                gestureId = state.Id, driverId = state.Driver.Id, mode = state.Mode.ToString(), geometryMatched,
                lastMotion = new { x = state.LastMouse.X, y = state.LastMouse.Y },
                releasePoint = new { x = point.X, y = point.Y }, frames = state.Frames, error
            });
        }

        private void Finish(bool cancel, string reason)
        {
            Gesture state = gesture;
            if (state == null || finishing) return;
            gesture = null;
            finishing = true;
            try
            {
                if (cancel && state.UndoStarted)
                {
                    if(state.GroupChild!=null)state.GroupChild.Edit.Restore(state.GroupChild.Edit.Plan.Initial);
                    else if(state.SelectionMove==null)PowerPointLayout.Apply(state.Context,state.Before);
                    else WriteSelectionFrame(state,state.SelectionMove.Plan.Initial);
                }
                List<ShapeGeometry> after = PowerPointLayout.Snapshot(state.Context);
                if (cancel && !SameGeometry(after, state.Before))
                    throw new InvalidOperationException("Cancelled gesture did not restore its initial layout.");
                if(cancel && state.SelectionMove!=null && !ReadSelectionFrame(state).Matches(state.SelectionMove.Plan.Initial))
                    throw new InvalidOperationException("Cancelled selection move did not restore every selected native object.");
                if(cancel && state.GroupChild!=null && !state.GroupChild.Edit.ReadFrame().Matches(state.GroupChild.Edit.Plan.Initial))
                    throw new InvalidOperationException("Cancelled group-child edit did not restore the actual member and group bounds.");
                // COM size/position updates can leave native selection adorners
                // based on stale bounds. Recreate the selection only after the
                // pointer is released, so the next handle drag uses current bounds.
                if (state.CaptureWindow != IntPtr.Zero && GetCapture() == state.CaptureWindow) ReleaseCapture();
                object selectionBefore = null, selectionAfter = null;
                bool childSelected=false;
                if ((reason == "mouse-up" || reason == "button-released" || reason == "escape") &&
                    (state.IsTest || GetAncestor(GetForegroundWindow(), GaRoot) == state.RootWindow))
                {
                    dynamic window = state.Context.Window;
                    // Do not commit a pending click into another slide/window.
                    if(SameComObject((object)application.ActiveWindow,state.Context.Window) &&
                        (int)window.View.Slide.SlideID==state.Context.SlideId)
                    {
                        selectionBefore = ReadSelection(window);
                        bool normalRelease=reason=="mouse-up" || reason=="button-released";
                        bool selectChild=GroupPointerPolicy.ShouldSelectChild(state.WasSelectedGroup,state.ClickChild!=null,
                            state.Mode==GestureMode.Move,state.DragThresholdPassed,cancel,normalRelease);
                        if(state.GroupChild!=null)
                        {
                            if(!cancel || state.DragThresholdPassed)
                            {
                                if(state.DragThresholdPassed)window.Selection.Unselect();
                                state.GroupChild.Edit.SelectExact();
                            }
                        }
                        else if(state.SelectionMove!=null)
                        {
                            if(!cancel || state.DragThresholdPassed)
                            {
                                if(state.DragThresholdPassed)window.Selection.Unselect();
                                SelectNativeShapes(state.SelectionMove.Shapes);
                            }
                        }
                        else if(selectChild && (int)((dynamic)ShapeCapabilities.TopLevel(state.ClickChild)).Id==
                            (int)((dynamic)state.Context.Shapes[state.Driver.Id]).Id)
                        {
                            ((dynamic)state.ClickChild).Select();
                            childSelected=true;
                        }
                        else if(!cancel || state.DragThresholdPassed)
                        {
                            if(state.DragThresholdPassed)window.Selection.Unselect();
                            ((dynamic)state.Context.Shapes[state.Driver.Id]).Select();
                        }
                        selectionAfter = ReadSelection(window);
                    }
                }
                if (state.UndoStarted) application.StartNewUndoEntry();
                SafeGestureLog(state, "gesture-end", new
                {
                    gestureId = state.Id, driverId = state.Driver.Id, mode = state.Mode.ToString(),
                    reason, cancelled = cancel, frames = state.Frames, before = state.Before, after,
                    selectionBefore, selectionAfter, childSelected, dragThresholdPassed = state.DragThresholdPassed, test = state.IsTest,
                    selectedNativeIds=state.SelectionMove==null ? null : state.SelectionMove.NativeIds,
                    beforeExtras=state.SelectionMove==null ? null : state.SelectionMove.Plan.Initial.Extras,
                    afterExtras=state.SelectionMove==null ? null : ReadSelectionFrame(state).Extras,
                    childNativeIds=state.GroupChild==null ? null : state.GroupChild.Edit.ChildRoots.Keys.ToArray(),
                    childBefore=state.GroupChild==null ? null : state.GroupChild.Edit.Plan.Initial.Extras,
                    childAfter=state.GroupChild==null ? null : state.GroupChild.Edit.ReadChildren()
                });
            }
            catch (Exception exception)
            {
                SafeGestureLog(state, "gesture-end-error", new
                {
                    gestureId = state.Id, reason, error = exception.GetType().Name, message = exception.Message
                });
                cancel = true;
                try { if (state.UndoStarted) application.StartNewUndoEntry(); } catch { }
            }
            finally
            {
                if (state.CaptureWindow != IntPtr.Zero && GetCapture() == state.CaptureWindow) ReleaseCapture();
                finishing = false;
                if (!state.IsTest) CanvasGestureFinished?.Invoke(cancel);
            }
        }

        private static ShapeGeometry ComputeTarget(ShapeGeometry start, GestureMode mode, double dx, double dy)
        {
            if (mode == GestureMode.Move)
                return new ShapeGeometry(start.Id, start.Left + dx, start.Top + dy, start.Width, start.Height);
            bool west = mode == GestureMode.ResizeW || mode == GestureMode.ResizeNW || mode == GestureMode.ResizeSW;
            bool east = mode == GestureMode.ResizeE || mode == GestureMode.ResizeNE || mode == GestureMode.ResizeSE;
            bool north = mode == GestureMode.ResizeN || mode == GestureMode.ResizeNW || mode == GestureMode.ResizeNE;
            bool south = mode == GestureMode.ResizeS || mode == GestureMode.ResizeSW || mode == GestureMode.ResizeSE;
            double width = west ? Math.Max(LayoutSolver.MinimumSize, start.Width - dx) :
                east ? Math.Max(LayoutSolver.MinimumSize, start.Width + dx) : start.Width;
            double height = north ? Math.Max(LayoutSolver.MinimumSize, start.Height - dy) :
                south ? Math.Max(LayoutSolver.MinimumSize, start.Height + dy) : start.Height;
            return new ShapeGeometry(start.Id, west ? start.Left + start.Width - width : start.Left,
                north ? start.Top + start.Height - height : start.Top, width, height);
        }

        private static GeometryEditMode EditMode(GestureMode mode)
        { return (GeometryEditMode)Enum.Parse(typeof(GeometryEditMode),mode.ToString()); }

        private static bool SameGeometry(IEnumerable<ShapeGeometry> first, IEnumerable<ShapeGeometry> second)
        {
            var expected = second.ToDictionary(item => item.Id, StringComparer.Ordinal);
            int count = 0;
            foreach (var item in first)
            {
                ShapeGeometry other;
                if (!expected.TryGetValue(item.Id, out other) || LayoutSolver.NeedsUpdate(item, other)) return false;
                count++;
            }
            return count == expected.Count;
        }

        private static int ActivateSlidePane(dynamic window)
        {
            int previous = (int)window.ActivePane.ViewType;
            if (previous != 1)
            {
                for (int i=1; i <= (int)window.Panes.Count; i++)
                    if ((int)window.Panes[i].ViewType == 1) { window.Panes[i].Activate(); break; }
                if ((int)window.ActivePane.ViewType != 1)
                    throw new InvalidOperationException("PowerPoint could not activate the slide editing pane.");
            }
            return previous;
        }

        private static bool SameComObject(object first,object second)
        {
            IntPtr firstIdentity=Marshal.GetIUnknownForObject(first);
            try
            {
                IntPtr secondIdentity=Marshal.GetIUnknownForObject(second);
                try{return firstIdentity==secondIdentity;}
                finally{Marshal.Release(secondIdentity);}
            }
            finally{Marshal.Release(firstIdentity);}
        }

        private static Point HandlePoint(Handle handle)
        { return new Point { X=(int)Math.Round(handle.X), Y=(int)Math.Round(handle.Y) }; }

        private static object[] ConstraintData(IEnumerable<LayoutConstraint> constraints)
        {
            return constraints.Select(c => (object)new { c.Id, c.SlideId, kind = c.Kind.ToString(), c.ShapeIds }).ToArray();
        }

        private static int SelectedShapeId(dynamic window)
        {
            dynamic selection = window.Selection;
            int type = (int)selection.Type;
            if (type != 2 && type != 3) return -1;
            try { return (int)selection.ShapeRange.Count == 1 ? (int)selection.ShapeRange[1].Id : -1; }
            catch (COMException) { return -1; }
        }

        private static int SelectedShapeCount(dynamic window)
        {
            dynamic selection=window.Selection;
            int type=(int)selection.Type;
            if(type!=2 && type!=3)return 0;
            try{return (int)selection.ShapeRange.Count;}catch(COMException){return 0;}
        }

        private static object ReadSelection(dynamic window)
        {
            try
            {
                dynamic selection = window.Selection;
                int type = (int)selection.Type;
                if (type != 2 && type != 3) return new { type };
                dynamic range = selection.ShapeRange;
                bool hasChild=(bool)selection.HasChildShapeRange;
                return new { type, count = (int)range.Count, nativeId = (int)range.Count == 1 ? (int)range[1].Id : -1,
                    nativeIds=Enumerable.Range(1,(int)range.Count).Select(index=>(int)range[index].Id).ToArray(),
                    hasChild, childNativeId=hasChild && (int)selection.ChildShapeRange.Count==1 ? (int)selection.ChildShapeRange[1].Id : -1,
                    left = (double)range.Left, top = (double)range.Top, width = (double)range.Width, height = (double)range.Height };
            }
            catch (Exception exception) { return new { error = exception.GetType().Name, message = exception.Message }; }
        }

        private static bool CanDrag(object shapeObject)
        {
            dynamic shape = shapeObject;
            return ShapeCapabilities.IsSupported(shapeObject) && (int)shape.Visible != 0 && !ShapeCapabilities.IsLocked(shapeObject);
        }

        private static bool HitBody(ShapeGeometry shape, ScreenMap map, Point point)
        {
            double left = map.OriginX + shape.Left * map.ScaleX, top = map.OriginY + shape.Top * map.ScaleY;
            return point.X >= left && point.X <= left + shape.Width * map.ScaleX &&
                point.Y >= top && point.Y <= top + shape.Height * map.ScaleY;
        }

        private static System.Drawing.RectangleF ScreenBounds(ShapeGeometry shape, ScreenMap map)
        {
            return new System.Drawing.RectangleF((float)(map.OriginX + shape.Left * map.ScaleX),
                (float)(map.OriginY + shape.Top * map.ScaleY), (float)(shape.Width * map.ScaleX), (float)(shape.Height * map.ScaleY));
        }

        private static System.Drawing.RectangleF ScreenBounds(SelectionMovePosition shape,ScreenMap map)
        {
            return new System.Drawing.RectangleF((float)(map.OriginX+shape.Left*map.ScaleX),
                (float)(map.OriginY+shape.Top*map.ScaleY),(float)(shape.Width*map.ScaleX),(float)(shape.Height*map.ScaleY));
        }

        private static GestureMode? HitHandle(ShapeGeometry shape, ScreenMap map, Point point, double radius)
        {
            Handle nearest = null;
            double distance = radius * radius;
            foreach (var handle in Handles(shape, map))
            {
                double dx = point.X - handle.X, dy = point.Y - handle.Y, squared = dx * dx + dy * dy;
                if (squared <= distance) { distance = squared; nearest = handle; }
            }
            return nearest == null ? (GestureMode?)null : nearest.Mode;
        }

        private static object ClosestHandle(ShapeGeometry shape, ScreenMap map, Point point)
        {
            var nearest = Handles(shape, map).OrderBy(h => (point.X-h.X)*(point.X-h.X)+(point.Y-h.Y)*(point.Y-h.Y)).First();
            double dx = point.X - nearest.X, dy = point.Y - nearest.Y;
            return new { id = shape.Id, mode = nearest.Mode.ToString(), x = nearest.X, y = nearest.Y, distance = Math.Sqrt(dx*dx+dy*dy) };
        }

        private static Handle[] Handles(ShapeGeometry shape, ScreenMap map)
        {
            double left = map.OriginX + shape.Left * map.ScaleX, top = map.OriginY + shape.Top * map.ScaleY;
            double right = left + shape.Width * map.ScaleX, bottom = top + shape.Height * map.ScaleY;
            double cx = (left + right) / 2, cy = (top + bottom) / 2;
            return new[] {
                new Handle(GestureMode.ResizeNW,left,top), new Handle(GestureMode.ResizeN,cx,top),
                new Handle(GestureMode.ResizeNE,right,top), new Handle(GestureMode.ResizeE,right,cy),
                new Handle(GestureMode.ResizeSE,right,bottom), new Handle(GestureMode.ResizeS,cx,bottom),
                new Handle(GestureMode.ResizeSW,left,bottom), new Handle(GestureMode.ResizeW,left,cy)
            };
        }

        private static string WindowClass(IntPtr window)
        {
            var name = new StringBuilder(256);
            GetClassName(window, name, name.Capacity);
            return name.ToString();
        }

        internal static bool IsCanvasWindow(IntPtr window)
        {
            // The tested slide canvas is the leaf mdiClass. Testing ancestors
            // also admits thumbnail/notes panes, which cannot map slide points.
            return String.Equals(WindowClass(window), "mdiClass", StringComparison.Ordinal);
        }

        private void LogRoute(MouseHookData mouse, int messageKind, string route, object selection = null,
            double? radius = null, object closest = null, IEnumerable<ShapeGeometry> geometry = null, int? previousPane = null)
        {
            var parents = new List<string>();
            for (IntPtr parent = mouse.Window; parent != IntPtr.Zero && parents.Count < 5; parent = GetParent(parent))
                parents.Add(WindowClass(parent));
            SafeLog("pointer-route", new { message = messageKind == WmLeftDoubleClick ? "double-click" : "down", route,
                cursor = new { x = mouse.Position.X, y = mouse.Position.Y }, window = mouse.Window.ToInt64(), classes = parents,
                capture = GetCapture().ToInt64(), ctrl = KeyDown(0x11), alt = KeyDown(0x12), shift = KeyDown(0x10),
                selection, radius, closest, geometry, previousPane });
        }

        private static bool KeyDown(int key) { return (GetAsyncKeyState(key) & 0x8000) != 0; }

        private void SafeLog(string type, object detail)
        {
            try { log(type, serializer.Serialize(detail)); } catch { /* Never unwind through a native hook. */ }
        }

        private void SafeGestureLog(Gesture state, string type, object detail)
        { SafeLog(state.IsTest ? "test-"+type : type, detail); }

        private void RequireUiThread()
        {
            if (GetCurrentThreadId() != uiThreadId)
                throw new InvalidOperationException("The gesture controller must run on PowerPoint's UI thread.");
        }

        public void Dispose()
        {
            RequireUiThread();
            if (disposed) return;
            disposed = true;
            Finish(true, "dispose");
            if (hook != IntPtr.Zero) { UnhookWindowsHookEx(hook); hook = IntPtr.Zero; }
            if (keyboardHook != IntPtr.Zero) { UnhookWindowsHookEx(keyboardHook); keyboardHook = IntPtr.Zero; }
            pendingLeftUp = false;
            SafeLog("gesture-hook-removed", new { thread = uiThreadId });
        }

        private enum GestureMode { Move, ResizeNW, ResizeN, ResizeNE, ResizeE, ResizeSE, ResizeS, ResizeSW, ResizeW }
        private sealed class Handle
        {
            public GestureMode Mode; public double X, Y;
            public Handle(GestureMode mode, double x, double y) { Mode = mode; X = x; Y = y; }
        }
        private sealed class Gesture
        {
            public string Id;
            public LayoutContext Context;
            public IReadOnlyList<LayoutConstraint> Constraints;
            public ShapeGeometry Driver;
            public GestureMode Mode;
            public Point StartMouse, LastMouse;
            public ScreenMap Map;
            public List<ShapeGeometry> Before, LastGeometry;
            public IntPtr RootWindow, CaptureWindow;
            public bool UndoStarted;
            public bool IsTest, DragThresholdPassed;
            public bool DeferSelection, WasSelectedGroup;
            public object ClickChild;
            public SelectionMoveCapture SelectionMove;
            public GroupChildCapture GroupChild;
            public bool CancelRequested;
            public DateTime? PhysicalReleaseSince;
            public double DragThreshold;
            public int Frames;
            public string LastRejection;
            public IReadOnlyList<LayoutConstraint> Restrictions=Array.Empty<LayoutConstraint>();
        }
        private sealed class SelectionMoveCapture
        {
            public object[] Shapes;
            public int[] NativeIds;
            public Dictionary<int,object> ExtraShapes;
            public SelectionMovePlan Plan;
            public SelectionMoveFrame LastFrame;
        }
        private sealed class GroupChildCapture
        {
            public GroupChildNativeEdit Edit;
            public bool PendingSelection;
        }
        private sealed class ScreenMap
        {
            public double OriginX, OriginY, ScaleX, ScaleY;
            public static ScreenMap From(object windowObject)
            {
                dynamic window = windowObject;
                var map = new ScreenMap();
                map.OriginX = (double)window.PointsToScreenPixelsX(0f);
                map.OriginY = (double)window.PointsToScreenPixelsY(0f);
                map.ScaleX = ((double)window.PointsToScreenPixelsX(100f) - map.OriginX) / 100.0;
                map.ScaleY = ((double)window.PointsToScreenPixelsY(100f) - map.OriginY) / 100.0;
                if (Double.IsNaN(map.ScaleX) || Double.IsInfinity(map.ScaleX) || map.ScaleX <= 0 ||
                    Double.IsNaN(map.ScaleY) || Double.IsInfinity(map.ScaleY) || map.ScaleY <= 0)
                    throw new InvalidOperationException("PowerPoint returned an invalid slide coordinate transform.");
                return map;
            }
            public Point Point(double x, double y)
            {
                return new Point { X = (int)Math.Round(OriginX + x * ScaleX), Y = (int)Math.Round(OriginY + y * ScaleY) };
            }
        }
        [StructLayout(LayoutKind.Sequential)] private struct Point { public int X, Y; }
        [StructLayout(LayoutKind.Sequential)] private struct MouseHookData
        {
            public Point Position;
            public IntPtr Window;
            public uint HitTestCode;
            public UIntPtr ExtraInfo;
        }
        [UnmanagedFunctionPointer(CallingConvention.Winapi)] private delegate IntPtr MouseHook(int code, IntPtr message, IntPtr data);
        [UnmanagedFunctionPointer(CallingConvention.Winapi)] private delegate IntPtr KeyboardHook(int code, IntPtr key, IntPtr flags);
        [DllImport("user32.dll", SetLastError=true)] private static extern IntPtr SetWindowsHookEx(int hookType, MouseHook callback, IntPtr module, uint threadId);
        [DllImport("user32.dll", EntryPoint="SetWindowsHookExW", SetLastError=true)] private static extern IntPtr SetKeyboardHookEx(int hookType, KeyboardHook callback, IntPtr module, uint threadId);
        [DllImport("user32.dll")] private static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")] private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr message, IntPtr data);
        [DllImport("user32.dll")] private static extern IntPtr SetCapture(IntPtr window);
        [DllImport("user32.dll")] private static extern bool ReleaseCapture();
        [DllImport("user32.dll")] private static extern IntPtr GetCapture();
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] private static extern IntPtr SetFocus(IntPtr window);
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
        [DllImport("user32.dll")] private static extern bool GetCursorPos(out Point point);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int maxCount);
        [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(Point point);
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr window);
        [DllImport("user32.dll", SetLastError=true)] private static extern uint MsgWaitForMultipleObjectsEx(uint count, IntPtr handles, uint timeout, uint wakeMask, uint flags);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
        [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();
    }
}
