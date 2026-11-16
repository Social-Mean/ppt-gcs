using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    public sealed class DragPair
    {
        public object A { get; private set; }
        public object B { get; private set; }
        public SameTopConstraint Constraint { get; private set; }

        public DragPair(object a, object b, SameTopConstraint constraint)
        {
            if (a == null) throw new ArgumentNullException("a");
            if (b == null) throw new ArgumentNullException("b");
            if (constraint == null) throw new ArgumentNullException("constraint");
            A = a;
            B = b;
            Constraint = constraint;
        }
    }

    // Runs exclusively on PowerPoint's UI thread. It owns the mouse gesture so
    // native drag previews do not conceal intermediate geometry from the COM model.
    public sealed class ConstrainedDragController : IDisposable
    {
        private const int WhMouse = 7;
        private const int HcAction = 0;
        private const int WmMouseMove = 0x0200;
        private const int WmLeftDown = 0x0201;
        private const int WmLeftUp = 0x0202;
        private const int WmLeftDoubleClick = 0x0203;
        private const uint GaRoot = 2;
        private const double HandleMargin = 8.0;

        private readonly dynamic application;
        private readonly Action<string, string> log;
        private readonly uint uiThreadId;
        private readonly MouseHook hookDelegate;
        private readonly List<DragPair> targets = new List<DragPair>();
        private IntPtr hook;
        private DragState drag;
        private bool disposed;
        private bool inCallback;
        private bool updating;

        public bool IsDragging { get { return drag != null; } }

        public ConstrainedDragController(object application, Action<string, string> log)
        {
            if (application == null) throw new ArgumentNullException("application");
            if (log == null) throw new ArgumentNullException("log");
            this.application = application;
            this.log = log;
            uiThreadId = GetCurrentThreadId();
            hookDelegate = OnMouse;
            hook = SetWindowsHookEx(WhMouse, hookDelegate, IntPtr.Zero, uiThreadId);
            if (hook == IntPtr.Zero)
                throw new InvalidOperationException("Cannot install PowerPoint mouse hook: " + Marshal.GetLastWin32Error());
            SafeLog("drag-hook-installed", "thread=" + uiThreadId);
        }

        public void SetTargets(IEnumerable<DragPair> pairs)
        {
            RequireUiThread();
            if (disposed) return;
            if (pairs == null) throw new ArgumentNullException("pairs");
            // The gesture retains the COM references and start geometry it began with.
            if (drag != null) return;
            targets.Clear();
            foreach (var pair in pairs)
                if (pair != null) targets.Add(pair);
        }

        // Read-only host diagnostics. This exercises the real Office mapping and
        // Win32 hit target without synthesizing input or changing the document.
        public string InspectTargets()
        {
            RequireUiThread();
            var result = new StringBuilder();
            IntPtr foreground = GetForegroundWindow();
            IntPtr root = GetAncestor(foreground, GaRoot);
            uint foregroundProcess;
            uint foregroundThread = GetWindowThreadProcessId(root, out foregroundProcess);
            result.AppendFormat(CultureInfo.InvariantCulture,
                "ownerProcess={0} ownerThread={1} currentThread={2} foregroundRoot={3} foregroundClass={4} foregroundProcess={5} foregroundThread={6}\r\n",
                GetCurrentProcessId(), uiThreadId, GetCurrentThreadId(), root, WindowClass(root), foregroundProcess, foregroundThread);
            if ((int)application.Presentations.Count == 0) return result.Append("No presentation.").ToString();
            dynamic window = application.ActiveWindow;
            result.AppendFormat(CultureInfo.InvariantCulture, "view={0} presentation={1}\r\n", (int)window.ViewType, (string)window.Presentation.FullName);
            if ((int)window.ViewType != 9) return result.ToString();
            dynamic slide = window.View.Slide;
            double originX = (double)window.PointsToScreenPixelsX(0f);
            double originY = (double)window.PointsToScreenPixelsY(0f);
            double scaleX = ((double)window.PointsToScreenPixelsX(100f) - originX) / 100.0;
            double scaleY = ((double)window.PointsToScreenPixelsY(100f) - originY) / 100.0;
            result.AppendFormat(CultureInfo.InvariantCulture, "slide={0} origin={1},{2} scale={3:F4},{4:F4}\r\n", (int)slide.SlideID, originX, originY, scaleX, scaleY);
            foreach (var pair in ResolveCurrentPairs((object)window))
            {
                foreach (object shapeObject in new[] { pair.A, pair.B })
                {
                    dynamic shape = shapeObject;
                    double left = originX + (double)shape.Left * scaleX;
                    double top = originY + (double)shape.Top * scaleY;
                    double right = left + (double)shape.Width * scaleX;
                    double bottom = top + (double)shape.Height * scaleY;
                    var point = new Point { X = (int)Math.Round((left + right) / 2), Y = (int)Math.Round((top + bottom) / 2) };
                    IntPtr hitWindow = WindowFromPoint(point);
                    string range;
                    try
                    {
                        dynamic hit = window.RangeFromPoint(point.X, point.Y);
                        range = hit == null ? "none" : ((int)hit.Id).ToString(CultureInfo.InvariantCulture);
                    }
                    catch (Exception exception) { range = exception.GetType().Name + ":" + exception.Message; }
                    result.AppendFormat(CultureInfo.InvariantCulture,
                        "shape={0} id={1} screen={2:F1},{3:F1},{4:F1},{5:F1} center={6},{7} hitWindow={8} hitClass={9} canvas={10} rangeId={11}\r\n",
                        (string)shape.Name, (int)shape.Id, left, top, right, bottom, point.X, point.Y,
                        hitWindow, WindowClass(hitWindow), IsCanvasWindow(hitWindow), range);
                }
            }
            return result.ToString();
        }

        public void Tick()
        {
            RequireUiThread();
            if (disposed || drag == null || updating) return;
            try
            {
                if ((GetAsyncKeyState(0x1B) & 0x8000) != 0)
                {
                    Finish(true, "escape");
                    return;
                }
                if (GetAncestor(GetForegroundWindow(), GaRoot) != drag.RootWindow)
                {
                    Finish(true, "focus-lost");
                    return;
                }
                if (GetCapture() != drag.CaptureWindow)
                {
                    Finish(true, "capture-lost");
                    return;
                }
                Point point;
                if (!GetCursorPos(out point)) return;
                bool mouseDown = (GetAsyncKeyState(1) & 0x8000) != 0;
                Update(point, mouseDown);
                // Covers a release outside the application or an interrupted input stream.
                if (!mouseDown) Finish(false, "button-released");
            }
            catch (Exception exception)
            {
                SafeLog("drag-error", exception.GetType().Name + ": " + exception.Message);
                Finish(true, "update-error");
            }
        }

        private IntPtr OnMouse(int code, IntPtr message, IntPtr data)
        {
            // HC_NOREMOVE may repeat the same event. Never mutate state for it.
            if (code != HcAction || disposed || inCallback)
                return CallNextHookEx(hook, code, message, data);
            int kind = message.ToInt32();
            bool owned = drag != null;
            inCallback = true;
            try
            {
                var mouse = (MouseHookData)Marshal.PtrToStructure(data, typeof(MouseHookData));
                if (drag != null)
                {
                    if (kind == WmLeftUp)
                    {
                        Update(mouse.Position, false);
                        Finish(false, "mouse-up");
                        return new IntPtr(1);
                    }
                    if (kind == WmMouseMove || kind == WmLeftDown || kind == WmLeftDoubleClick)
                        return new IntPtr(1);
                }
                else if (kind == WmLeftDown && TryBegin(mouse))
                {
                    return new IntPtr(1);
                }
            }
            catch (Exception exception)
            {
                SafeLog("drag-hook-error", exception.GetType().Name + ": " + exception.Message);
                if (drag != null) Finish(true, "hook-error");
                // If we already swallowed down, never deliver its trailing message to PPT.
                if (owned) return new IntPtr(1);
            }
            finally { inCallback = false; }
            return CallNextHookEx(hook, code, message, data);
        }

        private bool TryBegin(MouseHookData mouse)
        {
            if (targets.Count == 0 || GetCapture() != IntPtr.Zero) return false;
            if ((int)application.Presentations.Count == 0) return false;
            dynamic window = application.ActiveWindow;
            if (window == null || (int)window.ViewType != 9) return false;
            // PowerPoint.DocumentWindow has no HWND property. The hook supplies
            // the real input HWND; constrain it to this process's canvas instead.
            IntPtr root = GetAncestor(mouse.Window, GaRoot);
            if (root == IntPtr.Zero || GetAncestor(GetForegroundWindow(), GaRoot) != root ||
                !IsCanvasWindow(mouse.Window)) return false;
            uint processId;
            if (GetWindowThreadProcessId(mouse.Window, out processId) != uiThreadId || processId != GetCurrentProcessId()) return false;

            double originX = (double)window.PointsToScreenPixelsX(0f);
            double originY = (double)window.PointsToScreenPixelsY(0f);
            double scaleX = ((double)window.PointsToScreenPixelsX(100f) - originX) / 100.0;
            double scaleY = ((double)window.PointsToScreenPixelsY(100f) - originY) / 100.0;
            if (!FinitePositive(scaleX) || !FinitePositive(scaleY)) return false;
            int slideId = (int)window.View.Slide.SlideID;
            dynamic hitShape = window.RangeFromPoint(mouse.Position.X, mouse.Position.Y);
            if (hitShape == null) return false;
            int hitShapeId = (int)hitShape.Id;
            DragState candidate = null;
            // Resolve XML and unique IDs from the actual active presentation at
            // mouse-down, never reuse another window's cached COM shape objects.
            foreach (var pair in ResolveCurrentPairs((object)window))
            {
                dynamic a = pair.A;
                dynamic b = pair.B;
                if (!IsRectangle(a) || !IsRectangle(b)) continue;
                bool hitA = hitShapeId == (int)a.Id && HitInterior(a, mouse.Position, originX, originY, scaleX, scaleY);
                bool hitB = hitShapeId == (int)b.Id && HitInterior(b, mouse.Position, originX, originY, scaleX, scaleY);
                if (hitA && hitB) return false;
                if (!hitA && !hitB) continue;
                // Ambiguous overlap is left to PowerPoint rather than selecting an arbitrary pair.
                if (candidate != null) return false;
                dynamic driver = hitA ? a : b;
                candidate = new DragState
                {
                    Pair = pair,
                    Driver = driver,
                    Follower = hitA ? b : a,
                    DriverId = hitA ? pair.Constraint.ShapeA : pair.Constraint.ShapeB,
                    StartMouse = mouse.Position,
                    StartLeft = (double)driver.Left,
                    StartTopA = (double)a.Top,
                    StartTopB = (double)b.Top,
                    StartDriverTop = (double)driver.Top,
                    StartFollowerLeft = (double)(hitA ? b : a).Left,
                    LastLeft = (double)driver.Left,
                    LastTop = (double)driver.Top,
                    ScaleX = scaleX,
                    ScaleY = scaleY,
                    RootWindow = root,
                    CaptureWindow = mouse.Window
                };
            }
            if (candidate == null) return false;
            dynamic selected = candidate.Driver;
            selected.Select();
            SetCapture(candidate.CaptureWindow);
            if (GetCapture() != candidate.CaptureWindow) return false;
            drag = candidate;
            SafeLog("drag-start", String.Format(CultureInfo.InvariantCulture,
                "driver={0} slide={1} left={2:F3} a={3:F3} b={4:F3} cursor={5},{6} scale={7:F4},{8:F4}",
                candidate.DriverId, slideId, candidate.StartLeft, candidate.StartTopA, candidate.StartTopB,
                mouse.Position.X, mouse.Position.Y, scaleX, scaleY));
            return true;
        }

        private static List<DragPair> ResolveCurrentPairs(object windowObject)
        {
            dynamic window = windowObject;
            dynamic slide = window.View.Slide;
            dynamic parts = window.Presentation.CustomXMLParts.SelectByNamespace(ConstraintDocument.NamespaceUri);
            var result = new List<DragPair>();
            if ((int)parts.Count == 0) return result;
            if ((int)parts.Count != 1) throw new InvalidOperationException("Duplicate constraint XML parts.");
            ConstraintDocument document = ConstraintDocument.Parse((string)parts[1].XML);
            int slideId = (int)slide.SlideID;
            foreach (var constraint in document.Constraints)
            {
                if (constraint.SlideId != slideId) continue;
                object a = null, b = null;
                int countA = 0, countB = 0;
                for (int i = 1; i <= (int)slide.Shapes.Count; i++)
                {
                    dynamic shape = slide.Shapes[i];
                    string id = ((string)shape.Tags["PPTGCS_ID"]).ToUpperInvariant();
                    if (id == constraint.ShapeA) { a = shape; countA++; }
                    if (id == constraint.ShapeB) { b = shape; countB++; }
                }
                if (countA == 1 && countB == 1) result.Add(new DragPair(a, b, constraint));
            }
            return result;
        }

        private static string WindowClass(IntPtr window)
        {
            var name = new StringBuilder(256);
            GetClassName(window, name, name.Capacity);
            return name.ToString();
        }

        private static bool IsCanvasWindow(IntPtr window)
        {
            // The tested desktop PowerPoint editing surface is mdiClass.
            // RangeFromPoint additionally verifies the visible shape at the point.
            for (int i = 0; i < 8 && window != IntPtr.Zero; i++, window = GetParent(window))
            {
                if (String.Equals(WindowClass(window), "mdiClass", StringComparison.Ordinal)) return true;
            }
            return false;
        }

        private void Update(Point point, bool mouseDown)
        {
            if (drag == null || updating) return;
            updating = true;
            try
            {
                DragState state = drag;
                double left = state.StartLeft + (point.X - state.StartMouse.X) / state.ScaleX;
                double top = state.StartDriverTop + (point.Y - state.StartMouse.Y) / state.ScaleY;
                if (!SameTopSolver.NeedsUpdate(state.LastLeft, left) && !SameTopSolver.NeedsUpdate(state.LastTop, top)) return;
                if (!state.UndoStarted)
                {
                    application.StartNewUndoEntry();
                    state.UndoStarted = true;
                }
                dynamic driver = state.Driver;
                dynamic follower = state.Follower;
                driver.Left = (float)left;
                driver.Top = (float)top;
                follower.Top = (float)SameTopSolver.Solve(state.Pair.Constraint, state.DriverId, (double)driver.Top).Top;
                state.LastLeft = (double)driver.Left;
                state.LastTop = (double)driver.Top;
                state.Frames++;
                dynamic a = state.Pair.A;
                dynamic b = state.Pair.B;
                SafeLog("drag-frame", String.Format(CultureInfo.InvariantCulture,
                    "frame={0} mouseDown={1} driver={2} left={3:F3} a={4:F3} b={5:F3} followerLeft={6:F3} cursor={7},{8}",
                    state.Frames, mouseDown, state.DriverId, state.LastLeft, (double)a.Top, (double)b.Top,
                    (double)follower.Left, point.X, point.Y));
            }
            finally { updating = false; }
        }

        private void Finish(bool cancel, string reason)
        {
            DragState state = drag;
            if (state == null) return;
            // Clear before COM or capture release, either of which can pump messages.
            drag = null;
            try
            {
                dynamic a = state.Pair.A;
                dynamic b = state.Pair.B;
                dynamic driver = state.Driver;
                if (cancel && state.UndoStarted)
                {
                    driver.Left = (float)state.StartLeft;
                    a.Top = (float)state.StartTopA;
                    b.Top = (float)state.StartTopB;
                }
                if (state.UndoStarted) application.StartNewUndoEntry();
                SafeLog("drag-end", String.Format(CultureInfo.InvariantCulture,
                    "reason={0} cancelled={1} frames={2} driver={3} leftBefore={4:F3} leftAfter={5:F3} aBefore={6:F3} aAfter={7:F3} bBefore={8:F3} bAfter={9:F3}",
                    reason, cancel, state.Frames, state.DriverId, state.StartLeft, (double)driver.Left,
                    state.StartTopA, (double)a.Top, state.StartTopB, (double)b.Top));
            }
            catch (Exception exception)
            {
                SafeLog("drag-end-error", reason + " " + exception.GetType().Name + ": " + exception.Message);
            }
            finally
            {
                if (GetCapture() == state.CaptureWindow) ReleaseCapture();
            }
        }

        private static bool IsRectangle(dynamic shape)
        {
            return (int)shape.Type == 1 && (int)shape.AutoShapeType == 1 &&
                Math.Abs((double)shape.Rotation) < 0.01 && (int)shape.Visible != 0;
        }

        private static bool HitInterior(dynamic shape, Point point, double originX, double originY, double scaleX, double scaleY)
        {
            double left = originX + (double)shape.Left * scaleX;
            double top = originY + (double)shape.Top * scaleY;
            double right = left + (double)shape.Width * scaleX;
            double bottom = top + (double)shape.Height * scaleY;
            return point.X > left + HandleMargin && point.X < right - HandleMargin &&
                point.Y > top + HandleMargin && point.Y < bottom - HandleMargin;
        }

        private static bool FinitePositive(double value)
        {
            return !Double.IsNaN(value) && !Double.IsInfinity(value) && value > 0;
        }

        private void SafeLog(string type, string detail)
        {
            try { log(type, detail); } catch { /* A logger must not unwind through a native hook. */ }
        }

        private void RequireUiThread()
        {
            if (GetCurrentThreadId() != uiThreadId)
                throw new InvalidOperationException("The drag controller must run on PowerPoint's UI thread.");
        }

        public void Dispose()
        {
            RequireUiThread();
            if (disposed) return;
            disposed = true;
            Finish(true, "dispose");
            if (hook != IntPtr.Zero)
            {
                UnhookWindowsHookEx(hook);
                hook = IntPtr.Zero;
            }
            targets.Clear();
            SafeLog("drag-hook-removed", "thread=" + uiThreadId);
        }

        private sealed class DragState
        {
            public DragPair Pair;
            public object Driver;
            public object Follower;
            public string DriverId;
            public Point StartMouse;
            public double StartLeft;
            public double StartTopA;
            public double StartTopB;
            public double StartDriverTop;
            public double StartFollowerLeft;
            public double LastLeft;
            public double LastTop;
            public double ScaleX;
            public double ScaleY;
            public IntPtr RootWindow;
            public IntPtr CaptureWindow;
            public bool UndoStarted;
            public int Frames;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Point { public int X; public int Y; }

        [StructLayout(LayoutKind.Sequential)]
        private struct MouseHookData
        {
            public Point Position;
            public IntPtr Window;
            public uint HitTestCode;
            public UIntPtr ExtraInfo;
        }

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate IntPtr MouseHook(int code, IntPtr message, IntPtr data);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int hookType, MouseHook callback, IntPtr module, uint threadId);
        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hook);
        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr message, IntPtr data);
        [DllImport("user32.dll")]
        private static extern IntPtr SetCapture(IntPtr window);
        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        private static extern IntPtr GetCapture();
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int key);
        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out Point point);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr window, StringBuilder name, int maxCount);
        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr window);
        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(Point point);
        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentProcessId();
    }
}
