using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PptGcs.AddIn
{
    internal sealed class LayerMarqueeController : IDisposable
    {
        private readonly Func<SlideLayers> currentSlide;
        private readonly Func<SlideLayers, string> activeLayer;
        private readonly Action started, finished;
        private readonly Action<string, object> log;
        private Drag drag;
        private object armedWindow;
        private int armedSlide;
        private bool consumeUp, consumeEscapeUp;

        internal LayerMarqueeController(Func<SlideLayers> currentSlide, Func<SlideLayers, string> activeLayer,
            Action started, Action finished, Action<string, object> log)
        {
            this.currentSlide = currentSlide; this.activeLayer = activeLayer;
            this.started = started; this.finished = finished; this.log = log;
        }

        internal bool IsDragging { get { return drag != null; } }
        internal bool IsArmed(SlideLayers slide)
        { return slide != null && Equals(armedWindow, slide.Window) && armedSlide == slide.SlideId; }

        internal void Toggle(SlideLayers slide)
        {
            if (IsArmed(slide)) armedWindow = null;
            else { armedWindow = slide.Window; armedSlide = slide.SlideId; }
        }

        internal bool HandleBeforeDrag(int message, int x, int y, IntPtr target)
        {
            if (message == 0x201) consumeUp = false;
            if (drag != null)
            {
                if (message == 0x204 || message == 0x207 || message == 0x20A || message == 0x20E)
                {
                    Cancel();
                    return false;
                }
                if (message == 0x200 || message == 0x202)
                {
                    // Native mouse hooks also receive nested messages while COM
                    // calls pump the STA. Defer all COM selection to Tick.
                    drag.Last = Clamp(drag.Canvas, new Point(x, y));
                    drag.Released |= message == 0x202;
                    return true;
                }
                return message == 0x201 || message == 0x203;
            }
            if (message == 0x202 && consumeUp) { consumeUp = false; return true; }
            return message == 0x201 && armedWindow != null && TryStart(x, y, target, true);
        }

        internal bool TryStart(int x, int y, IntPtr target, bool armedOnly)
        {
            if (!ConstrainedDragController.IsCanvasWindow(target) || GetCapture() != IntPtr.Zero ||
                GetKeyState(0x12) < 0) return false;
            IntPtr root = GetAncestor(target, 2);
            if (root == IntPtr.Zero || root != GetForegroundWindow()) return false;
            var slide = currentSlide();
            if (slide == null) return false;
            bool armed = IsArmed(slide);
            if (!armed && (armedOnly || slide.Stack.Length < 2)) return false;
            var point = new Point(x, y);
            var canvas = ClientBounds(target);
            if (!canvas.Contains(point)) return false;
            dynamic window = slide.Window;
            ConstrainedDragController.ActivateSlidePane(window);
            var map = ConstrainedDragController.ScreenMap.From(slide.Window);
            if (!armed)
            {
                if ((int)window.Selection.Type == 3 ||
                    ((int)window.Selection.Type == 2 && (bool)window.Selection.HasChildShapeRange)) return false;
                if (window.RangeFromPoint(x, y) != null || HitsSelectionHandle(slide, map, point)) return false;
            }
            string layer = activeLayer(slide);
            SetFocus(target);
            SetCapture(target);
            if (GetCapture() != target) return false;
            armedWindow = null;
            drag = new Drag {
                Slide = slide, Layer = layer, Target = target, Root = root, Canvas = canvas,
                Map = map, Origin = point, Last = point, Armed = armed
            };
            started();
            return true;
        }

        internal bool HandleKeyboard(int key, bool down)
        {
            if (key != 0x1B) return false;
            if (!down && consumeEscapeUp) { consumeEscapeUp = false; return true; }
            if (!down || (drag == null && armedWindow == null)) return false;
            consumeEscapeUp = true;
            Cancel();
            return true;
        }

        internal void Tick()
        {
            try
            {
                if (drag == null)
                {
                    if (armedWindow != null && !IsArmed(currentSlide())) Cancel();
                    return;
                }
                Drag state = drag;
                if (GetForegroundWindow() != state.Root || GetCapture() != state.Target)
                { Cancel(); return; }
                var box = Rectangle.FromLTRB(Math.Min(state.Origin.X, state.Last.X), Math.Min(state.Origin.Y, state.Last.Y),
                    Math.Max(state.Origin.X, state.Last.X), Math.Max(state.Origin.Y, state.Last.Y));
                Size threshold = SystemInformation.DragSize;
                state.Started |= box.Width >= threshold.Width / 2 || box.Height >= threshold.Height / 2;
                if (!state.Released)
                {
                    if (state.Started && box != state.Box) Render(state, box);
                    return;
                }
                var slide = currentSlide();
                Cancel();
                if (slide == null || slide.SlideId != state.Slide.SlideId || !Equals(slide.Window, state.Slide.Window)) return;
                if (!state.Started)
                {
                    if (!state.Armed) slide.Select(new SlideLayers.Member[0]);
                    return;
                }
                box.Inflate(1, 1);
                var selected = slide.Members().Where(member => member.LayerId == state.Layer &&
                    (int)((dynamic)member.Shape).Visible != 0 && box.Contains(Bounds(member.Shape, state.Map))).ToArray();
                slide.Select(selected);
                log("layer-marquee", new { slideId = slide.SlideId, layer = state.Layer, selected = selected.Select(member => member.Id).ToArray() });
            }
            catch (Exception error)
            {
                Cancel();
                log("layer-marquee-error", new { error = error.Message });
            }
        }

        private static bool HitsSelectionHandle(SlideLayers slide, ConstrainedDragController.ScreenMap map, Point point)
        {
            var selection = slide.Selected();
            if (selection.Count == 0) return false;
            var padding = Math.Max(4, (int)(GetDpiForWindow(GetForegroundWindow()) / 12));
            var rectangles = slide.Members().Where(member => selection.Contains(member.Id)).Select(member => Bounds(member.Shape, map)).ToArray();
            if (rectangles.Length == 0) return false;
            var bounds = rectangles.Aggregate(Rectangle.Union);
            var inner = bounds;
            inner.Inflate(-padding, -padding);
            bounds.Inflate(padding, padding);
            if (bounds.Contains(point) && !inner.Contains(point)) return true;
            int center = inner.Left + inner.Width / 2;
            return Math.Abs(point.X - center) <= padding && point.Y >= bounds.Top - 3 * padding && point.Y < bounds.Top;
        }

        private static Rectangle Bounds(object shapeObject, ConstrainedDragController.ScreenMap map)
        {
            dynamic shape = shapeObject;
            double width = (double)shape.Width, height = (double)shape.Height;
            double x = (double)shape.Left + width / 2, y = (double)shape.Top + height / 2;
            double angle = (double)shape.Rotation * Math.PI / 180;
            double halfWidth = (Math.Abs(width * Math.Cos(angle)) + Math.Abs(height * Math.Sin(angle))) / 2;
            double halfHeight = (Math.Abs(width * Math.Sin(angle)) + Math.Abs(height * Math.Cos(angle))) / 2;
            var first = map.Point(x - halfWidth, y - halfHeight);
            var last = map.Point(x + halfWidth, y + halfHeight);
            return Rectangle.FromLTRB(first.X, first.Y, Math.Max(first.X + 1, last.X), Math.Max(first.Y + 1, last.Y));
        }

        private static void Render(Drag state, Rectangle box)
        {
            if (state.Surface == null)
            {
                state.Overlay = new Overlay();
                state.Bitmap = new Bitmap(state.Canvas.Width, state.Canvas.Height, PixelFormat.Format32bppPArgb);
                state.Surface = new LayeredWindowSurface(state.Canvas.Size);
                state.Overlay.Show(new Owner(state.Root));
            }
            state.Box = box;
            box.Offset(-state.Canvas.X, -state.Canvas.Y);
            box.Width = Math.Max(1, box.Width);
            box.Height = Math.Max(1, box.Height);
            using (var graphics = Graphics.FromImage(state.Bitmap))
            using (var fill = new SolidBrush(Color.FromArgb(42, 0, 112, 192)))
            using (var outline = new Pen(Color.FromArgb(235, 0, 112, 192)))
            {
                graphics.Clear(Color.Transparent);
                graphics.FillRectangle(fill, box);
                graphics.DrawRectangle(outline, box);
            }
            state.Surface.Present(state.Overlay.Handle, state.Bitmap, state.Canvas.Location);
        }

        internal void Cancel()
        {
            bool hadState = drag != null || armedWindow != null;
            armedWindow = null;
            Drag state = drag;
            drag = null;
            if (state != null)
            {
                consumeUp = !state.Released;
                if (GetCapture() == state.Target) ReleaseCapture();
                if (state.Overlay != null) state.Overlay.Dispose();
                if (state.Surface != null) state.Surface.Dispose();
                if (state.Bitmap != null) state.Bitmap.Dispose();
            }
            if (hadState) finished();
        }

        public void Dispose() { Cancel(); }

        private static Point Clamp(Rectangle bounds, Point point)
        { return new Point(Math.Max(bounds.Left, Math.Min(bounds.Right, point.X)), Math.Max(bounds.Top, Math.Min(bounds.Bottom, point.Y))); }

        private static Rectangle ClientBounds(IntPtr window)
        {
            Rect rect;
            var origin = new NativePoint();
            if (!GetClientRect(window, out rect) || !ClientToScreen(window, ref origin)) return Rectangle.Empty;
            return new Rectangle(origin.X, origin.Y, rect.Right, rect.Bottom);
        }

        private sealed class Drag
        {
            public SlideLayers Slide;
            public string Layer;
            public IntPtr Target, Root;
            public ConstrainedDragController.ScreenMap Map;
            public Point Origin, Last;
            public Rectangle Canvas, Box;
            public bool Released, Started, Armed;
            public Overlay Overlay;
            public LayeredWindowSurface Surface;
            public Bitmap Bitmap;
        }

        private sealed class Owner : IWin32Window
        { public IntPtr Handle { get; private set; } public Owner(IntPtr handle) { Handle = handle; } }

        private sealed class Overlay : Form
        {
            public Overlay() { FormBorderStyle = FormBorderStyle.None; ShowInTaskbar = false; StartPosition = FormStartPosition.Manual; Size = new Size(1, 1); }
            protected override bool ShowWithoutActivation { get { return true; } }
            protected override CreateParams CreateParams
            { get { var args = base.CreateParams; args.ExStyle |= 0x080800A0; return args; } }
        }

        [StructLayout(LayoutKind.Sequential)] private struct Rect { public int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] private struct NativePoint { public int X, Y; }
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out Rect rect);
        [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr window, ref NativePoint point);
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr SetFocus(IntPtr window);
        [DllImport("user32.dll")] private static extern IntPtr SetCapture(IntPtr window);
        [DllImport("user32.dll")] private static extern bool ReleaseCapture();
        [DllImport("user32.dll")] private static extern IntPtr GetCapture();
        [DllImport("user32.dll")] private static extern short GetKeyState(int key);
    }
}
