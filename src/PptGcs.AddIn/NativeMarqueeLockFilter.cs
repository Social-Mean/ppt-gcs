using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PptGcs.AddIn
{
    // PowerPoint can include locked shapes in the result of its own marquee.
    // Leave the pointer gesture and rectangle to Office and prune that result
    // using the same native Locked property as the two selection panes.
    internal sealed class NativeMarqueeLockFilter
    {
        private readonly dynamic app;
        private readonly Action<string, object> log;
        private object window;
        private IntPtr root;
        private int slideId;
        private Point origin;
        private bool moved, released;

        public NativeMarqueeLockFilter(object app, Action<string, object> log)
        { this.app = app; this.log = log; }

        public void Begin(int x, int y, IntPtr target)
        {
            if (!ConstrainedDragController.IsCanvasWindow(target) || (int)app.Presentations.Count == 0) return;
            var foreground = GetForegroundWindow();
            if (foreground == IntPtr.Zero || GetAncestor(target, 2) != foreground) return;
            dynamic activeWindow = app.ActiveWindow;
            if ((int)activeWindow.ViewType != 9) return;
            if ((int)activeWindow.ActivePane.ViewType == 1)
            {
                dynamic hit = activeWindow.RangeFromPoint(x, y);
                if (hit != null && !ShapeCapabilities.IsLocked((object)hit)) return;
            }
            slideId = (int)activeWindow.View.Slide.SlideID;
            window = (object)activeWindow;
            root = foreground;
            origin = new Point(x, y);
            moved = released = false;
        }

        public void ObserveMouse(int kind, int x, int y)
        {
            if (window == null) return;
            if (kind == 0x201 || kind == 0x203 || kind == 0x204 || kind == 0x207 || kind == 0x20A || kind == 0x20E)
            {
                window = null;
                return;
            }
            if (kind != 0x200 && kind != 0x202) return;
            Size threshold = SystemInformation.DragSize;
            moved |= Math.Abs(x - origin.X) >= Math.Max(3, threshold.Width / 2) ||
                Math.Abs(y - origin.Y) >= Math.Max(3, threshold.Height / 2);
            released |= kind == 0x202;
        }

        public void Cancel() { window = null; }

        public void Finish(bool selectionChanged = false)
        {
            if (window == null) return;
            if (GetForegroundWindow() != root) { Cancel(); return; }
            if (!released && (GetAsyncKeyState(1) < 0 || selectionChanged)) return;
            if (!selectionChanged && GetCapture() != IntPtr.Zero) return;
            var expectedWindow = window;
            Cancel();
            if (!moved) return;
            try
            {
                if ((int)app.Presentations.Count == 0) return;
                dynamic nativeWindow = app.ActiveWindow;
                if (!Equals(expectedWindow, (object)nativeWindow) || (int)nativeWindow.ViewType != 9) return;
                dynamic slide = nativeWindow.View.Slide;
                if ((int)slide.SlideID != slideId) return;
                dynamic selection = nativeWindow.Selection;
                if ((int)selection.Type != 2 || (bool)selection.HasChildShapeRange) return;
                dynamic range = selection.ShapeRange;
                int count = (int)range.Count;
                var unlocked = Enumerable.Range(1, count).Select(index => (object)range[index])
                    .Where(shape => !ShapeCapabilities.IsLocked(shape))
                    .Select(shape => (int)((dynamic)shape).Id).ToHashSet();
                if (unlocked.Count == count) return;
                dynamic shapes = slide.Shapes;
                object[] indexes = Enumerable.Range(1, (int)shapes.Count)
                    .Where(index => unlocked.Contains((int)shapes[index].Id)).Select(index => (object)index).ToArray();
                if (indexes.Length == 0) selection.Unselect();
                else shapes.Range(indexes).Select();
                log("layer-native-marquee-filtered", new { slideId, excluded = count - unlocked.Count });
            }
            catch (Exception error) { log("layer-native-marquee-error", new { error = error.Message }); }
        }

        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetCapture();
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
    }
}
