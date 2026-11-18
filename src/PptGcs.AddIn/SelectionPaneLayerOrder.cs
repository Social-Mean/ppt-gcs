using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace PptGcs.AddIn
{
    // The Selection Pane's list can reorder shapes without running any of the
    // four repurposable Ribbon commands. It closes its own native undo entry.
    internal sealed class SelectionPaneLayerOrder
    {
        private readonly object app;
        private readonly Func<SlideLayers> activeSlide;
        private readonly Action<string, object> log;
        private Gesture pending;
        private Replacement replacing;
        private string paneTitle;

        internal SelectionPaneLayerOrder(object app, Func<SlideLayers> activeSlide, Action<string, object> log)
        { this.app = app; this.activeSlide = activeSlide; this.log = log; }

        internal void ObserveMouse(int message, IntPtr target)
        {
            if (message == 0x202)
            {
                if (pending != null) pending.Released = DateTime.UtcNow;
                return;
            }
            if (message == 0x204 || message == 0x207 || message == 0x20A || message == 0x20E) Cancel();
            if (message != 0x201) return;
            Cancel();
            try
            {
                if (!IsSelectionPane(target)) return;
                var slide = activeSlide();
                if (slide == null || slide.Stack.Length < 2) return;
                var original = slide.Members();
                if (!original.SequenceEqual(slide.SortIntoLayers(original))) return;
                pending = new Gesture { Slide = slide, Original = original, Root = GetAncestor(target, 2) };
            }
            catch (Exception error) { log("layer-native-pane-error", new { error = error.Message }); }
        }

        internal void Cancel()
        {
            pending = null;
            if (replacing != null) replacing.Interrupted = true;
        }

        internal void Tick()
        {
            if (replacing != null) { FinishReplacement(); return; }
            Gesture gesture = pending;
            if (gesture == null) return;
            if (gesture.Root != GetForegroundWindow()) { Cancel(); return; }
            if (GetCapture() != IntPtr.Zero || GetAsyncKeyState(1) < 0) return;
            if (!gesture.Released.HasValue) gesture.Released = DateTime.UtcNow;
            var elapsed = DateTime.UtcNow - gesture.Released.Value;
            if (elapsed.TotalMilliseconds < 35) return;
            try
            {
                var slide = activeSlide();
                if (!Matches(gesture, slide)) { Cancel(); return; }
                var raw = slide.Members();
                var corrected = slide.SortIntoLayers(raw);
                if (raw.SequenceEqual(corrected))
                {
                    if (elapsed.TotalMilliseconds >= 250) Cancel();
                    return;
                }
                Cancel();
                if (!SameMembership(gesture.Original, raw)) return;
                ReplaceNativeUndoEntry(gesture, slide, raw, corrected);
            }
            catch (Exception error)
            {
                Cancel();
                log("layer-native-pane-error", new { error = error.ToString() });
            }
        }

        private void ReplaceNativeUndoEntry(Gesture gesture, SlideLayers slide, SlideLayers.Member[] raw, SlideLayers.Member[] corrected)
        {
            dynamic bars = ((dynamic)app).CommandBars;
            if (!(bool)bars.GetEnabledMso("Undo")) return;
            replacing = new Replacement {
                Gesture = gesture, Raw = raw, Corrected = corrected, Selected = slide.Selected(), Issued = DateTime.UtcNow
            };
            // PowerPoint can queue ExecuteMso when called from its own UI
            // thread. Inspect the resulting native history on a later tick.
            bars.ExecuteMso("Undo");
        }

        private void FinishReplacement()
        {
            var next = replacing;
            try
            {
                var slide = activeSlide();
                if (!Matches(next.Gesture, slide)) { replacing = null; return; }
                var current = slide.Members();
                if (SameOrder(next.Gesture.Original, current))
                {
                    replacing = null;
                    slide.ReapplyNativeOrder(app, next.Raw.Select(member => member.Id).ToArray(),
                        next.Corrected.Select(member => member.Id).ToArray());
                    if (next.Selected.Count != 0 && !slide.Selected().SetEquals(next.Selected))
                        slide.Select(slide.Members().Where(member => next.Selected.Contains(member.Id)));
                    log("layer-native-pane-order", new { slideId = slide.SlideId });
                    return;
                }
                if ((DateTime.UtcNow - next.Issued).TotalMilliseconds < 450) return;
                replacing = null;
                // If the captured edit was overtaken, restore whatever Undo
                // consumed, but never send Redo over subsequent user input.
                if (!next.Interrupted && !SameOrder(next.Raw, current))
                {
                    dynamic bars = ((dynamic)app).CommandBars;
                    if ((bool)bars.GetEnabledMso("Redo")) bars.ExecuteMso("Redo");
                }
                log("layer-native-pane-history-mismatch", new { slideId = slide.SlideId });
            }
            catch (Exception error)
            {
                replacing = null;
                log("layer-native-pane-error", new { error = error.ToString() });
            }
        }

        private static bool Matches(Gesture gesture, SlideLayers slide)
        {
            return slide != null && slide.SlideId == gesture.Slide.SlideId &&
                Equals(slide.Window, gesture.Slide.Window) && slide.Encoded == gesture.Slide.Encoded;
        }

        private static bool SameMembership(SlideLayers.Member[] left, SlideLayers.Member[] right)
        { return SameOrder(left.OrderBy(member => member.Id).ToArray(), right.OrderBy(member => member.Id).ToArray()); }

        private static bool SameOrder(SlideLayers.Member[] left, SlideLayers.Member[] right)
        {
            return left.Length == right.Length && left.Zip(right,
                (first, second) => first.Id == second.Id && first.StoredTag == second.StoredTag).All(same => same);
        }

        private bool IsSelectionPane(IntPtr target)
        {
            if (target == IntPtr.Zero || GetAncestor(target, 2) != GetForegroundWindow()) return false;
            var buffer = new StringBuilder(256);
            for (IntPtr parent = target; parent != IntPtr.Zero; parent = GetParent(parent))
            {
                GetClassName(parent, buffer, buffer.Capacity);
                if (buffer.ToString() != "MsoWorkPane") continue;
                if (paneTitle == null) paneTitle = (string)((dynamic)app).CommandBars["Selection"].NameLocal;
                GetWindowText(parent, buffer, buffer.Capacity);
                return buffer.ToString() == paneTitle;
            }
            return false;
        }

        private sealed class Gesture
        {
            public SlideLayers Slide;
            public SlideLayers.Member[] Original;
            public IntPtr Root;
            public DateTime? Released;
        }

        private sealed class Replacement
        {
            public Gesture Gesture;
            public SlideLayers.Member[] Raw, Corrected;
            public System.Collections.Generic.HashSet<int> Selected;
            public DateTime Issued;
            public bool Interrupted;
        }

        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetCapture();
        [DllImport("user32.dll")] private static extern IntPtr GetParent(IntPtr window);
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
    }
}
