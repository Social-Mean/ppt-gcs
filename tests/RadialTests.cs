using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using PptGcs.AddIn;
using PptGcs.Core;

internal static class RadialTests
{
    private static int passed;

    [STAThread]
    private static void Main()
    {
        // No PowerPoint COM object, window, form, or synthesized Win32 input is
        // created here. The fixture injects the already-open menu's pure state;
        // production mouse/keyboard handlers and preview callbacks do the work.
        int[,] directions = {
            {500, 300}, {571, 329}, {600, 400}, {571, 471},
            {500, 500}, {429, 471}, {400, 400}, {429, 329}
        };
        using (var fixture = new Fixture())
        {
            for (int index = 0; index < 8; index++)
                Check(fixture.Hit(directions[index, 0], directions[index, 1]) == index,
                    "fixed radial direction " + index);
            Check(fixture.Hit(500, 400) == -1, "center has no command");
            Check(fixture.Hit(520, 400) == -1, "inner dead zone has no command");
            Check(fixture.Hit(50, 50) == -1, "pointer outside the canvas has no command");
        }

        using (var fixture = new Fixture())
        {
            fixture.Move(600, 400);
            Check(fixture.PreviewCalls == 1 && fixture.ExecuteCalls == 0 && fixture.PaneCalls == 0 &&
                fixture.NormalizeCalls == 0 && fixture.SourceGeometryUnchanged,
                "hover previews without committing or changing document state");
            fixture.Move(500, 400);
            fixture.Move(600, 400);
            Check(fixture.PreviewCalls == 1 && fixture.ExecuteCalls == 0,
                "returning to a sector reuses its read-only preview");
        }

        using (var fixture = new Fixture())
        {
            fixture.ReleaseDuringNextPreview = true;
            fixture.Move(600, 400); // SameWidth hover; release below selects SameY.
            Check(fixture.NestedReleaseWasQueued, "nested release waits until preview returns");
            Check(fixture.ExecuteCalls == 1 && fixture.ExecutedKind == "SameY" &&
                fixture.WasClosedWhenExecuted && !fixture.Controller.IsOpen,
                "exact release sector closes the menu before one command");
        }

        using (var fixture = new Fixture())
        {
            fixture.Move(600, 400);
            bool consumed = fixture.Controller.HandleKeyboard(0x1B, true);
            Check(consumed && !fixture.Controller.IsOpen && fixture.ExecuteCalls == 0 &&
                fixture.PaneCalls == 0 && fixture.SourceGeometryUnchanged,
                "Escape cancels the preview without changing the document");
        }

        using (var fixture = new Fixture())
        {
            fixture.Move(600, 400);
            fixture.Controller.HandleMouse(0x0208, 500, 400, IntPtr.Zero);
            Check(!fixture.Controller.IsOpen && fixture.ExecuteCalls == 0 && fixture.PaneCalls == 0 &&
                fixture.SourceGeometryUnchanged,
                "returning to the center before release cancels the command");
        }

        using (var fixture = new Fixture())
        {
            fixture.Move(600, 400);
            bool consumed = fixture.Controller.HandleMouse(0x020A, 500, 400, IntPtr.Zero);
            Check(!consumed && !fixture.Controller.IsOpen && fixture.ExecuteCalls == 0,
                "wheel scrolling cancels the menu and remains native input");
        }

        Console.WriteLine("PASS: " + passed + " radial state tests (no physical mouse input)");
    }

    private static void Check(bool condition, string description)
    {
        if (!condition) throw new Exception(description);
        passed++;
        Console.WriteLine("PASS " + description);
    }

    private sealed class Fixture : IDisposable
    {
        private static readonly Type ControllerType = typeof(RadialMenuController);
        private const BindingFlags PrivateInstance = BindingFlags.Instance | BindingFlags.NonPublic;
        private readonly ShapeGeometry source = new ShapeGeometry("AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA", 110, 120, 180, 70);
        public readonly RadialMenuController Controller;
        public int PreviewCalls, ExecuteCalls, PaneCalls, NormalizeCalls;
        public string ExecutedKind;
        public bool ReleaseDuringNextPreview, NestedReleaseWasQueued, WasClosedWhenExecuted;

        public Fixture()
        {
            Controller = new RadialMenuController(() => new ConstraintPaneModel(), Preview,
                kind => {
                    WasClosedWhenExecuted = !Controller.IsOpen;
                    ExecuteCalls++;
                    ExecutedKind = kind;
                }, more => PaneCalls++, () => null, () => NormalizeCalls++, (kind, value) => { });
            Set("canvasBounds", new Rectangle(100, 100, 800, 600));
            Set("center", new PointF(400, 300));
            Set("startPoint", NativePoint(500, 400));
            Set("outerRadius", 136f); Set("innerRadius", 38f);
            Set("originX", 100d); Set("originY", 100d);
            Set("scaleX", 1d); Set("scaleY", 1d);
            Set("selectedCount", 2); Set("open", true); Set("currentSlot", -1);
            // Eligibility depends on native selection at Begin, which belongs to
            // the host tests. Here all directions are eligible so routing and
            // preview behavior can be exercised without any native window.
            foreach (object slot in (Array)Get("slots"))
            {
                slot.GetType().GetField("Enabled").SetValue(slot, true);
                slot.GetType().GetField("Previewed").SetValue(slot, false);
            }
        }

        public bool SourceGeometryUnchanged
        {
            get { return source.Left == 110 && source.Top == 120 && source.Width == 180 && source.Height == 70; }
        }

        private IList<ShapeGeometry> Preview(string kind)
        {
            PreviewCalls++;
            if (ReleaseDuringNextPreview)
            {
                ReleaseDuringNextPreview = false;
                Controller.HandleMouse(0x0208, 571, 329, IntPtr.Zero);
                NestedReleaseWasQueued = Controller.IsOpen && ExecuteCalls == 0;
            }
            return new List<ShapeGeometry> { source };
        }

        public void Move(int x, int y) { Controller.HandleMouse(0x0200, x, y, IntPtr.Zero); }
        public int Hit(int x, int y)
        { return (int)ControllerType.GetMethod("HitSlot", PrivateInstance).Invoke(Controller, new[] { NativePoint(x, y) }); }
        private void Set(string name, object value)
        { ControllerType.GetField(name, PrivateInstance).SetValue(Controller, value); }
        private object Get(string name)
        { return ControllerType.GetField(name, PrivateInstance).GetValue(Controller); }
        private static object NativePoint(int x, int y)
        {
            Type type = ControllerType.GetNestedType("NativePoint", BindingFlags.NonPublic);
            object point = Activator.CreateInstance(type);
            type.GetField("X").SetValue(point, x);
            type.GetField("Y").SetValue(point, y);
            return point;
        }
        public void Dispose() { Controller.Dispose(); }
    }
}
