using System;
using System.Runtime.InteropServices;
using PptGcs.AddIn;
using PptGcs.Core;

internal static class ShapeSupportTests
{
    private static int passed;
    private static void Main()
    {
        foreach (int type in new[] {1,5,6,11,13,17,28,29})
            Check(ShapeCapabilities.IsSupported(new SupportShape {Type=type}), "bounding-box type " + type);
        foreach (int type in new[] {3,7,9,12,14,16,19,24})
            Check(!ShapeCapabilities.IsSupported(new SupportShape {Type=type}), "excluded object type " + type);
        var shape=new SupportShape();
        shape.Rotation=360;
        Check(ShapeCapabilities.IsSupported(shape), "a complete turn has zero rotation");
        shape.Rotation=15;
        Check(!ShapeCapabilities.IsSupported(shape), "rotated shapes are rejected explicitly");
        shape.Rotation=0;shape.Width=0;
        Check(!ShapeCapabilities.IsSupported(shape), "zero width is rejected before solving");
        shape.Width=100;shape.Height=Double.NaN;
        Check(!ShapeCapabilities.IsSupported(shape), "invalid dimensions are rejected before solving");
        shape.Height=80;shape.Top=Double.PositiveInfinity;
        Check(!ShapeCapabilities.IsSupported(shape), "invalid coordinates are rejected before solving");
        shape.Top=20;
        var group=new SupportShape {Type=6};
        var nested=new SupportShape {Type=6,Child=-1,ParentGroup=group};
        shape.Child=-1;shape.ParentGroup=nested;
        Check(Object.ReferenceEquals(ShapeCapabilities.TopLevel(shape),group), "nested group hits resolve to the whole top-level group");
        Check(Object.ReferenceEquals(ShapeCapabilities.TopLevel(group),group), "top-level hit retains the original object");
        var context = new LayoutContext();
        const string groupId = "AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA";
        context.Shapes.Add(groupId, group);
        PowerPointLayout.Apply(context, new[] { new ShapeGeometry(groupId, 40, 50, 100, 80) });
        Check(group.Left == 40 && group.Top == 50, "constraint writeback works without the native Locked member");
        CheckTopAlignmentWithUnimplementedLock();
        Console.WriteLine("PASS: " + passed + " shape support tests");
    }

    private static void CheckTopAlignmentWithUnimplementedLock()
    {
        var driver = new SupportShapeWithLock { ReadLock = () => throw new NotImplementedException() };
        var follower = new SupportShapeWithLock { Left = 140, ReadLock = () => throw new COMException("E_NOTIMPL", unchecked((int)0x80004001)) };
        string driverId = Guid.NewGuid().ToString("D").ToUpperInvariant();
        string followerId = Guid.NewGuid().ToString("D").ToUpperInvariant();
        var context = new LayoutContext();
        context.Shapes.Add(driverId, driver);
        context.Shapes.Add(followerId, follower);
        var relation = new LayoutConstraint(Guid.NewGuid().ToString(), 1, ConstraintKind.SameY, new[] {driverId, followerId});
        var target = new ShapeGeometry(driverId, driver.Left, 50, driver.Width, driver.Height);
        var solved = LayoutSolver.Solve(PowerPointLayout.Snapshot(context), new[] {relation}, driverId, target, GeometryEditMode.Move);
        PowerPointLayout.Apply(context, solved);
        Check(driver.Top == 50 && follower.Top == 50 && follower.Left == 140, "top-aligned shapes follow when Office reports E_NOTIMPL for Locked");

        follower.Child = -1;
        follower.ParentGroup = new SupportShapeWithLock { ReadLock = () => -1 };
        var lockedTarget = new ShapeGeometry(followerId, 140, 60, follower.Width, follower.Height);
        try { PowerPointLayout.Apply(context, new[] {lockedTarget}); }
        catch (InvalidOperationException)
        {
            Check(follower.Top == 50, "a locked parent still protects a child with an unimplemented lock getter");
            return;
        }
        throw new Exception("Expected the locked parent to prevent the move");
    }
    private static void Check(bool condition,string name)
    { if(!condition)throw new Exception(name);passed++; }
}

public class SupportShape
{
    public int Type {get;set;}=1;
    // Accessing AutoShapeType on pictures/text/groups is invalid in Office.
    public int AutoShapeType {get {throw new Exception("Do not require a rectangle outline");}}
    public double Rotation {get;set;}
    public double Left {get;set;}=10;
    public double Top {get;set;}=20;
    public double Width {get;set;}=100;
    public double Height {get;set;}=80;
    public int LockAspectRatio {get;set;}
    public int Child {get;set;}
    public object ParentGroup {get;set;}
}

public sealed class SupportShapeWithLock : SupportShape
{
    public Func<int> ReadLock { get; set; }
    public int Locked { get { return ReadLock(); } }
}
