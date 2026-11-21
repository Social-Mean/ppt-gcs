using System;
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
        Console.WriteLine("PASS: " + passed + " shape support tests");
    }
    private static void Check(bool condition,string name)
    { if(!condition)throw new Exception(name);passed++; }
}

public sealed class SupportShape
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
