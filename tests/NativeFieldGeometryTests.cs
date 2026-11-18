using System;
using System.Linq;
using PptGcs.Core;

internal static class NativeFieldGeometryTests
{
    private const string A="aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa";
    private const string B="bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb";
    private const string C="cccccccc-cccc-cccc-cccc-cccccccccccc";
    private const string D="dddddddd-dddd-dddd-dddd-dddddddddddd";
    private static int passed,serial;
    private static void Main()
    {
        Run("native width pins only width and preserves native left/top",WidthOnly);
        Run("native height pins only height and preserves native left/top",HeightOnly);
        Run("native width lets an aspect ratio derive height without recentering",WidthRatio);
        Run("native height lets an aspect ratio derive width without recentering",HeightRatio);
        Run("native width moves a center-aligned follower while retaining driver left/top",WidthAlignment);
        Run("compatible native width and height drivers solve together",CompatibleNativeDrivers);
        Run("incompatible native dimension drivers reject without mutating their inputs",ConflictingNativeDrivers);
        Run("fixed native dimensions reject an incompatible typed value",FixedDimension);
        Run("a native value below linked minimum size is rejected",LinkedMinimum);
        Run("connected constraint discovery follows transitive enabled relations only",ConnectedClosure);
        Run("connected constraint discovery handles free drivers and validates inputs",ConnectedInputs);
        Run("native field edits leave unrelated unsatisfied geometry unchanged",UnrelatedComponent);
        Console.WriteLine("PASS: "+passed+" native field geometry tests");
    }
    private static void Run(string name,Action action){action();passed++;Console.WriteLine("PASS "+name);}
    private static void Assert(bool value,string message){if(!value)throw new Exception(message);}
    private static void Near(double actual,double expected){Assert(Math.Abs(actual-expected)<.000001,"Expected "+expected+", got "+actual);}
    private static void Throws<T>(Action action) where T:Exception
    {try{action();}catch(T){return;}throw new Exception("Expected "+typeof(T).Name);}
    private static ShapeGeometry Shape(string id,double x=10,double y=20,double w=80,double h=40){return new ShapeGeometry(id,x,y,w,h);}
    private static LayoutConstraint Relation(ConstraintKind kind,params string[] ids)
    {return new LayoutConstraint((++serial).ToString("x8")+"-0000-0000-0000-000000000001",256,kind,ids);}
    private static LayoutConstraint Value(ConstraintKind kind,double value,params string[] ids)
    {return new LayoutConstraint((++serial).ToString("x8")+"-0000-0000-0000-000000000001",256,kind,ids,value);}
    private static ShapeGeometry[] Solve(ShapeGeometry[] baseline,LayoutConstraint[] relations,params GeometryEdit[] edits)
    {return LayoutSolver.Solve(baseline,relations,edits).ToArray();}
    private static ShapeGeometry Get(ShapeGeometry[] shapes,string id){return shapes.Single(shape=>shape.Id==ShapeIdentity.Normalize(id));}
    private static void WidthOnly()
    {
        var before=new[]{Shape(A),Shape(B,200,300,80,55)};
        var after=Solve(before,new[]{Relation(ConstraintKind.SameWidth,A,B)},new GeometryEdit(Shape(A,17,25,120,999),GeometryEditMode.NativeWidth));
        var a=Get(after,A);var b=Get(after,B);
        Near(a.Width,120);Near(a.Height,40);Near(a.Left,17);Near(a.Top,25);
        Near(b.Width,120);Near(b.Height,55);Near(b.Left,200);Near(b.Top,300);
    }
    private static void HeightOnly()
    {
        var before=new[]{Shape(A),Shape(B,200,300,95,40)};
        var after=Solve(before,new[]{Relation(ConstraintKind.SameHeight,A,B)},new GeometryEdit(Shape(A,17,25,999,70),GeometryEditMode.NativeHeight));
        var a=Get(after,A);var b=Get(after,B);
        Near(a.Height,70);Near(a.Width,80);Near(a.Left,17);Near(a.Top,25);
        Near(b.Height,70);Near(b.Width,95);Near(b.Left,200);Near(b.Top,300);
    }
    private static void WidthRatio()
    {
        var after=Solve(new[]{Shape(A)},new[]{Value(ConstraintKind.AspectRatio,2,A)},new GeometryEdit(Shape(A,10,20,120,40),GeometryEditMode.NativeWidth));
        Near(after[0].Width,120);Near(after[0].Height,60);Near(after[0].Left,10);Near(after[0].Top,20);
    }
    private static void HeightRatio()
    {
        var after=Solve(new[]{Shape(A)},new[]{Value(ConstraintKind.AspectRatio,2,A)},new GeometryEdit(Shape(A,10,20,80,60),GeometryEditMode.NativeHeight));
        Near(after[0].Width,120);Near(after[0].Height,60);Near(after[0].Left,10);Near(after[0].Top,20);
    }
    private static void WidthAlignment()
    {
        var before=new[]{Shape(A),Shape(B,30,200,40,30)};
        var after=Solve(before,new[]{Relation(ConstraintKind.CenterXAligned,A,B)},new GeometryEdit(Shape(A,10,20,120,40),GeometryEditMode.NativeWidth));
        Near(Get(after,A).Left,10);Near(Get(after,A).Top,20);Near(Get(after,B).Left,50);Near(Get(after,B).Width,40);
    }
    private static void CompatibleNativeDrivers()
    {
        var before=new[]{Shape(A),Shape(B,200,300)};
        var relations=new[]{Relation(ConstraintKind.SameWidth,A,B),Relation(ConstraintKind.SameHeight,A,B),Value(ConstraintKind.AspectRatio,2,A)};
        var after=Solve(before,relations,new GeometryEdit(Shape(A,10,20,120,40),GeometryEditMode.NativeWidth),new GeometryEdit(Shape(B,200,300,80,60),GeometryEditMode.NativeHeight));
        foreach(var shape in after){Near(shape.Width,120);Near(shape.Height,60);}
        Near(Get(after,A).Left,10);Near(Get(after,B).Top,300);Near(LayoutSolver.MaxResidual(after,relations),0);
    }
    private static void ConflictingNativeDrivers()
    {
        var before=new[]{Shape(A),Shape(B,200,300)};
        var relations=new[]{Relation(ConstraintKind.SameWidth,A,B)};
        Throws<ConstraintConflictException>(()=>Solve(before,relations,new GeometryEdit(Shape(A,10,20,100),GeometryEditMode.NativeWidth),new GeometryEdit(Shape(B,200,300,110),GeometryEditMode.NativeWidth)));
        Near(before[0].Width,80);Near(before[1].Width,80);
    }
    private static void FixedDimension()
    {
        var before=new[]{Shape(A)};
        Throws<ConstraintConflictException>(()=>Solve(before,new[]{Value(ConstraintKind.FixedWidth,80,A)},new GeometryEdit(Shape(A,10,20,100),GeometryEditMode.NativeWidth)));
        Near(before[0].Width,80);
    }
    private static void LinkedMinimum()
    {
        Throws<ConstraintConflictException>(()=>Solve(new[]{Shape(A,10,20,80,8)},new[]{Value(ConstraintKind.AspectRatio,10,A)},new GeometryEdit(Shape(A,10,20,5,8),GeometryEditMode.NativeWidth)));
    }
    private static void ConnectedClosure()
    {
        var ab=Relation(ConstraintKind.SameY,A,B);var bc=Relation(ConstraintKind.SameWidth,B,C);var d=Value(ConstraintKind.FixedHeight,40,D);
        var disabled=new LayoutConstraint("ffffffff-ffff-ffff-ffff-ffffffffffff",256,ConstraintKind.SameX,new[]{C,D},enabled:false);
        var found=LayoutSolver.ConnectedConstraints(new[]{d,bc,disabled,ab},new[]{A.ToUpperInvariant(),A});
        Assert(found.Select(relation=>relation.Id).SequenceEqual(new[]{bc.Id,ab.Id}),"Enabled transitive relation membership and input order must be retained");
        Assert(!found.Contains(d)&&!found.Contains(disabled),"Unrelated or disabled relations must not be active");
    }
    private static void ConnectedInputs()
    {
        Assert(LayoutSolver.ConnectedConstraints(new[]{Relation(ConstraintKind.SameY,A,B)},new[]{C}).Count==0,"A free driver has no connected relations");
        Assert(LayoutSolver.ConnectedConstraints(new[]{Relation(ConstraintKind.SameY,A,B)},new string[0]).Count==0,"No drivers means no edit component");
        Throws<ArgumentNullException>(()=>LayoutSolver.ConnectedConstraints(null,new[]{A}));
        Throws<ArgumentNullException>(()=>LayoutSolver.ConnectedConstraints(new LayoutConstraint[0],null));
        Throws<ArgumentException>(()=>LayoutSolver.ConnectedConstraints(new LayoutConstraint[]{null},new[]{A}));
    }
    private static void UnrelatedComponent()
    {
        var before=new[]{Shape(A),Shape(B,200,20),Shape(C,300,100),Shape(D,400,120)};
        var relations=new[]{Relation(ConstraintKind.SameWidth,A,B),Relation(ConstraintKind.SameY,C,D)};
        var after=Solve(before,relations,new GeometryEdit(Shape(A,10,20,120),GeometryEditMode.NativeWidth));
        Near(Get(after,A).Width,120);Near(Get(after,B).Width,120);Near(Get(after,C).Top,100);Near(Get(after,D).Top,120);
        Near(LayoutSolver.MaxResidual(after,LayoutSolver.ConnectedConstraints(relations,new[]{A})),0);
    }
}
