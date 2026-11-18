using System;
using System.Collections.Generic;
using System.Linq;
using PptGcs.Core;

internal static class MultiEditTests
{
    private const string A="aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa";
    private const string B="bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb";
    private const string C="cccccccc-cccc-cccc-cccc-cccccccccccc";
    private const string D="dddddddd-dddd-dddd-dddd-dddddddddddd";
    private const string E="eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee";
    private static int passed, serial;

    private static void Main()
    {
        Run("multi-selection move pins every selected object and preserves all selected sizes", SelectionMove);
        Run("ordered two-axis offsets accept a consistent selection translation", OffsetMove);
        Run("two end drivers solve the middle equal gap from the actual unequal edges", EqualGapEnds);
        Run("negative gaps retain the stored member order", ReversedGapEnds);
        Run("two selected members preserve a fixed-gap chain", FixedGapSelection);
        Run("native exact edits pin all fields and propagate compatible proportional sizes", ExactDimensions);
        Run("conflicting selected positions reject the entire edit without mutation", PositionConflict);
        Run("conflicting selected dimensions reject the entire edit without mutation", DimensionConflict);
        Run("a fixed object rejects a translated selection atomically", FixedDriver);
        Run("multiple connected components solve together while unrelated geometry stays untouched", DisconnectedComponents);
        Run("unconstrained multi-edit reaches both exact targets", FreeExact);
        Run("all-driver equations accept compatible targets without a free variable", AllDriven);
        Run("an unrelated second driver does not change the minimum-movement follower solution", MinimumMovement);
        Run("mixed edge resize and exact edits retain the resize's opposite-edge anchor", MixedModes);
        Run("independent corner ratios retain pointer projection and opposite corners", CornerModes);
        Run("invalid and duplicate edit descriptions are rejected", InvalidEdits);
        Run("disabled missing-object relations do not block multiple unrelated edits", DisabledRelations);
        Run("large-coordinate multi-selection edits preserve edge gap formulas", LargeCoordinates);
        Run("repeated multi-selection edits do not accumulate coordinate drift", RepeatedEdits);
        Console.WriteLine("PASS: "+passed+" multi-edit tests");
    }

    private static ShapeGeometry Shape(string id,double x=0,double y=0,double w=20,double h=30)
    { return new ShapeGeometry(id,x,y,w,h); }
    private static LayoutConstraint Relation(ConstraintKind kind,params string[] ids)
    { return new LayoutConstraint((++serial).ToString("x8")+"-0000-0000-0000-000000000001",256,kind,ids); }
    private static LayoutConstraint Parameter(ConstraintKind kind,double value,double? second,params string[] ids)
    { return new LayoutConstraint((++serial).ToString("x8")+"-0000-0000-0000-000000000001",256,kind,ids,value,second); }
    private static GeometryEdit Move(ShapeGeometry target)
    { return new GeometryEdit(target,GeometryEditMode.Move); }
    private static Dictionary<string,ShapeGeometry> Solve(IEnumerable<ShapeGeometry> shapes,IEnumerable<LayoutConstraint> relations,params GeometryEdit[] edits)
    { return LayoutSolver.Solve(shapes,relations,edits).ToDictionary(shape=>shape.Id); }
    private static ShapeGeometry Get(IDictionary<string,ShapeGeometry> result,string id)
    { return result[ShapeIdentity.Normalize(id)]; }
    private static void Near(double value,double expected,double tolerance=0.000001)
    { if(Double.IsNaN(value)||Math.Abs(value-expected)>tolerance)throw new Exception("Expected "+expected+", got "+value); }
    private static void Same(ShapeGeometry actual,ShapeGeometry expected)
    { if(actual.Id!=expected.Id)throw new Exception("Identity mismatch");Near(actual.Left,expected.Left);Near(actual.Top,expected.Top);Near(actual.Width,expected.Width);Near(actual.Height,expected.Height); }
    private static void Assert(bool value,string message)
    { if(!value)throw new Exception(message); }
    private static void Throws<T>(Action action) where T:Exception
    { try{action();}catch(T){return;}throw new Exception("Expected "+typeof(T).Name); }
    private static void Run(string name,Action test)
    { test();passed++;Console.WriteLine("PASS "+name); }

    private static void SelectionMove()
    {
        var before=new[]{Shape(A,0,10,40,20),Shape(B,100,10,40,30),Shape(C,220,10,40,50),Shape(D,400,60,70,80)};
        var relations=new[]{Relation(ConstraintKind.SameY,A,B,C),Relation(ConstraintKind.SameWidth,A,B,C)};
        var result=Solve(before,relations,Move(Shape(A,17,-13,900,901)),Move(Shape(B,117,-13,902,903)));
        Same(Get(result,A),Shape(A,17,-13,40,20));Same(Get(result,B),Shape(B,117,-13,40,30));
        Same(Get(result,C),Shape(C,220,-13,40,50));Same(Get(result,D),before[3]);
        Near(LayoutSolver.MaxResidual(result.Values,relations),0);
    }

    private static void OffsetMove()
    {
        var before=new[]{Shape(A,10,20),Shape(B,110,50),Shape(C,160,30)};
        var relations=new[]{Parameter(ConstraintKind.Offset,100,30,A,B),Parameter(ConstraintKind.Offset,50,-20,B,C)};
        var result=Solve(before,relations,Move(Shape(A,25,12)),Move(Shape(B,125,42)));
        Same(Get(result,A),Shape(A,25,12));Same(Get(result,B),Shape(B,125,42));Same(Get(result,C),Shape(C,175,22));
        Near(LayoutSolver.MaxResidual(result.Values,relations),0);
    }

    private static ShapeGeometry[] GapRow(double shift=0)
    { return new[]{Shape(A,shift,10,20,30),Shape(B,shift+60,15,40,35),Shape(C,shift+140,20,30,40)}; }
    private static void EqualGapEnds()
    {
        var relations=new[]{Relation(ConstraintKind.EqualHorizontalGap,A,B,C)};
        var result=Solve(GapRow(),relations,Move(Shape(A,20,10,20,30)),Move(Shape(C,200,20,30,40)));
        Near(Get(result,B).Left,100);Near(Get(result,B).Top,15);
        Near(Get(result,B).Left-Get(result,A).Left-Get(result,A).Width,60);
        Near(Get(result,C).Left-Get(result,B).Left-Get(result,B).Width,60);
        Near(Get(result,A).Left,20);Near(Get(result,C).Left,200);
    }
    private static void ReversedGapEnds()
    {
        var relations=new[]{Relation(ConstraintKind.EqualHorizontalGap,A,B,C)};
        var result=Solve(GapRow(),relations,Move(Shape(A,210,10,20,30)),Move(Shape(C,10,20,30,40)));
        Near(Get(result,B).Left,100);
        Near(Get(result,B).Left-Get(result,A).Left-Get(result,A).Width,-130);
        Near(Get(result,C).Left-Get(result,B).Left-Get(result,B).Width,-130);
    }
    private static void FixedGapSelection()
    {
        var relations=new[]{Parameter(ConstraintKind.FixedHorizontalGap,40,null,A,B,C)};
        var result=Solve(GapRow(),relations,Move(Shape(A,-12,13,20,30)),Move(Shape(C,128,23,30,40)));
        Near(Get(result,B).Left,48);Near(Get(result,B).Top,15);
        Near(LayoutSolver.MaxResidual(result.Values,relations),0);
    }
    private static void ExactDimensions()
    {
        var before=new[]{Shape(A,0,0,20,30),Shape(B,100,100,20,10),Shape(C,200,200,20,60)};
        var relations=new[]{Relation(ConstraintKind.SameWidth,A,B,C),Parameter(ConstraintKind.AspectRatio,2,null,B)};
        var a=Shape(A,7,9,80,45);var b=Shape(B,210,90,80,40);
        var result=Solve(before,relations,new GeometryEdit(a),new GeometryEdit(b));
        Same(Get(result,A),a);Same(Get(result,B),b);Same(Get(result,C),Shape(C,200,200,80,60));
        Near(LayoutSolver.MaxResidual(result.Values,relations),0);
    }
    private static void PositionConflict()
    {
        var before=new[]{Shape(A,0,10),Shape(B,100,10),Shape(C,200,10)};
        var relations=new[]{Relation(ConstraintKind.SameY,A,B,C)};
        Throws<ConstraintConflictException>(()=>Solve(before,relations,Move(Shape(A,0,20)),Move(Shape(B,100,30))));
        Same(before[0],Shape(A,0,10));Same(before[1],Shape(B,100,10));Same(before[2],Shape(C,200,10));
    }
    private static void DimensionConflict()
    {
        var before=new[]{Shape(A),Shape(B,100),Shape(C,200)};
        var relations=new[]{Relation(ConstraintKind.SameWidth,A,B,C)};
        Throws<ConstraintConflictException>(()=>Solve(before,relations,new GeometryEdit(Shape(A,0,0,50)),new GeometryEdit(Shape(B,100,0,60))));
        Same(before[0],Shape(A));Same(before[1],Shape(B,100));Same(before[2],Shape(C,200));
    }
    private static void FixedDriver()
    {
        var before=new[]{Shape(A,10,20),Shape(B,100,20)};
        var relations=new[]{Relation(ConstraintKind.SameY,A,B),Parameter(ConstraintKind.FixedPosition,10,20,A)};
        Throws<ConstraintConflictException>(()=>Solve(before,relations,Move(Shape(A,15,25)),Move(Shape(B,105,25))));
        Same(before[0],Shape(A,10,20));Same(before[1],Shape(B,100,20));
    }
    private static void DisconnectedComponents()
    {
        var before=new[]{Shape(A,0,10),Shape(B,70,10),Shape(C,200,100),Shape(D,200,300),Shape(E,500,500,33,44)};
        var relations=new[]{Relation(ConstraintKind.SameY,A,B),Relation(ConstraintKind.SameX,C,D),Parameter(ConstraintKind.FixedWidth,20,null,E)};
        var result=Solve(before,relations,Move(Shape(A,5,25)),Move(Shape(C,240,120)));
        Same(Get(result,A),Shape(A,5,25));Same(Get(result,B),Shape(B,70,25));
        Same(Get(result,C),Shape(C,240,120));Same(Get(result,D),Shape(D,240,300));
        Same(Get(result,E),before[4]);
        Assert(ReferenceEquals(Get(result,E),before[4]),"Disconnected geometry should not be reconstructed");
    }
    private static void FreeExact()
    {
        var before=new[]{Shape(A),Shape(B,100),Shape(C,200)};
        var a=Shape(A,-40,50,70,80);var b=Shape(B,160,200,90,100);
        var result=Solve(before,new LayoutConstraint[0],new GeometryEdit(a),new GeometryEdit(b));
        Same(Get(result,A),a);Same(Get(result,B),b);Same(Get(result,C),before[2]);
    }
    private static void AllDriven()
    {
        var before=new[]{Shape(A),Shape(B,60,10,40,20)};
        var relations=new[]{Relation(ConstraintKind.CentersCoincident,A,B)};
        var a=Shape(A,10,20,80,40);var b=Shape(B,30,30,40,20);
        var result=Solve(before,relations,new GeometryEdit(a),new GeometryEdit(b));
        Same(Get(result,A),a);Same(Get(result,B),b);Near(LayoutSolver.MaxResidual(result.Values,relations),0);
        Throws<ConstraintConflictException>(()=>Solve(before,relations,new GeometryEdit(a),new GeometryEdit(Shape(B,31,30,40,20))));
    }
    private static void MinimumMovement()
    {
        var before=GapRow().Concat(new[]{Shape(D,400,400)}).ToArray();
        var relations=new[]{Relation(ConstraintKind.EqualHorizontalGap,A,B,C)};
        var result=Solve(before,relations,Move(Shape(A,20,10,20,30)),Move(Shape(D,450,470)));
        // Minimize db^2 + dc^2 subject to dc - 2*db = -20.
        Near(Get(result,B).Left,68);Near(Get(result,C).Left,136);Same(Get(result,D),Shape(D,450,470));
        Near(LayoutSolver.MaxResidual(result.Values,relations),0);
    }
    private static void MixedModes()
    {
        var before=new[]{Shape(A,10,30,80,40),Shape(B,200,100,80,60)};
        var relations=new[]{Relation(ConstraintKind.SameWidth,A,B),Parameter(ConstraintKind.AspectRatio,2,null,A)};
        var result=Solve(before,relations,new GeometryEdit(Shape(A,-10,30,100,40),GeometryEditMode.ResizeW),new GeometryEdit(Shape(B,210,110,100,60)));
        Same(Get(result,A),Shape(A,-10,25,100,50));Same(Get(result,B),Shape(B,210,110,100,60));
        Near(LayoutSolver.MaxResidual(result.Values,relations),0);
    }
    private static void CornerModes()
    {
        var before=new[]{Shape(A,10,20,100,50),Shape(B,200,300,40,80)};
        var relations=new[]{Parameter(ConstraintKind.AspectRatio,2,null,A),Parameter(ConstraintKind.AspectRatio,.5,null,B)};
        var result=Solve(before,relations,new GeometryEdit(Shape(A,0,0,140,100),GeometryEditMode.ResizeNW),new GeometryEdit(Shape(B,0,0,100,140),GeometryEditMode.ResizeSE));
        Same(Get(result,A),Shape(A,-42,-6,152,76));Same(Get(result,B),Shape(B,200,300,76,152));
    }
    private static void InvalidEdits()
    {
        var before=new[]{Shape(A),Shape(B,100)};var none=new LayoutConstraint[0];
        Throws<ArgumentNullException>(()=>new GeometryEdit(null));
        Throws<ArgumentOutOfRangeException>(()=>new GeometryEdit(Shape(A),(GeometryEditMode)999));
        Throws<ArgumentNullException>(()=>LayoutSolver.Solve(before,none,(IEnumerable<GeometryEdit>)null));
        Throws<ArgumentException>(()=>LayoutSolver.Solve(before,none,new GeometryEdit[0]));
        Throws<ArgumentException>(()=>LayoutSolver.Solve(before,none,new GeometryEdit[]{null}));
        Throws<ArgumentException>(()=>Solve(before,none,new GeometryEdit(Shape(A)),new GeometryEdit(Shape(A.ToUpperInvariant()))));
        Throws<ArgumentException>(()=>Solve(before,none,new GeometryEdit(Shape(A)),new GeometryEdit(Shape(C))));
        Throws<ArgumentException>(()=>Solve(new[]{Shape(A),Shape(A)},none,new GeometryEdit(Shape(A))));
    }
    private static void DisabledRelations()
    {
        var relation=new LayoutConstraint("ffffffff-ffff-ffff-ffff-ffffffffffff",256,ConstraintKind.SameY,new[]{A,C},enabled:false);
        var result=Solve(new[]{Shape(A),Shape(B,100)},new[]{relation},Move(Shape(A,10,20)),Move(Shape(B,110,30)));
        Same(Get(result,A),Shape(A,10,20));Same(Get(result,B),Shape(B,110,30));
    }
    private static void LargeCoordinates()
    {
        const double shift=100000000;
        var relations=new[]{Relation(ConstraintKind.EqualHorizontalGap,A,B,C)};
        var result=Solve(GapRow(shift),relations,Move(Shape(A,shift+20,10,20,30)),Move(Shape(C,shift+200,20,30,40)));
        Near(Get(result,A).Left-shift,20);Near(Get(result,B).Left-shift,100);Near(Get(result,C).Left-shift,200);
        Near(LayoutSolver.MaxResidual(result.Values,relations),0);
    }
    private static void RepeatedEdits()
    {
        var relations=new[]{Relation(ConstraintKind.SameY,A,B,C),Relation(ConstraintKind.SameWidth,A,B,C)};
        IReadOnlyList<ShapeGeometry> state=new[]{Shape(A,10,20),Shape(B,60,20),Shape(C,120,20)};
        for(int i=0;i<100;i++)
        {
            double dy=i%2==0 ? 17.125 : 20;
            state=LayoutSolver.Solve(state,relations,new[]{Move(Shape(A,10,dy)),Move(Shape(B,60,dy))});
            Near(LayoutSolver.MaxResidual(state,relations),0);
        }
        Same(state.Single(shape=>shape.Id==ShapeIdentity.Normalize(A)),Shape(A,10,20));
        Same(state.Single(shape=>shape.Id==ShapeIdentity.Normalize(B)),Shape(B,60,20));
        Same(state.Single(shape=>shape.Id==ShapeIdentity.Normalize(C)),Shape(C,120,20));
    }
}
