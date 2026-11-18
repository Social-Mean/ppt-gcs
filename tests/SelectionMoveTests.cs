using System;
using System.Drawing;
using System.Linq;
using PptGcs.AddIn;
using PptGcs.Core;

internal static class SelectionMoveTests
{
    private const string A="aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa";
    private const string B="bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb";
    private const string C="cccccccc-cccc-cccc-cccc-cccccccccccc";
    private static int passed;
    private static readonly RectangleF Bounds=new RectangleF(100,100,200,100);

    private static void Main()
    {
        Run("a selected constrained object and a zero-width line share the same absolute delta", ZeroWidthExtra);
        Run("all selected constrained members are drivers while an unselected follower moves only as required", SelectedDrivers);
        Run("later pointer samples use the frozen initial selection rather than accumulated deltas", AbsoluteDelta);
        Run("conflicting selection translation is rejected before any native write", RejectBeforeWrite);
        Run("a partially applied native frame restores followers and all extras together", AtomicWriteFailure);
        Run("native readback mismatch restores the prior complete frame", AtomicReadbackFailure);
        Run("constraint residual is checked even when individual shape readbacks are within tolerance", ResidualReadbackFailure);
        Run("an unrelated unsatisfied component does not block a valid selected translation", UnrelatedResidual);
        Run("a rejected later frame restores the last accepted frame rather than the initial state", LastAcceptedFrame);
        Run("failed rollback is reported as a hard failure with both causes", RollbackFailure);
        Run("frame equality uses both identities and complete extra bounds", FrameMatching);
        Run("selection plans reject incomplete or duplicate members", InvalidPlans);
        Run("a true selected-member hit owns translation", RealSelectedHit);
        Run("a foreign hit is never claimed even on the selected outline", ForeignHit);
        Run("a direct hit on an unselected object continues into the single-object route", ForeignRoute);
        Run("a covering unselected object cannot translate the old selected outline", ForeignOutlineRoute);
        Run("selection handles and blank gaps remain intentionally native", NativeRoute);
        Run("a blank selected outline has a limited move entry", BlankOutline);
        Run("blank canvas in a gap between selected members remains native", BlankInterior);
        Run("all eight aggregate resize handles stay native", NativeHandles);
        Run("a degenerate line keeps endpoint handles but retains its movable interior", LineHandles);
        Run("outline and handle tolerances follow display DPI", DpiRouting);
        Run("exact native group ownership does not depend on a guessed COM box", NativeHitOutsideBox);
        Console.WriteLine("PASS: "+passed+" selection move tests");
    }
    private static ShapeGeometry Shape(string id,double x,double y,double w=30,double h=40)
    {return new ShapeGeometry(id,x,y,w,h);}
    private static SelectionMovePlan Plan(string[] drivers=null,SelectionMovePosition[] extras=null)
    {
        return new SelectionMovePlan(new[]{Shape(A,10,20),Shape(B,100,20),Shape(C,200,20)},
            new[]{new LayoutConstraint("11111111-1111-1111-1111-111111111111",256,ConstraintKind.SameY,new[]{A,B,C})},
            drivers??new[]{A},extras??new[]{new SelectionMovePosition(90,300,400,0,80)});
    }
    private static void Assert(bool value,string message){if(!value)throw new Exception(message);}
    private static void Near(double actual,double expected){Assert(Math.Abs(actual-expected)<.000001,"Expected "+expected+", got "+actual);}
    private static void Throws<T>(Action action) where T:Exception
    {try{action();}catch(T){return;}throw new Exception("Expected "+typeof(T).Name);}
    private static void Run(string name,Action action){action();passed++;Console.WriteLine("PASS "+name);}
    private static ShapeGeometry Get(SelectionMoveFrame frame,string id){return frame.Geometry.Single(shape=>shape.Id==ShapeIdentity.Normalize(id));}
    private static void ZeroWidthExtra()
    {
        var plan=Plan();var frame=plan.Solve(12.5,-8.25);var line=frame.Extras.Single();
        Near(Get(frame,A).Left,22.5);Near(Get(frame,A).Top,11.75);Near(Get(frame,B).Top,11.75);Near(Get(frame,C).Top,11.75);
        Near(line.Left,312.5);Near(line.Top,391.75);Near(line.Width,0);Near(line.Height,80);
        Near(plan.Initial.Extras.Single().Left,300);Near(Get(plan.Initial,A).Left,10);
    }
    private static void SelectedDrivers()
    {
        var plan=Plan(new[]{A,B},new SelectionMovePosition[0]);var frame=plan.Solve(8,9);
        Near(Get(frame,A).Left,18);Near(Get(frame,B).Left,108);Near(Get(frame,C).Left,200);
        foreach(var shape in frame.Geometry){Near(shape.Top,29);Near(shape.Width,30);Near(shape.Height,40);}
    }
    private static void AbsoluteDelta()
    {
        var plan=Plan();var first=plan.Solve(10,20);var second=plan.Solve(12,24);
        Near(Get(first,A).Left,20);Near(Get(second,A).Left,22);
        Near(first.Extras.Single().Left,310);Near(second.Extras.Single().Left,312);
    }
    private static void RejectBeforeWrite()
    {
        var plan=new SelectionMovePlan(new[]{Shape(A,10,20)},
            new[]{new LayoutConstraint("22222222-2222-2222-2222-222222222222",256,ConstraintKind.FixedPosition,new[]{A},10,20)},
            new[]{A},new[]{new SelectionMovePosition(90,300,400,0,80)});
        int writes=0;
        Throws<ConstraintConflictException>(()=>{var target=plan.Solve(5,10);plan.ApplyAtomic(plan.Initial,target,frame=>writes++,()=>plan.Initial);});
        Assert(writes==0,"An incompatible edit must not start writing selected members");
    }
    private static void AtomicWriteFailure()
    {
        var plan=Plan(extras:new[]{new SelectionMovePosition(90,300,400,0,80),new SelectionMovePosition(91,500,600,100,0)});
        var target=plan.Solve(15,25);var current=plan.Initial;int writes=0;
        Throws<SelectionMoveRejectedException>(()=>plan.ApplyAtomic(plan.Initial,target,frame=>{
            writes++;
            if(writes==1){current=new SelectionMoveFrame(frame.Geometry,new[]{frame.Extras[0],plan.Initial.Extras[1]});throw new InvalidOperationException("native setter rejected second extra");}
            current=frame;
        },()=>current));
        Assert(writes==2&&current.Matches(plan.Initial),"Followers and both line extras must be restored together");
    }
    private static void AtomicReadbackFailure()
    {
        var plan=Plan();var current=plan.Initial;var target=plan.Solve(5,7);int writes=0;
        Throws<SelectionMoveRejectedException>(()=>plan.ApplyAtomic(plan.Initial,target,frame=>{
            writes++;
            current=writes==1 ? new SelectionMoveFrame(frame.Geometry,new[]{new SelectionMovePosition(90,306,407,0,80)}) : frame;
        },()=>current));
        Assert(current.Matches(plan.Initial)&&writes==2,"A silent one-point native drift must not be accepted");
    }
    private static void ResidualReadbackFailure()
    {
        var plan=Plan();var current=plan.Initial;var target=plan.Solve(5,7);int writes=0;
        Throws<SelectionMoveRejectedException>(()=>plan.ApplyAtomic(plan.Initial,target,frame=>{
            writes++;
            current=writes==1 ? new SelectionMoveFrame(frame.Geometry.Select(shape=>new ShapeGeometry(shape.Id,shape.Left,
                shape.Top+(shape.Id==ShapeIdentity.Normalize(A) ? .009 : -.009),shape.Width,shape.Height)),frame.Extras) : frame;
        },()=>current));
        Assert(current.Matches(plan.Initial)&&writes==2,"Residual validation must preserve the previous complete frame");
    }
    private static void LastAcceptedFrame()
    {
        var plan=Plan();var current=plan.Initial;
        var accepted=plan.ApplyAtomic(current,plan.Solve(5,7),frame=>current=frame,()=>current);
        int writes=0;
        Throws<SelectionMoveRejectedException>(()=>plan.ApplyAtomic(accepted,plan.Solve(20,30),frame=>{
            writes++;current=frame;if(writes==1)throw new InvalidOperationException("native setter failure");
        },()=>current));
        Assert(current.Matches(accepted)&&!current.Matches(plan.Initial),"Rejected input must preserve the latest accepted pointer position");
    }
    private static void UnrelatedResidual()
    {
        string d="dddddddd-dddd-dddd-dddd-dddddddddddd";
        var initial=new[]{Shape(A,10,20),Shape(B,100,20),Shape(C,200,100),Shape(d,300,120)};
        var relations=new[]{
            new LayoutConstraint("33333333-3333-3333-3333-333333333333",256,ConstraintKind.SameY,new[]{A,B}),
            new LayoutConstraint("44444444-4444-4444-4444-444444444444",256,ConstraintKind.SameY,new[]{C,d})};
        var plan=new SelectionMovePlan(initial,relations,new[]{A},new[]{new SelectionMovePosition(90,400,400,0,50)});
        var requested=plan.Solve(.5,2);var current=plan.Initial;int writes=0;
        var actual=plan.ApplyAtomic(current,requested,frame=>{current=frame;writes++;},()=>current);
        Assert(writes==1 && actual.Matches(requested),"A valid component must be committed without rollback");
        Near(Get(actual,A).Top,22);Near(Get(actual,B).Top,22);
        Near(Get(actual,C).Top,100);Near(Get(actual,d).Top,120);
        Near(LayoutSolver.MaxResidual(actual.Geometry,relations),20);
        Near(LayoutSolver.MaxResidual(actual.Geometry,LayoutSolver.ConnectedConstraints(relations,new[]{A})),0);
    }
    private static void RollbackFailure()
    {
        var plan=Plan();Exception failure=null;
        try{plan.ApplyAtomic(plan.Initial,plan.Solve(5,7),frame=>{throw new InvalidOperationException("unavailable native shape");},()=>plan.Initial);}
        catch(Exception error){failure=error;}
        Assert(failure is InvalidOperationException && !(failure is SelectionMoveRejectedException) &&
            failure.InnerException is AggregateException && ((AggregateException)failure.InnerException).InnerExceptions.Count==2,
            "A failed rollback cannot be reported as an ordinary restored rejection");
    }
    private static void FrameMatching()
    {
        var plan=Plan();var initial=plan.Initial;
        Assert(initial.Matches(new SelectionMoveFrame(initial.Geometry.Reverse(),initial.Extras.Reverse())),"Order is not geometry identity");
        Assert(!initial.Matches(new SelectionMoveFrame(initial.Geometry,new[]{new SelectionMovePosition(90,300,400,1,80)})),"A zero-width line must not silently acquire width");
        Assert(!initial.Matches(new SelectionMoveFrame(initial.Geometry,new[]{new SelectionMovePosition(91,300,400,0,80)})),"A native extra cannot be substituted by another ID");
        Throws<ArgumentException>(()=>new SelectionMoveFrame(initial.Geometry,new[]{initial.Extras[0],initial.Extras[0]}));
    }
    private static void InvalidPlans()
    {
        Throws<ArgumentException>(()=>Plan(new string[0]));
        Throws<ArgumentException>(()=>Plan(new[]{A,A}));
        Throws<ArgumentException>(()=>Plan(new[]{"dddddddd-dddd-dddd-dddd-dddddddddddd"}));
        Near(Get(Plan(new[]{A},new SelectionMovePosition[0]).Solve(4,6),A).Left,14);
        Throws<ArgumentOutOfRangeException>(()=>new SelectionMovePosition(90,0,0,-1,10));
        Throws<ArgumentOutOfRangeException>(()=>Plan().Solve(Double.PositiveInfinity,0));
    }
    private static void RealSelectedHit(){Assert(MultiSelectionPointerPolicy.CanMove(Bounds,new PointF(160,130),8,6,true),"Native selected-member body hit must translate the whole selection");}
    private static void ForeignHit(){Assert(!MultiSelectionPointerPolicy.CanMove(Bounds,new PointF(97,130),8,6,false),"A covering foreign object wins over the selected outline");}
    private static void ForeignRoute()
    {
        Assert(MultiSelectionPointerPolicy.RouteMove(Bounds,new PointF(400,230),8,6,false)==null,
            "Pressing a real unselected C must proceed to C's own constraint route, even when A/B are selected");
        Assert(MultiSelectionPointerPolicy.RouteMove(Bounds,new PointF(160,130),8,6,true)==true,
            "Pressing a selected member must still move the full original selection");
    }
    private static void ForeignOutlineRoute()
    {
        Assert(MultiSelectionPointerPolicy.RouteMove(Bounds,new PointF(97,130),8,6,false)==null,
            "A foreign covering object must route independently rather than translating the old selection through its outline band");
        Assert(MultiSelectionPointerPolicy.RouteMove(Bounds,new PointF(97,130),8,6,null)==true,
            "The same outline point may move the selection only when native hit testing proves it is blank");
    }
    private static void NativeRoute()
    {
        Assert(MultiSelectionPointerPolicy.RouteMove(Bounds,new PointF(160,130),8,6,null)==false,
            "A blank gap must remain native rather than becoming a single-object hit");
        foreach(bool? native in new bool?[]{true,false,null})
            Assert(MultiSelectionPointerPolicy.RouteMove(Bounds,new PointF(100,100),8,6,native)==false,
                "An aggregate handle must retain native resize routing regardless of the underlying hit result");
    }
    private static void BlankOutline()
    {
        Assert(MultiSelectionPointerPolicy.CanMove(Bounds,new PointF(94,130),8,6,null),"The exact outline tolerance is available");
        Assert(!MultiSelectionPointerPolicy.CanMove(Bounds,new PointF(93,130),8,6,null),"Blank canvas beyond the outline is not claimed");
    }
    private static void BlankInterior(){Assert(!MultiSelectionPointerPolicy.CanMove(Bounds,new PointF(160,130),8,6,null),"An empty group gap is not a guessed selected-member hit");}
    private static void NativeHandles()
    {
        foreach(var point in new[]{new PointF(100,100),new PointF(200,100),new PointF(300,100),new PointF(300,150),new PointF(300,200),new PointF(200,200),new PointF(100,200),new PointF(100,150)})
            Assert(!MultiSelectionPointerPolicy.CanMove(Bounds,point,8,6,true),"Resize handle became a selection move: "+point);
    }
    private static void LineHandles()
    {
        var line=new RectangleF(100,100,0,80);
        Assert(MultiSelectionPointerPolicy.HitHandle(line,new PointF(100,100),8)&&MultiSelectionPointerPolicy.HitHandle(line,new PointF(100,180),8),"Line endpoint handles remain native");
        Assert(!MultiSelectionPointerPolicy.HitHandle(line,new PointF(100,140),8)&&MultiSelectionPointerPolicy.CanMove(line,new PointF(100,140),8,6,true),"A real line interior hit remains available to move the full selection");
    }
    private static void DpiRouting()
    {
        foreach(float scale in new[]{1.25f,1.5f,2f})
        {
            var bounds=new RectangleF(100*scale,100*scale,200*scale,100*scale);
            Assert(MultiSelectionPointerPolicy.CanMove(bounds,new PointF(94*scale,130*scale),8*scale,6*scale,null),"Scaled selected outline did not own its band");
            Assert(!MultiSelectionPointerPolicy.CanMove(bounds,new PointF(100*scale,100*scale),8*scale,6*scale,true),"Scaled corner handle became a move");
        }
    }
    private static void NativeHitOutsideBox(){Assert(MultiSelectionPointerPolicy.CanMove(Bounds,new PointF(400,230),8,6,true),"A verified native selected-group hit must not be lost to an approximate box");}
}
