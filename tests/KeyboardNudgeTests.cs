using System;
using System.Linq;
using PptGcs.AddIn;
using PptGcs.Core;

internal static class KeyboardNudgeTests
{
    private const string A="aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa";
    private const string B="bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb";
    private const string C="cccccccc-cccc-cccc-cccc-cccccccccccc";
    private static int passed;

    private static void Main()
    {
        Run("shape and exact group-member selections on the active canvas own arrow input",OwnShapeInput);
        Run("text caret and non-shape selections retain native navigation",NativeEditingInput);
        Run("Shift resize and Alt system navigation are never nudges",NativeModifiers);
        Run("Ribbon fields, notes and background windows retain their arrows",NativeFields);
        Run("the long coordinate baseline preserves one-pixel precision at 115 percent zoom",NativePixelScale);
        Run("unsnapped arrows translate by one screen pixel on either axis",PixelStep);
        Run("snap-to-grid uses the current presentation spacing",GridStep);
        Run("Ctrl overrides snap-to-grid with one screen pixel",ControlStep);
        Run("one-pixel translation follows zoom rather than a hardcoded point",ZoomScale);
        Run("all four directions have independent horizontal and vertical scale",Directions);
        Run("invalid native scale, grid distance and direction are rejected",InvalidNativeInput);
        Run("single-shape nudge drives constrained followers without changing dimensions",SingleDriver);
        Run("multiple selected drivers and a zero-width extra translate rigidly",MultipleDrivers);
        Run("held-key repeats use the frozen baseline without compounding float-rounding drift",RepeatedNudge);
        Run("fixed-position conflicts reject the complete nudge before extras move",RejectedNudge);
        Run("a rejected native write restores the last complete keyboard frame",AtomicNudge);
        Run("held arrow steps move the actual child and drive native parent bounds without parent writes",GroupChildNudge);
        Console.WriteLine("PASS: "+passed+" keyboard nudge tests");
    }
    private static void Run(string name,Action action){action();passed++;Console.WriteLine("PASS "+name);}
    private static void Assert(bool condition,string message){if(!condition)throw new Exception(message);}
    private static void Near(double actual,double expected,double tolerance=.000001)
    {Assert(Math.Abs(actual-expected)<=tolerance,"Expected "+expected+", got "+actual);}
    private static void Throws<T>(Action action) where T:Exception
    {try{action();}catch(T){return;}throw new Exception("Expected "+typeof(T).Name);}
    private static ShapeGeometry Shape(string id,double x,double y){return new ShapeGeometry(id,x,y,50,40);}
    private static LayoutConstraint SameY(){return new LayoutConstraint("11111111-1111-1111-1111-111111111111",256,ConstraintKind.SameY,new[]{A,B,C});}
    private static SelectionMovePlan Plan(string[] drivers=null,SelectionMovePosition[] extras=null)
    {return new SelectionMovePlan(new[]{Shape(A,10,20),Shape(B,100,20),Shape(C,200,20)},new[]{SameY()},drivers??new[]{A},extras??new SelectionMovePosition[0]);}
    private static ShapeGeometry Get(SelectionMoveFrame frame,string id){return frame.Geometry.Single(shape=>shape.Id==ShapeIdentity.Normalize(id));}
    private static bool Own(int key=0x28,int type=2,bool shift=false,bool alt=false,bool canvas=true,bool active=true)
    {return KeyboardNudgePolicy.CanOwn(key,type,shift,alt,canvas,active);}

    private static void OwnShapeInput()
    {
        foreach(int key in new[]{0x25,0x26,0x27,0x28})Assert(Own(key),"Shape arrow must be owned");
        foreach(int key in new[]{0x09,0x0D,0x1B,0x41})Assert(!Own(key),"A non-arrow must remain native");
    }
    private static void NativeEditingInput(){Assert(!Own(type:3)&&!Own(type:0)&&!Own(type:1),"Caret and non-shape selections must stay native");}
    private static void NativeModifiers(){Assert(!Own(shift:true)&&!Own(alt:true)&&!Own(shift:true,alt:true),"Resize/system modifiers must not start a move");}
    private static void NativeFields(){Assert(!Own(canvas:false)&&!Own(active:false),"Native text fields and inactive canvases retain their arrows");}
    private static void NativePixelScale()
    {
        double scale=KeyboardNudgePolicy.PixelsPerPoint(401,15678);
        Near(scale,1.5277);
        Near(KeyboardNudgePolicy.StepPoints(false,5.669921875,false,scale),.6545787785559992,1e-12);
        // The 100-point baseline would round this transform to 1.53 pixels/pt.
        Assert(Math.Abs(1.53-scale)>.002,"The calibration must retain the long-baseline precision");
    }
    private static void PixelStep(){Near(KeyboardNudgePolicy.StepPoints(false,5.669921875,false,2),.5);Near(KeyboardNudgePolicy.StepPoints(false,0,false,1.5),2.0/3);}
    private static void GridStep(){Near(KeyboardNudgePolicy.StepPoints(true,5.669921875,false,1.5277),5.669921875);}
    private static void ControlStep(){Near(KeyboardNudgePolicy.StepPoints(true,5.669921875,true,2),.5);Near(KeyboardNudgePolicy.StepPoints(true,Double.NaN,true,2),.5);}
    private static void ZoomScale(){Near(KeyboardNudgePolicy.StepPoints(false,10,false,4),.25);Near(KeyboardNudgePolicy.StepPoints(false,10,false,1),1);}
    private static void Directions()
    {
        var left=KeyboardNudgePolicy.Delta(0x25,.5,.75);var up=KeyboardNudgePolicy.Delta(0x26,.5,.75);
        var right=KeyboardNudgePolicy.Delta(0x27,.5,.75);var down=KeyboardNudgePolicy.Delta(0x28,.5,.75);
        Near(left.Item1,-.5);Near(left.Item2,0);Near(right.Item1,.5);Near(right.Item2,0);
        Near(up.Item1,0);Near(up.Item2,-.75);Near(down.Item1,0);Near(down.Item2,.75);
    }
    private static void InvalidNativeInput()
    {
        Throws<InvalidOperationException>(()=>KeyboardNudgePolicy.PixelsPerPoint(100,100));
        Throws<InvalidOperationException>(()=>KeyboardNudgePolicy.PixelsPerPoint(Double.NaN,100));
        Throws<ArgumentOutOfRangeException>(()=>KeyboardNudgePolicy.StepPoints(false,1,false,0));
        Throws<InvalidOperationException>(()=>KeyboardNudgePolicy.StepPoints(true,0,false,1));
        Throws<InvalidOperationException>(()=>KeyboardNudgePolicy.StepPoints(true,Double.PositiveInfinity,false,1));
        Throws<ArgumentOutOfRangeException>(()=>KeyboardNudgePolicy.Delta(0x41,1,1));
        Throws<ArgumentOutOfRangeException>(()=>KeyboardNudgePolicy.Delta(0x25,-1,1));
    }
    private static void SingleDriver()
    {
        var result=Plan().Solve(0,.5);
        foreach(var shape in result.Geometry){Near(shape.Top,20.5);Near(shape.Width,50);Near(shape.Height,40);}
        Near(Get(result,A).Left,10);Near(Get(result,B).Left,100);Near(Get(result,C).Left,200);
    }
    private static void MultipleDrivers()
    {
        var result=Plan(new[]{A,B},new[]{new SelectionMovePosition(90,300,400,0,80)}).Solve(.5,.75);
        Near(Get(result,A).Left,10.5);Near(Get(result,B).Left,100.5);Near(Get(result,C).Left,200);
        foreach(var shape in result.Geometry)Near(shape.Top,20.75);
        Near(result.Extras.Single().Left,300.5);Near(result.Extras.Single().Top,400.75);Near(result.Extras.Single().Width,0);Near(result.Extras.Single().Height,80);
    }
    private static void RepeatedNudge()
    {
        var plan=Plan();double scale=1.5277,step=KeyboardNudgePolicy.StepPoints(false,5,false,scale),offset=0;
        SelectionMoveFrame frame=plan.Initial;
        for(int index=0;index<100;index++){offset+=step;frame=plan.Solve(0,offset);}
        Near(Get(frame,A).Top,20+100/scale,1e-9);Near(Get(frame,B).Top,20+100/scale,1e-9);
        Near(Get(plan.Initial,A).Top,20);
    }
    private static void RejectedNudge()
    {
        var plan=new SelectionMovePlan(new[]{Shape(A,10,20)},
            new[]{new LayoutConstraint("22222222-2222-2222-2222-222222222222",256,ConstraintKind.FixedPosition,new[]{A},10,20)},
            new[]{A},new[]{new SelectionMovePosition(90,300,400,0,80)});
        int writes=0;
        Throws<ConstraintConflictException>(()=>{var frame=plan.Solve(.5,0);plan.ApplyAtomic(plan.Initial,frame,next=>writes++,()=>plan.Initial);});
        Assert(writes==0,"Rejected keyboard input must not alter the native selection");
    }
    private static void AtomicNudge()
    {
        var plan=Plan(extras:new[]{new SelectionMovePosition(90,300,400,0,80)});var current=plan.Initial;
        var accepted=plan.ApplyAtomic(current,plan.Solve(0,.5),next=>current=next,()=>current);
        int writes=0;
        Throws<SelectionMoveRejectedException>(()=>plan.ApplyAtomic(accepted,plan.Solve(0,1),next=>{
            writes++;current=next;if(writes==1)throw new InvalidOperationException("native setter rejected");
        },()=>current));
        Assert(current.Matches(accepted)&&writes==2,"Keyboard rollback must restore the most recent accepted frame");
    }
    private static void GroupChildNudge()
    {
        var child=new SelectionMovePosition(10,10,20,20,30);
        var sibling=new SelectionMovePosition(11,50,30,40,20);
        Func<ShapeGeometry> group=()=>new ShapeGeometry(A,Math.Min(child.Left,sibling.Left),Math.Min(child.Top,sibling.Top),
            Math.Max(child.Left+child.Width,sibling.Left+sibling.Width)-Math.Min(child.Left,sibling.Left),
            Math.Max(child.Top+child.Height,sibling.Top+sibling.Height)-Math.Min(child.Top,sibling.Top));
        ShapeGeometry peer=new ShapeGeometry(B,150,20,80,40);
        Func<System.Collections.Generic.IReadOnlyList<ShapeGeometry>> read=()=>new[]{group(),peer};
        var relation=new LayoutConstraint("11111111-1111-1111-1111-111111111111",256,ConstraintKind.SameY,new[]{A,B});
        var plan=new GroupChildEditPlan(read(),new[]{relation},new[]{A},new[]{child},GeometryEditMode.Move,false);
        SelectionMoveFrame frame=plan.Initial;double step=KeyboardNudgePolicy.StepPoints(false,5,false,2);
        for(int repeat=1;repeat<=3;repeat++)
        {
            var delta=KeyboardNudgePolicy.Delta(0x28,step,step);
            frame=plan.ApplyAtomic(frame,plan.Target(delta.Item1*repeat,delta.Item2*repeat),values=>child=values.Single(),
                ()=>new[]{child},()=>new[]{group()},values=>{
                    Assert(values.Count==1&&values[0].Id==ShapeIdentity.Normalize(B),"The keyboard writer must exclude the edited parent");
                    peer=values.Single();
                },read);
        }
        Near(child.Top,21.5);Near(Get(frame,A).Top,21.5);Near(peer.Top,21.5);Near(sibling.Top,30);
        Near(child.Width,20);Near(child.Height,30);Near(Get(plan.Initial,A).Top,20);
    }
}
