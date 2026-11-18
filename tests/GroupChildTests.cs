using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using PptGcs.AddIn;
using PptGcs.Core;

internal static class GroupChildTests
{
    private const string Group="aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa";
    private const string Peer="bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb";
    private const string Other="cccccccc-cccc-cccc-cccc-cccccccccccc";
    private const string SecondGroup="dddddddd-dddd-dddd-dddd-dddddddddddd";
    private static int passed;

    private static void Main()
    {
        Run("moving a true child drives the actual new group box without moving a sibling", MoveChild);
        Run("an interior child move remains an edit even when the outer group box is unchanged", InteriorMove);
        Run("child edge resizing preserves its opposite edge and propagates actual group width", ResizeChild);
        Run("a native locked aspect ratio uses the existing ratio solver and side anchors", LockedAspectSide);
        Run("locked-aspect corner resizing preserves the opposite native child corner", LockedAspectCorner);
        Run("an external fixed group size rejects the child edit and restores the actual child", GroupConstraintConflict);
        Run("an unrelated broken constraint cannot reject an otherwise valid group-child edit", UnrelatedBrokenComponent);
        Run("partial native child write failure restores child and followers without parent writes", ChildWriteFailure);
        Run("partial external follower failure restores both sides without parent scaling", FollowerWriteFailure);
        Run("a native child readback mismatch is rejected and restored", ChildReadbackFailure);
        Run("native child ownership changes cannot drive a different parent", WrongParentReadback);
        Run("later rejected input restores the last accepted child and follower frame", RestoreLastAccepted);
        Run("a failed rollback reports both failures rather than claiming recovery", RollbackFailure);
        Run("zero-width native members can move without forcing GCS minimum dimensions", DegenerateMove);
        Run("resize minimum conflicts are rejected before an Office write", ResizeMinimum);
        Run("unknown groups and mismatched native child identities are rejected", InvalidPlans);
        Run("selected members of one group move rigidly while an unselected sibling stays fixed", MultipleMembersMove);
        Run("the mathematical plan drives independently supplied root bounds simultaneously", MultipleRootGroups);
        Run("a mathematical multi-root conflict atomically restores all supplied children", MultipleRootConflict);
        Run("a partial second child write restores every child and external follower", MultipleChildWriteFailure);
        Run("a follower failure restores the latest accepted multi-child frame", MultipleRestoreAccepted);
        Run("multi-child inputs retain exact membership and aggregate resizing stays unsupported", MultipleInvalidInputs);
        Run("the observed outward child move is rejected before the simulated undo boundary", PreflightObservedConflict);
        Run("fixed root position still allows movement inside its enclosing rectangle", PreflightInteriorMove);
        Run("inward root shrink is deliberately left for authoritative native readback", PreflightCannotProveShrink);
        Run("fixed width rejects required child span even when every child individually fits", PreflightWidthSpan);
        Run("fixed height rejects vertical span while leaving horizontal movement free", PreflightHeightSpan);
        Run("fixed position and dimensions reject anchored right and bottom overflow", PreflightAnchoredEdges);
        Run("root envelopes stay independent and disabled or unrelated constraints are ignored", PreflightConstraintScope);
        Run("missing safe geometry supplies no invented parent bound", PreflightUnknownGeometry);
        Run("rotated, 3D and self-sizing objects cannot supply an enclosure proof", PreflightCapabilities);
        Run("a free zero-width root stays native without entering the GCS snapshot", NativeFreeZeroRoot);
        Run("unsupported mixed roots are rejected before reading a free zero-width box", NativeMixedZeroRootRejected);
        Run("the native single-root rule also rejects two nondegenerate linked roots", NativeMixedLinkedRootsRejected);
        Run("same-root native members still move and restore without writing their parent", NativeSameRootMove);
        Run("native follower failure restores actual children through the shared helper", NativeSameRootRollback);
        Run("associated capture errors retain ownership while non-child selections stay native", NativeCaptureFailureOwnership);
        Run("an intended related pointer edit consumes a failed capture and its matching up", NativePointerCaptureRejected);
        Console.WriteLine("PASS: "+passed+" group child tests");
    }
    private static void Run(string name,Action action)
    {
        try{action();passed++;Console.WriteLine("PASS "+name);}
        catch(Exception error)
        {
            Console.Error.WriteLine("FAIL "+name+": "+error.GetType().Name+": "+error.Message);
            var aggregate=error.InnerException as AggregateException;
            if(aggregate!=null)foreach(var inner in aggregate.InnerExceptions)Console.Error.WriteLine(inner.GetType().Name+": "+inner.Message);
            throw;
        }
    }
    private static void Assert(bool value,string message){if(!value)throw new Exception(message);}
    private static void Near(double actual,double expected){Assert(Math.Abs(actual-expected)<.000001,"Expected "+expected+", got "+actual);}
    private static void Throws<T>(Action action) where T:Exception
    {try{action();}catch(T){return;}throw new Exception("Expected "+typeof(T).Name);}
    private static ShapeGeometry Get(SelectionMoveFrame frame,string id){return frame.Geometry.Single(shape=>shape.Id==ShapeIdentity.Normalize(id));}
    private static LayoutConstraint Relation(ConstraintKind kind,params string[] ids)
    {return new LayoutConstraint((kind==ConstraintKind.SameY ? "11111111" : "22222222")+"-0000-0000-0000-000000000001",256,kind,ids);}

    private sealed class NativeScene
    {
        public SelectionMovePosition Child;
        public readonly SelectionMovePosition Sibling;
        private readonly Dictionary<string,ShapeGeometry> followers=new Dictionary<string,ShapeGeometry>(StringComparer.Ordinal);
        public int ParentWriteAttempts,ChildWrites,FollowerWrites;
        public NativeScene(SelectionMovePosition child=null,SelectionMovePosition sibling=null)
        {
            Child=child??new SelectionMovePosition(10,10,20,20,30);
            Sibling=sibling??new SelectionMovePosition(11,50,30,40,20);
            var group=ReadGroup();followers.Add(ShapeIdentity.Normalize(Peer),new ShapeGeometry(Peer,150,group.Top,group.Width,40));
        }
        public ShapeGeometry ReadGroup()
        {
            // Independent native-oracle model. A parent has no write API here:
            // it is always the union of the actual edited and untouched child.
            double left=Math.Min(Child.Left,Sibling.Left),top=Math.Min(Child.Top,Sibling.Top);
            double right=Math.Max(Child.Left+Child.Width,Sibling.Left+Sibling.Width);
            double bottom=Math.Max(Child.Top+Child.Height,Sibling.Top+Sibling.Height);
            return new ShapeGeometry(Group,left,top,right-left,bottom-top);
        }
        public IReadOnlyList<ShapeGeometry> ReadLayout(){return new[]{ReadGroup()}.Concat(followers.Values).ToList().AsReadOnly();}
        public void WriteChild(SelectionMovePosition value){ChildWrites++;Child=value;}
        public void WriteFollowers(IReadOnlyList<ShapeGeometry> values)
        {
            FollowerWrites++;
            foreach(var value in values)
            {
                if(value.Id==ShapeIdentity.Normalize(Group)){ParentWriteAttempts++;throw new Exception("The parent must never be resized to edit or restore a child");}
                followers[value.Id]=value;
            }
        }
        public GroupChildEditPlan Plan(GeometryEditMode mode=GeometryEditMode.Move,bool aspect=false,IEnumerable<LayoutConstraint> relations=null)
        {return new GroupChildEditPlan(ReadLayout(),relations??new[]{Relation(ConstraintKind.SameY,Group,Peer),Relation(ConstraintKind.SameWidth,Group,Peer)},new[]{Group},new[]{Child},mode,aspect);}
        public SelectionMoveFrame Frame(){return new SelectionMoveFrame(ReadLayout(),new[]{Child});}
        public SelectionMoveFrame Apply(GroupChildEditPlan plan,SelectionMoveFrame before,IEnumerable<SelectionMovePosition> target,
            Action<SelectionMovePosition> childWrite=null,Action<IReadOnlyList<ShapeGeometry>> followerWrite=null,
            Func<SelectionMovePosition> childRead=null,Func<ShapeGeometry> groupRead=null)
        {return plan.ApplyAtomic(before,target,values=>(childWrite??WriteChild)(values.Single()),
            ()=>new[]{(childRead??(()=>Child))()},()=>new[]{(groupRead??ReadGroup)()},followerWrite??WriteFollowers,ReadLayout);}
    }
    private static void MoveChild()
    {
        var scene=new NativeScene();var plan=scene.Plan();var frame=scene.Apply(plan,plan.Initial,plan.Target(-15,5));
        Near(scene.Child.Left,-5);Near(scene.Child.Top,25);Near(scene.Sibling.Left,50);Near(scene.Sibling.Top,30);
        Near(Get(frame,Group).Left,-5);Near(Get(frame,Group).Top,25);Near(Get(frame,Group).Width,95);Near(Get(frame,Group).Height,30);
        Near(Get(frame,Peer).Left,150);Near(Get(frame,Peer).Top,25);Near(Get(frame,Peer).Width,95);Near(Get(frame,Peer).Height,40);
        Assert(scene.ParentWriteAttempts==0,"The parent's own box must remain native-derived");
    }
    private static void InteriorMove()
    {
        var scene=new NativeScene(new SelectionMovePosition(10,30,30,10,10),new SelectionMovePosition(11,10,20,80,60));
        var plan=scene.Plan();var frame=scene.Apply(plan,plan.Initial,plan.Target(5,5));
        foreach(var before in plan.Initial.Geometry)
        {
            var after=frame.Geometry.Single(shape=>shape.Id==before.Id);
            Assert(!LayoutSolver.NeedsUpdate(before,after),"An interior move must not move the unaffected outer layout");
        }
        Assert(!frame.Matches(plan.Initial),"The actual child still moved even though the group box did not");
        Near(scene.Child.Left,35);Near(scene.Child.Top,35);Assert(scene.ChildWrites==1,"The child write must not be skipped");
    }
    private static void ResizeChild()
    {
        var scene=new NativeScene();var plan=scene.Plan(GeometryEditMode.ResizeW);var frame=scene.Apply(plan,plan.Initial,plan.Target(-10,0));
        Near(scene.Child.Left,0);Near(scene.Child.Width,30);Near(scene.Child.Left+scene.Child.Width,30);Near(scene.Child.Height,30);
        Near(Get(frame,Group).Width,90);Near(Get(frame,Peer).Width,90);Near(scene.Sibling.Left,50);
        Assert(scene.ParentWriteAttempts==0,"A real child resize must not be implemented as a group resize");
    }
    private static void LockedAspectSide()
    {
        var scene=new NativeScene();var plan=scene.Plan(GeometryEditMode.ResizeE,true);var target=plan.Target(10,0)[0];
        Near(target.Left,10);Near(target.Top,12.5);Near(target.Width,30);Near(target.Height,45);
        var frame=scene.Apply(plan,plan.Initial,new[]{target});
        Near(Get(frame,Group).Top,12.5);Near(Get(frame,Group).Height,45);Near(Get(frame,Peer).Top,12.5);
        Near(scene.Sibling.Left,50);Near(scene.Sibling.Top,30);
    }
    private static void LockedAspectCorner()
    {
        var scene=new NativeScene();var plan=scene.Plan(GeometryEditMode.ResizeNW,true);var target=plan.Target(-20,-10)[0];
        Near(target.Width/target.Height,2.0/3);Near(target.Left+target.Width,30);Near(target.Top+target.Height,50);
        scene.Apply(plan,plan.Initial,new[]{target});Assert(scene.ParentWriteAttempts==0,"A corner edit must leave the parent out of writeback");
    }
    private static void GroupConstraintConflict()
    {
        var scene=new NativeScene();
        var relations=new[]{new LayoutConstraint("33333333-3333-3333-3333-333333333333",256,ConstraintKind.FixedWidth,new[]{Group},80)};
        var plan=scene.Plan(relations:relations);GroupChildEditRejectedException failure=null;
        try{scene.Apply(plan,plan.Initial,plan.Target(-15,5));}catch(GroupChildEditRejectedException error){failure=error;}
        Assert(failure!=null && failure.InnerException is ConstraintConflictException,"The external width conflict must be explained");
        Assert(scene.Frame().Matches(plan.Initial)&&scene.ChildWrites==2&&scene.ParentWriteAttempts==0,"The actual child must restore the native parent bounds");
    }
    private static void UnrelatedBrokenComponent()
    {
        var scene=new NativeScene();var unrelated=new ShapeGeometry(Other,500,400,17,15);
        Func<IReadOnlyList<ShapeGeometry>> read=()=>scene.ReadLayout().Concat(new[]{unrelated}).ToList().AsReadOnly();
        var relations=new[]{Relation(ConstraintKind.SameY,Group,Peer),
            new LayoutConstraint("33333333-3333-3333-3333-333333333333",256,ConstraintKind.FixedWidth,new[]{Other},99)};
        var plan=new GroupChildEditPlan(read(),relations,new[]{Group},new[]{scene.Child},GeometryEditMode.Move,false);
        var frame=plan.ApplyAtomic(plan.Initial,plan.Target(-15,5),values=>scene.WriteChild(values.Single()),()=>new[]{scene.Child},()=>new[]{scene.ReadGroup()},values=>{
            Assert(!LayoutSolver.NeedsUpdate(unrelated,values.Single(value=>value.Id==ShapeIdentity.Normalize(Other))),
                "The disconnected broken object must remain untouched");
            scene.WriteFollowers(values.Where(value=>value.Id!=ShapeIdentity.Normalize(Other)).ToList().AsReadOnly());
        },read);
        Near(Get(frame,Peer).Top,25);Near(Get(frame,Other).Width,17);
        Assert(LayoutSolver.MaxResidual(frame.Geometry,relations)>LayoutSolver.PositionTolerance,
            "The unrelated broken constraint must remain visible rather than being silently repaired");
        Assert(scene.ChildWrites==1&&scene.ParentWriteAttempts==0,"The independent group edit must be accepted without rollback");
    }
    private static void ChildWriteFailure()
    {
        var scene=new NativeScene();var plan=scene.Plan();int writes=0;
        Throws<GroupChildEditRejectedException>(()=>scene.Apply(plan,plan.Initial,plan.Target(-15,5),childWrite:target=>{
            scene.WriteChild(target);if(++writes==1)throw new InvalidOperationException("native child setter failed after a partial change");
        }));
        Assert(scene.Frame().Matches(plan.Initial)&&scene.ParentWriteAttempts==0,"Failed native child writes must not be repaired by scaling the parent");
    }
    private static void FollowerWriteFailure()
    {
        var scene=new NativeScene();var plan=scene.Plan();int writes=0;
        Throws<GroupChildEditRejectedException>(()=>scene.Apply(plan,plan.Initial,plan.Target(-15,5),followerWrite:targets=>{
            scene.WriteFollowers(targets);if(++writes==1)throw new InvalidOperationException("native follower setter failed");
        }));
        Assert(scene.Frame().Matches(plan.Initial)&&scene.ParentWriteAttempts==0&&writes==2,"Both child and external followers must restore together");
    }
    private static void ChildReadbackFailure()
    {
        var scene=new NativeScene();var plan=scene.Plan();int reads=0;
        Throws<GroupChildEditRejectedException>(()=>scene.Apply(plan,plan.Initial,plan.Target(-15,5),childRead:()=>{
            var child=scene.Child;return ++reads==1 ? new SelectionMovePosition(child.NativeId,child.Left,child.Top+1,child.Width,child.Height) : child;
        }));
        Assert(scene.Frame().Matches(plan.Initial),"A silent child-position drift must not be accepted");
    }
    private static void WrongParentReadback()
    {
        var scene=new NativeScene();var plan=scene.Plan();
        Throws<GroupChildEditRejectedException>(()=>scene.Apply(plan,plan.Initial,plan.Target(-15,5),groupRead:()=>new ShapeGeometry(Other,0,0,100,100)));
        Assert(scene.Frame().Matches(plan.Initial),"A member cannot silently switch which group drives the solver");
    }
    private static void RestoreLastAccepted()
    {
        var scene=new NativeScene();var plan=scene.Plan();var accepted=scene.Apply(plan,plan.Initial,plan.Target(-15,5));int writes=0;
        Throws<GroupChildEditRejectedException>(()=>scene.Apply(plan,accepted,plan.Target(-20,10),followerWrite:targets=>{
            scene.WriteFollowers(targets);if(++writes==1)throw new InvalidOperationException("follower unavailable");
        }));
        Assert(scene.Frame().Matches(accepted)&&!scene.Frame().Matches(plan.Initial),"Rollback must preserve the latest accepted child frame");
    }
    private static void RollbackFailure()
    {
        var scene=new NativeScene();var plan=scene.Plan();Exception failure=null;
        try{scene.Apply(plan,plan.Initial,plan.Target(-15,5),followerWrite:targets=>{throw new InvalidOperationException("persistent follower failure");});}
        catch(Exception error){failure=error;}
        Assert(failure is InvalidOperationException && !(failure is GroupChildEditRejectedException) && failure.InnerException is AggregateException,
            "A failed rollback must not be reported as a recovered child edit");
        Assert(scene.ParentWriteAttempts==0,"No failure path is allowed to resize the parent");
    }
    private static void DegenerateMove()
    {
        var scene=new NativeScene(new SelectionMovePosition(10,10,20,0,30));var plan=scene.Plan();
        var frame=scene.Apply(plan,plan.Initial,plan.Target(-10,-5));
        Near(scene.Child.Width,0);Near(scene.Child.Left,0);Near(scene.Child.Top,15);Near(Get(frame,Peer).Top,15);
        Throws<ArgumentException>(()=>scene.Plan(GeometryEditMode.ResizeE));
    }
    private static void ResizeMinimum()
    {
        var scene=new NativeScene(new SelectionMovePosition(10,10,20,20,10));var plan=scene.Plan(GeometryEditMode.ResizeE,true);
        Throws<ConstraintConflictException>(()=>plan.Target(-30,0));
        Assert(scene.ChildWrites==0&&scene.FollowerWrites==0,"An impossible local aspect/minimum edit must not touch Office");
    }
    private static void InvalidPlans()
    {
        var scene=new NativeScene();var rows=scene.ReadLayout();
        Throws<ArgumentException>(()=>new GroupChildEditPlan(rows,new LayoutConstraint[0],new[]{Other},new[]{scene.Child},GeometryEditMode.Move,false));
        Throws<ArgumentOutOfRangeException>(()=>scene.Plan(GeometryEditMode.Exact));
        var plan=scene.Plan();
        Throws<ArgumentException>(()=>scene.Apply(plan,plan.Initial,new[]{new SelectionMovePosition(99,10,20,20,30)}));
        Assert(scene.ChildWrites==0,"A wrong child ID cannot be written before validation");
    }

    private sealed class MultiNativeScene
    {
        public readonly Dictionary<int,SelectionMovePosition> Children=new Dictionary<int,SelectionMovePosition>();
        private readonly Dictionary<string,int[]> roots=new Dictionary<string,int[]>(StringComparer.Ordinal);
        private ShapeGeometry peer;
        public int ChildWrites,ParentWriteAttempts;
        public MultiNativeScene(bool twoRoots=false)
        {
            Children.Add(10,new SelectionMovePosition(10,10,20,20,30));
            Children.Add(11,new SelectionMovePosition(11,50,30,40,20));
            Children.Add(12,new SelectionMovePosition(12,30,25,10,10));
            roots.Add(ShapeIdentity.Normalize(Group),new[]{10,11,12});
            if(twoRoots)
            {
                Children.Add(20,new SelectionMovePosition(20,200,20,20,30));
                Children.Add(21,new SelectionMovePosition(21,250,30,30,20));
                roots.Add(ShapeIdentity.Normalize(SecondGroup),new[]{20,21});
            }
            peer=new ShapeGeometry(Peer,350,20,80,40);
        }
        public IReadOnlyList<ShapeGeometry> ReadGroups()
        {
            return roots.Select(root=>{
                var members=root.Value.Select(id=>Children[id]).ToList();
                double left=members.Min(child=>child.Left),top=members.Min(child=>child.Top);
                return new ShapeGeometry(root.Key,left,top,members.Max(child=>child.Left+child.Width)-left,
                    members.Max(child=>child.Top+child.Height)-top);
            }).ToList().AsReadOnly();
        }
        public IReadOnlyList<ShapeGeometry> ReadLayout(){return ReadGroups().Concat(new[]{peer}).ToList().AsReadOnly();}
        public IReadOnlyList<SelectionMovePosition> ReadChildren(int[] ids){return ids.Select(id=>Children[id]).ToList().AsReadOnly();}
        public void WriteChildren(IReadOnlyList<SelectionMovePosition> values)
        {foreach(var child in values){ChildWrites++;Children[child.NativeId]=child;}}
        public void WriteFollowers(IReadOnlyList<ShapeGeometry> values)
        {
            foreach(var target in values)
            {
                if(roots.ContainsKey(target.Id)){ParentWriteAttempts++;throw new Exception("No edited root may enter the writer");}
                Assert(target.Id==ShapeIdentity.Normalize(Peer),"The follower writer must receive only external objects");peer=target;
            }
        }
        public GroupChildEditPlan Plan(int[] selected,IEnumerable<LayoutConstraint> relations)
        {return new GroupChildEditPlan(ReadLayout(),relations,roots.Keys,ReadChildren(selected),GeometryEditMode.Move,false);}
        public SelectionMoveFrame Apply(GroupChildEditPlan plan,SelectionMoveFrame before,IReadOnlyList<SelectionMovePosition> target,
            Action<IReadOnlyList<SelectionMovePosition>> childWriter=null,Action<IReadOnlyList<ShapeGeometry>> followerWriter=null)
        {
            var ids=plan.Initial.Extras.Select(child=>child.NativeId).ToArray();
            return plan.ApplyAtomic(before,target,childWriter??WriteChildren,()=>ReadChildren(ids),ReadGroups,
                followerWriter??WriteFollowers,ReadLayout);
        }
        public SelectionMoveFrame Frame(int[] selected){return new SelectionMoveFrame(ReadLayout(),ReadChildren(selected));}
    }
    private static void MultipleMembersMove()
    {
        var scene=new MultiNativeScene();int[] selected={10,11};
        var plan=scene.Plan(selected,new[]{Relation(ConstraintKind.SameY,Group,Peer),Relation(ConstraintKind.SameWidth,Group,Peer)});
        var frame=scene.Apply(plan,plan.Initial,plan.Target(-15,5));
        Near(scene.Children[10].Left,-5);Near(scene.Children[10].Top,25);Near(scene.Children[11].Left,35);Near(scene.Children[11].Top,35);
        Near(scene.Children[11].Left-scene.Children[10].Left,40);Near(scene.Children[12].Left,30);Near(scene.Children[12].Top,25);
        Near(Get(frame,Group).Left,-5);Near(Get(frame,Group).Top,25);Near(Get(frame,Group).Width,80);Near(Get(frame,Peer).Top,25);
        Assert(scene.ParentWriteAttempts==0,"Both selected members must move without a parent transform");
    }
    private static void MultipleRootGroups()
    {
        var scene=new MultiNativeScene(true);int[] selected={10,20};
        var plan=scene.Plan(selected,new[]{Relation(ConstraintKind.SameY,Group,SecondGroup,Peer)});
        var frame=scene.Apply(plan,plan.Initial,plan.Target(-15,5));
        Near(Get(frame,Group).Left,-5);Near(Get(frame,Group).Top,25);Near(Get(frame,Group).Width,95);
        Near(Get(frame,SecondGroup).Left,185);Near(Get(frame,SecondGroup).Top,25);Near(Get(frame,SecondGroup).Width,95);
        Near(Get(frame,Peer).Top,25);Near(scene.Children[11].Left,50);Near(scene.Children[21].Left,250);
        Assert(scene.ParentWriteAttempts==0&&frame.Extras.Count==2,"Every edited root must be a read-only exact driver");
    }
    private static void MultipleRootConflict()
    {
        var scene=new MultiNativeScene(true);int[] selected={11,20};
        var plan=scene.Plan(selected,new[]{Relation(ConstraintKind.SameWidth,Group,SecondGroup)});
        GroupChildEditRejectedException rejected=null;
        try{scene.Apply(plan,plan.Initial,plan.Target(15,0));}catch(GroupChildEditRejectedException error){rejected=error;}
        Assert(rejected!=null&&rejected.InnerException is ConstraintConflictException,
            "Two incompatible authoritative root widths must be rejected rather than resizing an edited parent");
        Assert(scene.Frame(selected).Matches(plan.Initial)&&scene.ParentWriteAttempts==0,"Every child and derived root must restore together");
    }
    private static void MultipleChildWriteFailure()
    {
        var scene=new MultiNativeScene();int[] selected={10,11};var plan=scene.Plan(selected,new[]{Relation(ConstraintKind.SameY,Group,Peer)});int attempts=0;
        Throws<GroupChildEditRejectedException>(()=>scene.Apply(plan,plan.Initial,plan.Target(-15,5),childWriter:values=>{
            scene.WriteChildren(new[]{values[0]});if(++attempts==1)throw new InvalidOperationException("second selected child setter is unavailable");
            scene.WriteChildren(values.Skip(1).ToList().AsReadOnly());
        }));
        Assert(scene.Frame(selected).Matches(plan.Initial)&&scene.ParentWriteAttempts==0,"Partial child writes cannot leave any selected member translated");
    }
    private static void MultipleRestoreAccepted()
    {
        var scene=new MultiNativeScene(true);int[] selected={10,20};
        var plan=scene.Plan(selected,new[]{Relation(ConstraintKind.SameY,Group,SecondGroup,Peer)});
        var accepted=scene.Apply(plan,plan.Initial,plan.Target(-15,5));int attempts=0;
        Throws<GroupChildEditRejectedException>(()=>scene.Apply(plan,accepted,plan.Target(-20,3),followerWriter:values=>{
            scene.WriteFollowers(values);if(++attempts==1)throw new InvalidOperationException("external follower rejected its new position");
        }));
        Assert(scene.Frame(selected).Matches(accepted)&&!accepted.Matches(plan.Initial)&&scene.ParentWriteAttempts==0,
            "Failure must restore the last accepted multi-child frame instead of the initial frame");
    }
    private static void MultipleInvalidInputs()
    {
        var scene=new MultiNativeScene();int[] selected={10,11};var plan=scene.Plan(selected,new LayoutConstraint[0]);
        Throws<ArgumentException>(()=>new GroupChildEditPlan(scene.ReadLayout(),new LayoutConstraint[0],new[]{Group},scene.ReadChildren(selected),GeometryEditMode.ResizeE,false));
        Throws<ArgumentException>(()=>scene.Apply(plan,plan.Initial,new[]{plan.Target(1,1)[0]}));
        Throws<ArgumentException>(()=>scene.Apply(plan,plan.Initial,new[]{plan.Target(1,1)[0],plan.Target(1,1)[0]}));
        Assert(scene.ChildWrites==0,"Missing or repeated selected members must fail before any Office write");
        Assert(GroupChildEditPlan.SameChildren(plan.Initial.Extras,plan.Initial.Extras.Reverse()),"Selection membership comparison must ignore Office range enumeration order");
    }
    private static LayoutConstraint Fixed(ConstraintKind kind,string root,double value,double? second=null,bool enabled=true)
    {return new LayoutConstraint(Guid.NewGuid().ToString("D"),256,kind,new[]{root},value,second,enabled);}
    private static GroupChildBound Bound(string root,int native,double left,double top,double width=20,double height=20)
    {return new GroupChildBound(root,new SelectionMovePosition(native,left,top,width,height));}
    private static void PreflightObservedConflict()
    {
        const double left=163.48487854003906,top=78.908821105957031;
        var constraints=new[]{Fixed(ConstraintKind.FixedPosition,Group,left,top)};
        int undoBoundaries=0,nativeWrites=0;
        foreach(double offset in new[]{-13.0,-28.0})
            Throws<ConstraintConflictException>(()=>{
                GroupChildPreflight.ThrowIfProvablyImpossible(constraints,new[]{Bound(Group,3,left,top+offset,140,135.16),
                    Bound(Group,4,323.48,298.79,182.09,70)});
                undoBoundaries++;nativeWrites++;
            });
        Assert(undoBoundaries==0&&nativeWrites==0,"Every rejected pointer sample must precede both native history and geometry writes");
    }
    private static void PreflightInteriorMove()
    {
        var constraints=new[]{Fixed(ConstraintKind.FixedPosition,Group,10,20),Fixed(ConstraintKind.FixedWidth,Group,100),Fixed(ConstraintKind.FixedHeight,Group,80)};
        GroupChildPreflight.ThrowIfProvablyImpossible(constraints,new[]{Bound(Group,10,35,40),Bound(Group,11,10,20),Bound(Group,12,90,80)});
    }
    private static void PreflightCannotProveShrink()
    {
        // The new child minimum is below the old top. This gate cannot claim
        // the real group retains empty space there or predict its tight box.
        GroupChildPreflight.ThrowIfProvablyImpossible(new[]{Fixed(ConstraintKind.FixedPosition,Group,10,20)},
            new[]{Bound(Group,10,10,25),Bound(Group,11,50,30)});
    }
    private static void PreflightWidthSpan()
    {
        var constraints=new[]{Fixed(ConstraintKind.FixedWidth,Group,80)};
        Throws<ConstraintConflictException>(()=>GroupChildPreflight.ThrowIfProvablyImpossible(constraints,
            new[]{Bound(Group,10,0,0),Bound(Group,11,100,0)}));
        GroupChildPreflight.ThrowIfProvablyImpossible(constraints,new[]{Bound(Group,10,300,0),Bound(Group,11,340,0)});
    }
    private static void PreflightHeightSpan()
    {
        var constraints=new[]{Fixed(ConstraintKind.FixedHeight,Group,70)};
        Throws<ConstraintConflictException>(()=>GroupChildPreflight.ThrowIfProvablyImpossible(constraints,
            new[]{Bound(Group,10,0,0),Bound(Group,11,0,60)}));
        GroupChildPreflight.ThrowIfProvablyImpossible(constraints,new[]{Bound(Group,10,-1000,30),Bound(Group,11,1000,60)});
    }
    private static void PreflightAnchoredEdges()
    {
        var constraints=new[]{Fixed(ConstraintKind.FixedPosition,Group,10,20),Fixed(ConstraintKind.FixedWidth,Group,100),Fixed(ConstraintKind.FixedHeight,Group,80)};
        Throws<ConstraintConflictException>(()=>GroupChildPreflight.ThrowIfProvablyImpossible(constraints,new[]{Bound(Group,10,95,30)}));
        Throws<ConstraintConflictException>(()=>GroupChildPreflight.ThrowIfProvablyImpossible(constraints,new[]{Bound(Group,10,30,85)}));
        GroupChildPreflight.ThrowIfProvablyImpossible(constraints,new[]{Bound(Group,10,10-.005,20-.005),Bound(Group,11,90+.005,80+.005)});
    }
    private static void PreflightConstraintScope()
    {
        var constraints=new[]{Fixed(ConstraintKind.FixedWidth,Group,80),Fixed(ConstraintKind.FixedWidth,SecondGroup,80),
            Fixed(ConstraintKind.FixedPosition,Group,500,500,false),Fixed(ConstraintKind.FixedPosition,Other,500,500)};
        GroupChildPreflight.ThrowIfProvablyImpossible(constraints,new[]{Bound(Group,10,0,0),Bound(SecondGroup,20,1000,1000)});
    }
    private static void PreflightUnknownGeometry()
    {
        GroupChildPreflight.ThrowIfProvablyImpossible(new[]{Fixed(ConstraintKind.FixedPosition,Group,10,20)},new GroupChildBound[0]);
        Throws<ArgumentException>(()=>GroupChildPreflight.ThrowIfProvablyImpossible(new LayoutConstraint[0],new GroupChildBound[]{null}));
    }
    private static void PreflightCapabilities()
    {
        Assert(GroupChildPreflight.IsAxisAlignedWithout3D(0,false,0,0,0)&&GroupChildPreflight.IsAxisAlignedWithout3D(360,false,0,0,0),"Known identity rotations can supply a proof");
        Assert(!GroupChildPreflight.IsAxisAlignedWithout3D(.001,false,0,0,0)&&!GroupChildPreflight.IsAxisAlignedWithout3D(0,true,0,0,0)&&
            !GroupChildPreflight.IsAxisAlignedWithout3D(0,false,1,0,0)&&!GroupChildPreflight.IsAxisAlignedWithout3D(0,false,0,1,0)&&
            !GroupChildPreflight.IsAxisAlignedWithout3D(0,false,0,0,1)&&!GroupChildPreflight.IsAxisAlignedWithout3D(Double.NaN,false,0,0,0),
            "A projection or even small genuine rotation must not be approximated as an unrotated box");
        Assert(GroupChildPreflight.HasStableShapeBounds(1,false)&&GroupChildPreflight.HasStableShapeBounds(17,false),"Fixed geometry shapes and text boxes qualify");
        Assert(!GroupChildPreflight.HasStableShapeBounds(1,true)&&!GroupChildPreflight.HasStableShapeBounds(6,false)&&
            !GroupChildPreflight.HasStableShapeBounds(19,false),"Self-sizing text, selected nested groups and tables supply no unproved bounds");
    }

    private sealed class NativeCaptureScene
    {
        public readonly GroupChildFakeShape A=new GroupChildFakeShape(10,1,10,20,20,30);
        public readonly GroupChildFakeShape B=new GroupChildFakeShape(11,1,50,30,40,20);
        public readonly GroupChildFakeShape Interior=new GroupChildFakeShape(12,1,30,25,10,10);
        public readonly GroupChildFakeShape FreeA=new GroupChildFakeShape(20,9,200,20,0,30);
        public readonly GroupChildFakeShape FreeB=new GroupChildFakeShape(21,9,200,70,0,30);
        public readonly GroupChildFakeShape Root,FreeRoot;
        public readonly GroupChildFakeShape Follower=new GroupChildFakeShape(102,1,150,20,80,30);
        public readonly LayoutContext Context;
        public NativeCaptureScene(bool linkSecondRoot=false)
        {
            Root=GroupChildFakeShape.Group(100,A,B,Interior);FreeRoot=GroupChildFakeShape.Group(200,FreeA,FreeB);
            Context=new LayoutContext();
            Context.Shapes.Add(ShapeIdentity.Normalize(Group),Root);Context.Shapes.Add(ShapeIdentity.Normalize(Peer),Follower);
            Context.Geometry.Add(PowerPointLayout.Geometry(Root,Group));Context.Geometry.Add(PowerPointLayout.Geometry(Follower,Peer));
            Context.Constraints.Add(Relation(ConstraintKind.SameY,Group,Peer));Context.Constraints.Add(Relation(ConstraintKind.SameWidth,Group,Peer));
            if(linkSecondRoot)
            {
                FreeA.SetInitialBounds(200,20,20,30);FreeB.SetInitialBounds(250,30,30,20);
                Context.Shapes.Add(ShapeIdentity.Normalize(SecondGroup),FreeRoot);
                Context.Geometry.Add(PowerPointLayout.Geometry(FreeRoot,SecondGroup));
                Context.Constraints.Add(new LayoutConstraint("44444444-4444-4444-4444-444444444444",256,ConstraintKind.SameY,new[]{Group,SecondGroup}));
            }
        }
        public GroupChildNativeEdit Capture(params object[] children)
        {return GroupChildNativeEdit.Capture(Context,children,GeometryEditMode.Move,false);}
        public void AssertNoParentWrites()
        {Assert(Root.ParentWriteAttempts==0&&FreeRoot.ParentWriteAttempts==0,"No edited parent may enter any native geometry writer");}
    }
    private static void NativeFreeZeroRoot()
    {
        var scene=new NativeCaptureScene();Near(scene.FreeRoot.Width,0);scene.FreeA.ThrowOnWidthRead=true;
        Assert(scene.Capture(scene.FreeA)==null,"An unconstrained native child must stay native");
        Assert(scene.FreeA.ThrowOnWidthRead&&scene.Context.Shapes.Count==2&&scene.Context.Geometry.Count==2,
            "A free root must not be read as ShapeGeometry or injected into the constraint context");
        scene.AssertNoParentWrites();
    }
    private static void NativeMixedZeroRootRejected()
    {
        var scene=new NativeCaptureScene();scene.FreeA.ThrowOnWidthRead=true;GroupChildCaptureException rejected=null;
        try{scene.Capture(scene.A,scene.FreeA);}catch(GroupChildCaptureException error){rejected=error;}
        Assert(rejected!=null&&rejected.InnerException.Message.Contains("跨根组合"),"The native selection limit must explain the rejection");
        Assert(scene.FreeA.ThrowOnWidthRead&&scene.Context.Shapes.Count==2&&scene.Context.Geometry.Count==2,
            "The invalid cross-root request must fail before a degenerate root is added to Core");
        Assert(scene.A.Writes==0&&scene.FreeA.Writes==0,"A rejected native selection must not alter either group member");scene.AssertNoParentWrites();
    }
    private static void NativeMixedLinkedRootsRejected()
    {
        var scene=new NativeCaptureScene(true);Near(scene.FreeRoot.Width,80);
        Throws<GroupChildCaptureException>(()=>scene.Capture(scene.A,scene.FreeA));
        Assert(scene.Context.Shapes.Count==3&&scene.Context.Geometry.Count==3&&scene.A.Writes==0&&scene.FreeA.Writes==0,
            "A solvable multi-root equation must not be mistaken for a representable native ChildShapeRange");
        scene.AssertNoParentWrites();
    }
    private static void NativeSameRootMove()
    {
        var scene=new NativeCaptureScene();var edit=scene.Capture(scene.A,scene.B);var before=edit.Plan.Initial;
        var frame=edit.Apply(edit.Plan.Target(-10,5));
        Assert(edit.Groups.Count==1&&edit.Plan.GroupIds.Count==1&&scene.Context.Shapes.Count==2,"Only the actual constrained root may drive Core");
        Near(scene.A.Left,0);Near(scene.A.Top,25);Near(scene.B.Left,40);Near(scene.B.Top,35);
        Near(scene.Interior.Left,30);Near(scene.Interior.Top,25);Near(Get(frame,Group).Top,25);Near(scene.Follower.Top,25);
        edit.Restore(before);Assert(edit.ReadFrame().Matches(before),"Restoring actual children must reproduce the initial native root bounds");scene.AssertNoParentWrites();
    }
    private static void NativeSameRootRollback()
    {
        var scene=new NativeCaptureScene();var edit=scene.Capture(scene.A);bool failed=false;
        scene.Follower.AfterWrite=(axis,value)=>{if(axis=="Top"&&!failed){failed=true;throw new InvalidOperationException("native follower setter failed after writing");}};
        Throws<GroupChildEditRejectedException>(()=>edit.Apply(edit.Plan.Target(-10,5)));
        Assert(failed&&edit.ReadFrame().Matches(edit.Plan.Initial)&&edit.LastFrame.Matches(edit.Plan.Initial),"A failed native follower must restore children, root and external geometry");
        scene.AssertNoParentWrites();
    }
    private static void NativeCaptureFailureOwnership()
    {
        var scene=new NativeCaptureScene();scene.A.ThrowOnWidthRead=true;GroupChildCaptureException rejected=null;
        try{scene.Capture(scene.A);}catch(GroupChildCaptureException error){rejected=error;}
        Assert(rejected!=null&&rejected.InnerException is InvalidOperationException,"An associated metadata failure must be distinguished from a native miss");
        Assert(scene.Capture(new object[0])==null&&scene.Capture(scene.Root)==null,
            "An empty or top-level selection must not be claimed by the child edit router");
        Assert(scene.Context.Shapes.Count==2&&scene.A.Writes==0,"Capture rejection must be nonmutating");scene.AssertNoParentWrites();
    }
    private static void NativePointerCaptureRejected()
    {
        var scene=new NativeCaptureScene();scene.A.ThrowOnWidthRead=true;
        var type=typeof(ConstrainedDragController);var method=type.GetMethod("BeginGroupChildren",BindingFlags.NonPublic|BindingFlags.Instance);
        var parameters=method.GetParameters();
        // Invoke only the pure, pre-input capture branch. No controller
        // constructor, Win32 hook, focus/capture API or real input is involved.
        object controller=FormatterServices.GetUninitializedObject(type);
        object map=Activator.CreateInstance(parameters[1].ParameterType,true),mouse=Activator.CreateInstance(parameters[2].ParameterType);
        object mode=Enum.Parse(parameters[9].ParameterType,"Move");
        bool owned=(bool)method.Invoke(controller,new object[]{scene.Context,map,mouse,0x0201,IntPtr.Zero,(uint)96,8.0,1,new object[]{scene.A},mode,true});
        bool paired=(bool)type.GetField("pendingLeftUp",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(controller);
        Assert(owned&&paired&&scene.A.Writes==0,"A known related child capture failure must not leak its native DOWN or matching UP");
        IntPtr payload=Marshal.AllocHGlobal(Marshal.SizeOf(mouse));
        try
        {
            Marshal.StructureToPtr(mouse,payload,false);
            var result=(IntPtr)type.GetMethod("OnMouse",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(controller,new object[]{0,new IntPtr(0x0202),payload});
            Assert(result==new IntPtr(1)&&!(bool)type.GetField("pendingLeftUp",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(controller),
                "The rejected capture's matching UP must be consumed exactly once without native dispatch");
        }
        finally{Marshal.FreeHGlobal(payload);}
        bool doubleClick=(bool)type.GetMethod("TryBegin",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(controller,new object[]{mouse,0x0203});
        Assert(!doubleClick,"Native double-click text entry must remain outside the child capture rejection route");
        scene.AssertNoParentWrites();
    }
}

// Public so the production add-in's dynamic Office calls can bind against an
// independent native geometry oracle in this executable test assembly.
public sealed class GroupChildFakeShape
{
    private double left,top,width,height;
    public int Id {get;private set;}
    public int Type {get;private set;}
    public int Child {get{return ParentGroup==null?0:-1;}}
    public GroupChildFakeShape ParentGroup {get;private set;}
    public readonly List<GroupChildFakeShape> Members=new List<GroupChildFakeShape>();
    public int LockAspectRatio {get;set;}
    public int Writes,ParentWriteAttempts;
    public bool ThrowOnWidthRead;
    public Action<string,double> AfterWrite;
    public GroupChildFakeShape(int id,int type,double x,double y,double w,double h)
    {Id=id;Type=type;SetInitialBounds(x,y,w,h);}
    public void SetInitialBounds(double x,double y,double w,double h){left=x;top=y;width=w;height=h;}
    public static GroupChildFakeShape Group(int id,params GroupChildFakeShape[] children)
    {
        var root=new GroupChildFakeShape(id,6,0,0,0,0);root.Members.AddRange(children);
        foreach(var child in children)child.ParentGroup=root;return root;
    }
    public double Left {get{return Type==6?Members.Min(child=>child.Left):left;}set{Write("Left",value);left=value;Notify("Left",value);}}
    public double Top {get{return Type==6?Members.Min(child=>child.Top):top;}set{Write("Top",value);top=value;Notify("Top",value);}}
    public double Width
    {
        get
        {
            if(ThrowOnWidthRead){ThrowOnWidthRead=false;throw new InvalidOperationException("native width getter failed");}
            return Type==6?Members.Max(child=>child.Left+child.Width)-Left:width;
        }
        set{Write("Width",value);width=value;Notify("Width",value);}
    }
    public double Height {get{return Type==6?Members.Max(child=>child.Top+child.Height)-Top:height;}set{Write("Height",value);height=value;Notify("Height",value);}}
    private void Write(string axis,double value)
    {if(Type==6){ParentWriteAttempts++;throw new InvalidOperationException("The test exposes no parent geometry setter");}Writes++;}
    private void Notify(string axis,double value){if(AfterWrite!=null)AfterWrite(axis,value);}
}
