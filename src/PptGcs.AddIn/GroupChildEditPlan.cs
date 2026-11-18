using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    public sealed class GroupChildEditRejectedException : InvalidOperationException
    {
        public GroupChildEditRejectedException(string message,Exception inner):base(message,inner){}
    }

    // Edit the real native member, then use the actual resulting group bounds
    // as a GCS driver. The group itself is deliberately never passed to a writer:
    // restoring its width/height would rescale the members being edited.
    // Multiple supplied root boxes are a mathematical capability. Native UI
    // selection expressibility is checked separately by GroupChildNativeEdit.
    public sealed class GroupChildEditPlan
    {
        private readonly List<LayoutConstraint> constraints;
        private readonly GeometryEditMode mode;
        private readonly bool keepAspectRatio;
        private readonly string localChildId=Guid.NewGuid().ToString("D");
        public ReadOnlyCollection<string> GroupIds { get; private set; }
        public SelectionMoveFrame Initial { get; private set; }

        public GroupChildEditPlan(IEnumerable<ShapeGeometry> layout,IEnumerable<LayoutConstraint> relations,
            IEnumerable<string> groupIds,IEnumerable<SelectionMovePosition> children,GeometryEditMode mode,bool keepAspectRatio)
        {
            if(relations==null)throw new ArgumentNullException("relations");
            if(groupIds==null)throw new ArgumentNullException("groupIds");
            if(children==null)throw new ArgumentNullException("children");
            if(mode!=GeometryEditMode.Move && mode!=GeometryEditMode.ResizeNW && mode!=GeometryEditMode.ResizeN &&
                mode!=GeometryEditMode.ResizeNE && mode!=GeometryEditMode.ResizeE && mode!=GeometryEditMode.ResizeSE &&
                mode!=GeometryEditMode.ResizeS && mode!=GeometryEditMode.ResizeSW && mode!=GeometryEditMode.ResizeW)
                throw new ArgumentOutOfRangeException("mode");
            Initial=new SelectionMoveFrame(layout,children);
            if(Initial.Extras.Count==0)throw new ArgumentException("At least one native member is required.","children");
            if(mode!=GeometryEditMode.Move && Initial.Extras.Count!=1)
                throw new ArgumentException("Aggregate native member resizing is not supported.","mode");
            if(mode!=GeometryEditMode.Move && (Initial.Extras[0].Width<LayoutSolver.MinimumSize || Initial.Extras[0].Height<LayoutSolver.MinimumSize))
                throw new ArgumentException("A degenerate native member can move, but its endpoint resizing remains native.","child");
            var ids=groupIds.Select(ShapeIdentity.Normalize).ToList();
            if(ids.Count==0 || ids.Distinct(StringComparer.Ordinal).Count()!=ids.Count)
                throw new ArgumentException("The edited root groups must be nonempty and unique.","groupIds");
            if(ids.Any(id=>!Initial.Geometry.Any(shape=>shape.Id==id)))
                throw new ArgumentException("An edited root group is absent from the layout.","groupIds");
            GroupIds=ids.AsReadOnly();
            constraints=LayoutSolver.ConnectedConstraints(relations,GroupIds).ToList();
            this.mode=mode;this.keepAspectRatio=keepAspectRatio;
        }

        public IReadOnlyList<SelectionMovePosition> Target(double dx,double dy)
        {
            if(Double.IsNaN(dx)||Double.IsInfinity(dx)||Double.IsNaN(dy)||Double.IsInfinity(dy))throw new ArgumentOutOfRangeException("delta");
            var before=Initial.Extras[0];
            if(mode==GeometryEditMode.Move)return Initial.Extras.Select(child=>child.Translate(dx,dy)).ToList().AsReadOnly();
            bool west=mode==GeometryEditMode.ResizeW || mode==GeometryEditMode.ResizeNW || mode==GeometryEditMode.ResizeSW;
            bool east=mode==GeometryEditMode.ResizeE || mode==GeometryEditMode.ResizeNE || mode==GeometryEditMode.ResizeSE;
            bool north=mode==GeometryEditMode.ResizeN || mode==GeometryEditMode.ResizeNW || mode==GeometryEditMode.ResizeNE;
            bool south=mode==GeometryEditMode.ResizeS || mode==GeometryEditMode.ResizeSW || mode==GeometryEditMode.ResizeSE;
            double width=west ? Math.Max(LayoutSolver.MinimumSize,before.Width-dx) : east ? Math.Max(LayoutSolver.MinimumSize,before.Width+dx) : before.Width;
            double height=north ? Math.Max(LayoutSolver.MinimumSize,before.Height-dy) : south ? Math.Max(LayoutSolver.MinimumSize,before.Height+dy) : before.Height;
            var original=new ShapeGeometry(localChildId,before.Left,before.Top,before.Width,before.Height);
            var requested=new ShapeGeometry(localChildId,west ? before.Left+before.Width-width : before.Left,
                north ? before.Top+before.Height-height : before.Top,width,height);
            var ratio=keepAspectRatio ? new[]{new LayoutConstraint(Guid.NewGuid().ToString("D"),1,ConstraintKind.AspectRatio,
                new[]{localChildId},before.Width/before.Height)} : new LayoutConstraint[0];
            var target=LayoutSolver.Solve(new[]{original},ratio,localChildId,requested,mode)[0];
            return new[]{new SelectionMovePosition(before.NativeId,target.Left,target.Top,target.Width,target.Height)};
        }

        public SelectionMoveFrame ApplyAtomic(SelectionMoveFrame previous,IEnumerable<SelectionMovePosition> targets,
            Action<IReadOnlyList<SelectionMovePosition>> writeChildren,Func<IReadOnlyList<SelectionMovePosition>> readChildren,
            Func<IReadOnlyList<ShapeGeometry>> readGroups,
            Action<IReadOnlyList<ShapeGeometry>> writeFollowers,Func<IReadOnlyList<ShapeGeometry>> readLayout)
        {
            if(previous==null)throw new ArgumentNullException("previous");
            if(targets==null)throw new ArgumentNullException("targets");
            if(writeChildren==null||readChildren==null||readGroups==null||writeFollowers==null||readLayout==null)
                throw new ArgumentNullException("nativeAccess");
            var target=targets.ToList().AsReadOnly();
            var nativeIds=new HashSet<int>(Initial.Extras.Select(child=>child.NativeId));
            if(target.Any(child=>child==null) || target.Count!=nativeIds.Count ||
                target.Select(child=>child.NativeId).Distinct().Count()!=nativeIds.Count || target.Any(child=>!nativeIds.Contains(child.NativeId)) ||
                previous.Extras.Count!=nativeIds.Count || previous.Extras.Any(child=>!nativeIds.Contains(child.NativeId)))
                throw new ArgumentException("The edit must retain every exact native member identity.");
            try
            {
                writeChildren(target);
                var actualChildren=readChildren();
                if(!SameChildren(actualChildren,target))throw new InvalidOperationException("PowerPoint did not retain the requested native member geometry.");
                var actualGroups=readGroups();
                if(actualGroups.Count!=GroupIds.Count || actualGroups.Select(group=>group.Id).Distinct().Count()!=GroupIds.Count ||
                    actualGroups.Any(group=>!GroupIds.Contains(group.Id)))
                    throw new InvalidOperationException("The native members changed root group ownership.");
                var edits=actualGroups.Select(group=>new GeometryEdit(group,GeometryEditMode.Exact)).ToList();
                IReadOnlyList<ShapeGeometry> solved;
                try {solved=LayoutSolver.Solve(Initial.Geometry,constraints,edits);}
                catch(ConstraintConflictException error)
                {throw new ConstraintConflictException(error.Message,LayoutDiagnostics.ExplainEdit(Initial.Geometry,constraints,edits),error);}
                writeFollowers(solved.Where(shape=>!GroupIds.Contains(shape.Id)).ToList().AsReadOnly());
                var actual=new SelectionMoveFrame(readLayout(),readChildren());
                var expected=new SelectionMoveFrame(solved,actualChildren);
                if(!actual.Matches(expected))throw new InvalidOperationException("PowerPoint did not retain the complete group-child edit.");
                if(LayoutSolver.MaxResidual(actual.Geometry,constraints)>LayoutSolver.PositionTolerance)
                    throw new InvalidOperationException("The group-child edit violates an external layout constraint.");
                return actual;
            }
            catch(Exception error)
            {
                try
                {
                    writeChildren(previous.Extras);
                    writeFollowers(previous.Geometry.Where(shape=>!GroupIds.Contains(shape.Id)).ToList().AsReadOnly());
                    if(!new SelectionMoveFrame(readLayout(),readChildren()).Matches(previous))
                        throw new InvalidOperationException("The previous child and follower geometry was not restored.");
                }
                catch(Exception rollback)
                {throw new InvalidOperationException("Cannot restore a rejected group-child edit without resizing its parent.",new AggregateException(error,rollback));}
                throw new GroupChildEditRejectedException("The group-child edit was rejected and its previous complete frame was restored.",error);
            }
        }

        public static bool SameChild(SelectionMovePosition first,SelectionMovePosition second)
        {
            return first!=null&&second!=null&&first.NativeId==second.NativeId &&
                Math.Abs(first.Left-second.Left)<=LayoutSolver.PositionTolerance && Math.Abs(first.Top-second.Top)<=LayoutSolver.PositionTolerance &&
                Math.Abs(first.Width-second.Width)<=LayoutSolver.PositionTolerance && Math.Abs(first.Height-second.Height)<=LayoutSolver.PositionTolerance;
        }

        public static bool SameChildren(IEnumerable<SelectionMovePosition> first,IEnumerable<SelectionMovePosition> second)
        {
            if(first==null||second==null)return false;
            var actual=first.ToList();var expected=second.ToList();
            if(actual.Count!=expected.Count || actual.Any(child=>child==null) || expected.Any(child=>child==null) ||
                expected.Select(child=>child.NativeId).Distinct().Count()!=expected.Count)return false;
            var byId=expected.ToDictionary(child=>child.NativeId);
            return actual.Select(child=>child.NativeId).Distinct().Count()==actual.Count &&
                actual.All(child=>byId.ContainsKey(child.NativeId)&&SameChild(child,byId[child.NativeId]));
        }
    }
}
