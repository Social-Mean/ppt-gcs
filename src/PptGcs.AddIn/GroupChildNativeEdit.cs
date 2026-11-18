using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.InteropServices;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    // Raised only after a native child selection is known to include a
    // constrained root. Input routers must reject that intended edit instead
    // of passing its DOWN through as an unconstrained native move.
    public sealed class GroupChildCaptureException : InvalidOperationException
    {
        public GroupChildCaptureException(string message,Exception inner):base(message,inner){}
    }

    // One native transaction for pointer, keyboard and Ribbon field edits of
    // real group members. Edited parent boxes are always read-only drivers.
    public sealed class GroupChildNativeEdit
    {
        private readonly Dictionary<string,object> groups;
        private readonly Dictionary<int,string> childRoots;
        private readonly Dictionary<int,object> childrenById;
        private readonly LayoutContext followers;
        private readonly List<LayoutConstraint> preflightConstraints;
        public LayoutContext Context {get;private set;}
        public GroupChildEditPlan Plan {get;private set;}
        public ReadOnlyCollection<object> Children {get;private set;}
        public IReadOnlyDictionary<string,object> Groups {get{return groups;}}
        public IReadOnlyDictionary<int,string> ChildRoots {get{return childRoots;}}
        public SelectionMoveFrame LastFrame {get;private set;}

        private GroupChildNativeEdit(LayoutContext context,object[] children,GeometryEditMode mode,bool keepAspectRatio)
        {
            Context=context;Children=Array.AsReadOnly(children);
            childrenById=children.ToDictionary(child=>(int)((dynamic)child).Id);
            var roots=children.Select(ShapeCapabilities.TopLevel).GroupBy(group=>(int)((dynamic)group).Id).Select(group=>group.First()).ToList();
            groups=new Dictionary<string,object>(StringComparer.Ordinal);
            foreach(object group in roots)
            {
                string id=FindGroupId(context,group);
                if(id==null)throw new InvalidOperationException("Only a constrained native root may enter the child solver context.");
                groups.Add(id,group);
            }
            var rootIds=groups.ToDictionary(pair=>(int)((dynamic)pair.Value).Id,pair=>pair.Key);
            childRoots=children.ToDictionary(child=>(int)((dynamic)child).Id,
                child=>rootIds[(int)((dynamic)ShapeCapabilities.TopLevel(child)).Id]);
            preflightConstraints=context.Constraints.Where(relation=>relation.Enabled&&relation.ShapeIds.Count==1&&groups.ContainsKey(relation.ShapeIds[0])&&
                (relation.Kind==ConstraintKind.FixedPosition||relation.Kind==ConstraintKind.FixedWidth||relation.Kind==ConstraintKind.FixedHeight)).ToList();
            followers=new LayoutContext {Window=context.Window,Presentation=context.Presentation,Slide=context.Slide,
                SlideId=context.SlideId,Document=context.Document,Constraints=context.Constraints.ToList(),
                Geometry=context.Geometry.Where(shape=>!groups.ContainsKey(shape.Id)).ToList(),
                Shapes=context.Shapes.Where(pair=>!groups.ContainsKey(pair.Key)).ToDictionary(pair=>pair.Key,pair=>pair.Value,StringComparer.Ordinal)};
            Plan=new GroupChildEditPlan(context.Geometry,context.Constraints,groups.Keys,children.Select(ReadPosition),mode,keepAspectRatio);
            LastFrame=Plan.Initial;
        }

        public static GroupChildNativeEdit Capture(LayoutContext context,object[] children,GeometryEditMode mode,bool keepAspectRatio)
        {
            if(context==null)throw new ArgumentNullException("context");
            if(children==null)throw new ArgumentNullException("children");
            if(children.Length==0 || children.Any(child=>child==null || (int)((dynamic)child).Child==0 || ShapeCapabilities.IsLocked(child)))return null;
            var roots=children.Select(ShapeCapabilities.TopLevel).GroupBy(group=>(int)((dynamic)group).Id).Select(group=>group.First()).ToList();
            if(!roots.Any(group=>FindGroupId(context,group)!=null))return null;
            try
            {
                // PowerPoint's native ChildShapeRange cannot represent members
                // of distinct root groups. Keep the multi-root mathematical
                // plan independent, but reject this native request before any
                // Select call or geometry write. Free roots never enter Core.
                if(roots.Count!=1)
                    throw new InvalidOperationException("PowerPoint 不支持跨根组合的组内多选；请在同一组合内选择成员。");
                if(children.Select(child=>(int)((dynamic)child).Id).Distinct().Count()!=children.Length)
                    throw new InvalidOperationException("The native child selection contains a repeated object.");
                return new GroupChildNativeEdit(context,children.ToArray(),mode,keepAspectRatio);
            }
            catch(Exception error)
            {throw new GroupChildCaptureException("无法捕获已关联约束的组内编辑，操作已拒绝。",error);}
        }

        public SelectionMoveFrame Apply(IEnumerable<SelectionMovePosition> targets)
        {
            var accepted=Plan.ApplyAtomic(LastFrame,targets,WriteChildren,ReadChildren,ReadGroups,
                values=>PowerPointLayout.Apply(followers,values),()=>PowerPointLayout.Snapshot(Context));
            LastFrame=accepted;return accepted;
        }

        public void ThrowIfProvablyImpossible(IEnumerable<SelectionMovePosition> targets)
        {
            if(targets==null)throw new ArgumentNullException("targets");
            var requested=targets.ToList();
            if(requested.Any(target=>target==null) || requested.Count!=Children.Count ||
                requested.Select(target=>target.NativeId).Distinct().Count()!=Children.Count ||
                requested.Any(target=>!childrenById.ContainsKey(target.NativeId)))
                throw new ArgumentException("Preflight must retain every exact selected native member.","targets");
            if(preflightConstraints.Count==0)return;
            ReadGroups();
            var byId=requested.ToDictionary(target=>target.NativeId);
            var enclosed=new List<GroupChildBound>();
            foreach(var root in groups)
                if(preflightConstraints.Any(relation=>relation.ShapeIds[0]==root.Key)&&IsAxisAlignedWithout3D(root.Value))
                    AddEnclosedBounds(root.Value,root.Key,byId,enclosed);
            GroupChildPreflight.ThrowIfProvablyImpossible(preflightConstraints,enclosed);
        }

        public void Restore(SelectionMoveFrame frame)
        {
            if(frame==null)throw new ArgumentNullException("frame");
            WriteChildren(frame.Extras);
            PowerPointLayout.Apply(followers,frame.Geometry.Where(shape=>!groups.ContainsKey(shape.Id)));
            if(!ReadFrame().Matches(frame))throw new InvalidOperationException("The actual native group members and followers were not restored.");
            LastFrame=frame;
        }

        public SelectionMoveFrame ReadFrame()
        {ReadGroups();return new SelectionMoveFrame(PowerPointLayout.Snapshot(Context),ReadChildren());}

        public IReadOnlyList<SelectionMovePosition> ReadChildren()
        {return Children.Select(ReadPosition).ToList().AsReadOnly();}

        public IReadOnlyList<ShapeGeometry> ReadGroups()
        {
            foreach(object child in Children)
            {
                string id=childRoots[(int)((dynamic)child).Id];
                if((int)((dynamic)child).Child==0 || (int)((dynamic)ShapeCapabilities.TopLevel(child)).Id!=(int)((dynamic)groups[id]).Id)
                    throw new InvalidOperationException("An edited native member is no longer in its original root group.");
            }
            return groups.Select(pair=>PowerPointLayout.Geometry(pair.Value,pair.Key)).ToList().AsReadOnly();
        }

        public void SelectExact()
        {
            Select(groups.Values.ToArray());Select(Children);
            dynamic window=Context.Window;
            if(!(bool)window.Selection.HasChildShapeRange)
                throw new InvalidOperationException("PowerPoint did not select the exact native group members.");
            dynamic range=window.Selection.ChildShapeRange;
            var ids=Enumerable.Range(1,(int)range.Count).Select(index=>(int)range[index].Id).ToArray();
            if(ids.Length!=Children.Count || !new HashSet<int>(ids).SetEquals(childRoots.Keys))
                throw new InvalidOperationException("PowerPoint did not preserve every exact selected native group member.");
        }

        private void WriteChildren(IReadOnlyList<SelectionMovePosition> targets)
        {
            ReadGroups();
            foreach(var target in targets)
            {
                object native;
                if(!childrenById.TryGetValue(target.NativeId,out native))throw new InvalidOperationException("The native child identity changed.");
                dynamic child=native;
                bool resized=Math.Abs((double)child.Width-target.Width)>.0001 || Math.Abs((double)child.Height-target.Height)>.0001;
                int aspect=(int)child.LockAspectRatio;
                if(resized && aspect!=0)child.LockAspectRatio=0;
                try
                {
                    if(Math.Abs((double)child.Width-target.Width)>.0001)child.Width=(float)target.Width;
                    if(Math.Abs((double)child.Height-target.Height)>.0001)child.Height=(float)target.Height;
                    if(Math.Abs((double)child.Left-target.Left)>.0001)child.Left=(float)target.Left;
                    if(Math.Abs((double)child.Top-target.Top)>.0001)child.Top=(float)target.Top;
                }
                finally{if(resized && aspect!=0)child.LockAspectRatio=aspect;}
            }
        }

        private static void AddEnclosedBounds(object native,string rootId,Dictionary<int,SelectionMovePosition> targets,List<GroupChildBound> enclosed)
        {
            dynamic shape=native;
            int type=(int)shape.Type;
            if(type!=6&&!GroupChildPreflight.HasStableShapeBounds(type,false))return;
            if(!IsAxisAlignedWithout3D(native))return;
            SelectionMovePosition target;
            bool selected=targets.TryGetValue((int)shape.Id,out target);
            if(type==6)
            {
                // A selected nested group can rescale complex descendants.
                // Its target is deliberately not predicted, nor are its old
                // descendant coordinates treated as unchanged bounds.
                if(selected)return;
                for(int index=1;index<=(int)shape.GroupItems.Count;index++)
                    AddEnclosedBounds((object)shape.GroupItems[index],rootId,targets,enclosed);
                return;
            }
            bool autoSizesShape;
            try
            {
                autoSizesShape=false;
                if((int)shape.HasTextFrame!=0)
                {
                    int legacy=(int)shape.TextFrame.AutoSize,modern=(int)shape.TextFrame2.AutoSize;
                    autoSizesShape=legacy!=0 || (modern!=0&&modern!=2);
                }
            }
            catch(COMException){return;}
            if(!GroupChildPreflight.HasStableShapeBounds(type,autoSizesShape))return;
            enclosed.Add(new GroupChildBound(rootId,selected ? target : ReadPosition(native)));
        }

        private static bool IsAxisAlignedWithout3D(object native)
        {
            try
            {
                dynamic shape=native;dynamic threeD=shape.ThreeD;
                return GroupChildPreflight.IsAxisAlignedWithout3D((double)shape.Rotation,(int)threeD.Visible!=0,
                    (double)threeD.RotationX,(double)threeD.RotationY,(double)threeD.RotationZ);
            }
            // Office documents that some shape kinds reject ThreeD queries.
            // Such a rectangle supplies no proof; the real native edit remains
            // authoritative rather than being rejected by this optional gate.
            catch(COMException){return false;}
        }

        private static void Select(IReadOnlyList<object> shapes)
        {for(int index=0;index<shapes.Count;index++)((dynamic)shapes[index]).Select(index==0 ? -1 : 0);}
        private static SelectionMovePosition ReadPosition(object native)
        {dynamic shape=native;return new SelectionMovePosition((int)shape.Id,(double)shape.Left,(double)shape.Top,(double)shape.Width,(double)shape.Height);}
        private static string FindGroupId(LayoutContext context,object group)
        {int nativeId=(int)((dynamic)group).Id;return context.Shapes.Where(pair=>(int)((dynamic)pair.Value).Id==nativeId).Select(pair=>pair.Key).FirstOrDefault();}
    }
}
