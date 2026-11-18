using System;
using System.Collections.Generic;
using System.Linq;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    // Each supplied rectangle is known to be enclosed by this root. Their
    // envelope is only a lower bound on the root's size, never a predicted box.
    public sealed class GroupChildBound
    {
        public string RootId {get;private set;}
        public SelectionMovePosition Bounds {get;private set;}
        public GroupChildBound(string rootId,SelectionMovePosition bounds)
        {
            if(bounds==null)throw new ArgumentNullException("bounds");
            RootId=ShapeIdentity.Normalize(rootId);Bounds=bounds;
        }
    }

    public static class GroupChildPreflight
    {
        // Passing means no impossibility was proved. The native group bounds
        // must still be read and solved after the real child edit.
        public static void ThrowIfProvablyImpossible(IEnumerable<LayoutConstraint> relations,IEnumerable<GroupChildBound> enclosedBounds)
        {
            if(relations==null)throw new ArgumentNullException("relations");
            if(enclosedBounds==null)throw new ArgumentNullException("enclosedBounds");
            var constraints=relations.Where(relation=>relation.Enabled).ToList();
            var bounds=enclosedBounds.ToList();
            if(bounds.Any(bound=>bound==null))throw new ArgumentException("An enclosed bound cannot be null.","enclosedBounds");
            foreach(var root in bounds.GroupBy(bound=>bound.RootId,StringComparer.Ordinal))
            {
                var direct=constraints.Where(relation=>relation.ShapeIds.Count==1&&relation.ShapeIds[0]==root.Key).ToList();
                var positions=direct.Where(relation=>relation.Kind==ConstraintKind.FixedPosition).ToList();
                var widths=direct.Where(relation=>relation.Kind==ConstraintKind.FixedWidth).ToList();
                var heights=direct.Where(relation=>relation.Kind==ConstraintKind.FixedHeight).ToList();
                double left=root.Min(bound=>bound.Bounds.Left),top=root.Min(bound=>bound.Bounds.Top);
                double right=root.Max(bound=>bound.Bounds.Left+bound.Bounds.Width),bottom=root.Max(bound=>bound.Bounds.Top+bound.Bounds.Height);
                foreach(var position in positions)
                {
                    if(left<position.Value.Value-LayoutSolver.PositionTolerance || top<position.SecondaryValue.Value-LayoutSolver.PositionTolerance)
                        throw new ConstraintConflictException("组内对象的目标位置越过了组合固定的左边或顶边。");
                    foreach(var width in widths)
                        if(right>position.Value.Value+width.Value.Value+LayoutSolver.PositionTolerance)
                            throw new ConstraintConflictException("组内对象的目标位置超出了组合固定的宽度。");
                    foreach(var height in heights)
                        if(bottom>position.SecondaryValue.Value+height.Value.Value+LayoutSolver.PositionTolerance)
                            throw new ConstraintConflictException("组内对象的目标位置超出了组合固定的高度。");
                }
                foreach(var width in widths)
                    if(right-left>width.Value.Value+LayoutSolver.PositionTolerance)
                        throw new ConstraintConflictException("组内对象所需的水平跨度超过了组合固定的宽度。");
                foreach(var height in heights)
                    if(bottom-top>height.Value.Value+LayoutSolver.PositionTolerance)
                        throw new ConstraintConflictException("组内对象所需的垂直跨度超过了组合固定的高度。");
            }
        }

        public static bool IsAxisAlignedWithout3D(double rotation,bool threeDVisible,double rotationX,double rotationY,double rotationZ)
        {
            // A conservative gate: tolerance is for COM zero noise, not an
            // approximation to a rotated bounding box.
            return !threeDVisible&&ZeroAngle(rotation)&&ZeroAngle(rotationX)&&ZeroAngle(rotationY)&&ZeroAngle(rotationZ);
        }

        public static bool HasStableShapeBounds(int shapeType,bool autoSizesShape)
        {
            if(autoSizesShape)return false;
            switch(shapeType)
            {
                case 1: case 5: case 11: case 13: case 17: case 28: case 29:return true;
                default:return false;
            }
        }

        private static bool ZeroAngle(double value)
        {
            if(Double.IsNaN(value)||Double.IsInfinity(value))return false;
            double normalized=value%360;
            return Math.Min(Math.Abs(normalized),360-Math.Abs(normalized))<=.000001;
        }
    }
}
