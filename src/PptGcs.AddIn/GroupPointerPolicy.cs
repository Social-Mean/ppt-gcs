using System;
using System.Drawing;

namespace PptGcs.AddIn
{
    // A group body click and drag start at the same native child hit. Own the
    // candidate gesture, then select that child only if the pointer did not drag.
    public static class GroupPointerPolicy
    {
        public static bool HasDragDistance(double dx,double dy,double threshold)
        {return Math.Abs(dx)>=threshold || Math.Abs(dy)>=threshold;}

        public static bool AcceptNativeHit(int shapeType,bool hitBodyOrBorder)
        {
            // The caller has already matched RangeFromPoint's top-level owner.
            // Native group hits can extend beyond the reported COM bounds.
            return shapeType==6 || hitBodyOrBorder;
        }

        public static bool ShouldSelectChild(bool wasSelectedGroup,bool exactChild,bool moveMode,
            bool dragThresholdPassed,bool cancelled,bool normalRelease)
        {
            return wasSelectedGroup && exactChild && moveMode && !dragThresholdPassed && !cancelled && normalRelease;
        }

        public static bool HitMoveBorder(int shapeType,bool selected,
            RectangleF bounds,PointF point,float borderWidth)
        {
            if(shapeType!=6 || !selected)return false;
            var outside=bounds;outside.Inflate(borderWidth,borderWidth);
            float border=Math.Min(borderWidth,Math.Min(bounds.Width,bounds.Height)/4);
            var inside=bounds;inside.Inflate(-border,-border);
            bool interior=point.X>inside.Left && point.X<inside.Right && point.Y>inside.Top && point.Y<inside.Bottom;
            return point.X>=outside.Left && point.X<=outside.Right && point.Y>=outside.Top && point.Y<=outside.Bottom &&
                !interior;
        }
    }
}
