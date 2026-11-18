using System;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    public enum NativeFieldProperty { Width,Height,Left,Top }
    public enum NativePositionOrigin { TopLeft,SlideCenter }

    // Input coordinates have already been translated from the observed native
    // reference origin. Native position edits retain size and the other axis.
    public static class NativeFieldTarget
    {
        public static bool IsDimension(NativeFieldProperty property)
        {return property==NativeFieldProperty.Width||property==NativeFieldProperty.Height;}
        public static double AbsolutePosition(NativeFieldProperty property,NativePositionOrigin origin,double displayedPoints,double slideWidth,double slideHeight)
        {
            if(property!=NativeFieldProperty.Left&&property!=NativeFieldProperty.Top)throw new ArgumentOutOfRangeException("property");
            if(Double.IsNaN(displayedPoints)||Double.IsInfinity(displayedPoints))throw new ArgumentOutOfRangeException("displayedPoints");
            if(origin==NativePositionOrigin.TopLeft)return displayedPoints;
            if(origin!=NativePositionOrigin.SlideCenter)throw new ArgumentOutOfRangeException("origin");
            // Independently verified for the horizontal and vertical fields:
            // center + 0 positions the bbox edge at half the slide dimension.
            double size=property==NativeFieldProperty.Left?slideWidth:slideHeight;
            if(Double.IsNaN(size)||Double.IsInfinity(size)||size<=0)throw new ArgumentOutOfRangeException("slideSize");
            return size/2+displayedPoints;
        }
        public static SelectionMovePosition Create(SelectionMovePosition before,NativeFieldProperty property,double points,bool aspect)
        {
            if(before==null)throw new ArgumentNullException("before");
            if(!Enum.IsDefined(typeof(NativeFieldProperty),property))throw new ArgumentOutOfRangeException("property");
            if(Double.IsNaN(points)||Double.IsInfinity(points)||Math.Abs(points)>Single.MaxValue)
                throw new ArgumentOutOfRangeException("points");
            if(property==NativeFieldProperty.Left)return new SelectionMovePosition(before.NativeId,points,before.Top,before.Width,before.Height);
            if(property==NativeFieldProperty.Top)return new SelectionMovePosition(before.NativeId,before.Left,points,before.Width,before.Height);
            if(points<LayoutSolver.MinimumSize)throw new ArgumentOutOfRangeException("points");
            bool width=property==NativeFieldProperty.Width;
            double newWidth=width?points:before.Width,newHeight=width?before.Height:points;
            if(aspect)
            {
                double old=width?before.Width:before.Height;
                if(old<LayoutSolver.MinimumSize)throw new InvalidOperationException("锁定宽高比的对象当前尺寸过小，无法推导原生缩放。");
                if(width)newHeight=before.Height*points/old;else newWidth=before.Width*points/old;
            }
            return new SelectionMovePosition(before.NativeId,before.Left,before.Top,newWidth,newHeight);
        }
    }
}
