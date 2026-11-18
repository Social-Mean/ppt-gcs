using System;
using System.Drawing;

namespace PptGcs.AddIn
{
    public static class MultiSelectionPointerPolicy
    {
        // true owns the current selection, false intentionally stays native,
        // null lets the caller handle a real hit on an unselected object.
        public static bool? RouteMove(RectangleF selectionBounds,PointF point,double handleRadius,
            float outlineWidth,bool? nativeHitSelected)
        {
            if(HitHandle(selectionBounds,point,handleRadius))return false;
            if(nativeHitSelected==false)return null;
            return CanMove(selectionBounds,point,handleRadius,outlineWidth,nativeHitSelected);
        }

        public static bool HitHandle(RectangleF bounds,PointF point,double radius)
        {
            if(bounds.Width==0 || bounds.Height==0)
            {
                double x1=point.X-bounds.Left,y1=point.Y-bounds.Top,x2=point.X-bounds.Right,y2=point.Y-bounds.Bottom;
                return x1*x1+y1*y1<=radius*radius || x2*x2+y2*y2<=radius*radius;
            }
            float[] xs={bounds.Left,(bounds.Left+bounds.Right)/2,bounds.Right};
            float[] ys={bounds.Top,(bounds.Top+bounds.Bottom)/2,bounds.Bottom};
            for(int x=0;x<3;x++)for(int y=0;y<3;y++)
            {
                if(x==1 && y==1)continue;
                double dx=point.X-xs[x],dy=point.Y-ys[y];
                if(dx*dx+dy*dy<=radius*radius)return true;
            }
            return false;
        }

        // nativeHitSelected: true = a real hit on a selected member, false =
        // another object, null = blank canvas. Only a narrow outline fills null.
        public static bool CanMove(RectangleF selectionBounds,PointF point,double handleRadius,
            float outlineWidth,bool? nativeHitSelected)
        {
            if(HitHandle(selectionBounds,point,handleRadius)||nativeHitSelected==false)return false;
            if(nativeHitSelected==true)return true;
            var outside=selectionBounds;outside.Inflate(outlineWidth,outlineWidth);
            float inset=Math.Min(outlineWidth,Math.Min(selectionBounds.Width,selectionBounds.Height)/4);
            var inside=selectionBounds;inside.Inflate(-inset,-inset);
            bool interior=point.X>inside.Left&&point.X<inside.Right&&point.Y>inside.Top&&point.Y<inside.Bottom;
            return point.X>=outside.Left&&point.X<=outside.Right&&point.Y>=outside.Top&&point.Y<=outside.Bottom&&!interior;
        }
    }
}
