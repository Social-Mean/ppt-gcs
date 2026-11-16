using System;
using System.Drawing;

namespace PptGcs.AddIn
{
    // Match PowerPoint text-box interaction: native clicks in the interior,
    // constrained movement on the selected border. Resize handles take priority.
    public static class TextPointerPolicy
    {
        public static bool UseNativeTextInput(int shapeType, bool editingText,
            RectangleF bounds, PointF point, float borderWidth)
        {
            if (shapeType != 17 && !editingText) return false;
            float border = Math.Min(borderWidth, Math.Min(bounds.Width, bounds.Height) / 4);
            bounds.Inflate(-border, -border);
            return point.X > bounds.Left && point.X < bounds.Right && point.Y > bounds.Top && point.Y < bounds.Bottom;
        }

        public static bool HitMoveBorder(int shapeType, bool editingText,
            RectangleF bounds, PointF point, float borderWidth)
        {
            if (shapeType != 17 && !editingText) return false;
            var outside = bounds;
            outside.Inflate(borderWidth, borderWidth);
            return point.X >= outside.Left && point.X <= outside.Right && point.Y >= outside.Top && point.Y <= outside.Bottom &&
                !UseNativeTextInput(shapeType, editingText, bounds, point, borderWidth);
        }

        // A selection outline may be outside the native object hit area. Never
        // steal an overlapping object's hit, or guess an unselected blank hit.
        public static bool CanUseBorderHit(bool selected, bool? nativeHitMatches)
        {
            return nativeHitMatches == true || (selected && !nativeHitMatches.HasValue);
        }
    }
}
