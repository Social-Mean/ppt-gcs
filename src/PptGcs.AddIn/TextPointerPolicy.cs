using System;
using System.Drawing;

namespace PptGcs.AddIn
{
    // Match PowerPoint text-box interaction: native clicks in the interior,
    // constrained movement on the selected border. Resize handles take priority.
    public static class TextPointerPolicy
    {
        public static bool UseNativeTextInput(int shapeType, bool selected, bool editingText,
            RectangleF bounds, PointF point, float borderWidth)
        {
            if (shapeType != 17 && !editingText) return false;
            if (!bounds.Contains(point)) return false;
            if (!selected) return true;
            float border = Math.Min(borderWidth, Math.Min(bounds.Width, bounds.Height) / 4);
            bounds.Inflate(-border, -border);
            return bounds.Contains(point);
        }
    }
}
