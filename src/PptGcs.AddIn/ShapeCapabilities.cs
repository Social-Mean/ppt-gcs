using System;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    // Layout constraints use the unrotated Office bounding box, irrespective of
    // the outline inside it. A group is one object; its children remain native.
    public static class ShapeCapabilities
    {
        public static bool IsSupported(object shape)
        { return UnsupportedReason(shape) == null; }

        public static void Validate(object shape)
        {
            string reason = UnsupportedReason(shape);
            if (reason != null) throw new InvalidOperationException(reason);
        }

        public static string UnsupportedReason(object shapeObject)
        {
            if (shapeObject == null) return "请选择形状对象。";
            dynamic shape = shapeObject;
            switch ((int)shape.Type)
            {
                case 1:  // msoAutoShape
                case 5:  // msoFreeform
                case 6:  // msoGroup (as a whole)
                case 11: // msoLinkedPicture
                case 13: // msoPicture
                case 17: // msoTextBox
                case 28: // msoGraphic (SVG / Office icons)
                case 29: // msoLinkedGraphic
                    break;
                default:
                    return "此对象类型暂不支持布局约束。支持自选图形、文本框、图片、SVG 图标、自由曲线和组合整体；线条、连接符、表格、图表及 SmartArt 暂不支持。";
            }
            double rotation = (double)shape.Rotation;
            double normalized = rotation % 360;
            if (!Finite(rotation) || Math.Min(Math.Abs(normalized), 360 - Math.Abs(normalized)) > 0.01)
                return "请先将对象旋转角度设为 0°，再建立布局约束。";
            double width = (double)shape.Width, height = (double)shape.Height;
            if (!Finite((double)shape.Left) || !Finite((double)shape.Top) ||
                !Finite(width) || !Finite(height) || width < LayoutSolver.MinimumSize || height < LayoutSolver.MinimumSize)
                return "对象的宽度和高度必须至少为 1 磅，坐标及尺寸必须有效。";
            return null;
        }

        public static object TopLevel(object shapeObject)
        {
            dynamic shape = shapeObject;
            while ((int)shape.Child != 0) shape = shape.ParentGroup;
            return (object)shape;
        }

        internal static bool IsLocked(object shapeObject)
        {
            dynamic shape = shapeObject;
            while (true)
            {
                if (ReadNativeLock((object)shape).GetValueOrDefault() != 0) return true;
                if ((int)shape.Child == 0) return false;
                shape = shape.ParentGroup;
            }
        }

        internal static int? ReadNativeLock(object shape)
        {
            const int memberNotFound = unchecked((int)0x80020003);
            const int unknownName = unchecked((int)0x80020006);
            const int notImplemented = unchecked((int)0x80004001);
            // Some PowerPoint builds expose Locked but return E_NOTIMPL from its getter.
            object value;
            try { value = ((dynamic)shape).Locked; }
            catch (RuntimeBinderException) { return null; }
            catch (NotImplementedException) { return null; }
            catch (COMException error) when (error.ErrorCode == memberNotFound || error.ErrorCode == unknownName || error.ErrorCode == notImplemented) { return null; }
            return (int)(dynamic)value;
        }

        private static bool Finite(double value)
        { return !Double.IsNaN(value) && !Double.IsInfinity(value); }
    }
}
