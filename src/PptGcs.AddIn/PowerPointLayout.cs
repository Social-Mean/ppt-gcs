using System;
using System.Collections.Generic;
using System.Linq;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    public sealed class LayoutContext
    {
        public object Window;
        public object Presentation;
        public object Slide;
        public int SlideId;
        public ConstraintDocument Document;
        public Dictionary<string, object> Shapes = new Dictionary<string, object>();
        public List<ShapeGeometry> Geometry = new List<ShapeGeometry>();
        public List<LayoutConstraint> Constraints = new List<LayoutConstraint>();
    }

    public static class PowerPointLayout
    {
        public const string TagName = "PPTGCS_ID";

        public static ConstraintDocument ReadDocument(object presentationObject)
        {
            return NativeConstraintStore.Read(presentationObject);
        }

        public static void WriteDocument(object presentationObject, ConstraintDocument document)
        {
            NativeConstraintStore.Write(presentationObject, document);
        }

        public static void ValidateShape(object shapeObject)
        {
            dynamic shape = shapeObject;
            if ((int)shape.Type != 1 || (int)shape.AutoShapeType != 1 || Math.Abs((double)shape.Rotation) > 0.01)
                throw new InvalidOperationException("当前支持未旋转、未分组的普通矩形。");
        }

        public static LayoutContext Read(object applicationObject)
        {
            dynamic application = applicationObject;
            if ((int)application.Presentations.Count == 0) throw new InvalidOperationException("请先打开演示文稿。");
            dynamic window = application.ActiveWindow;
            if (window == null || (int)window.ViewType != 9) throw new InvalidOperationException("请切换到普通幻灯片编辑视图。");
            dynamic slide = window.View.Slide;
            var context = new LayoutContext
            {
                Window = window, Presentation = window.Presentation, Slide = slide,
                SlideId = (int)slide.SlideID,
                Document = ReadDocument((object)window.Presentation)
            };
            context.Constraints = context.Document.Constraints.Where(c => c.SlideId == context.SlideId && c.Enabled).ToList();
            var required = new HashSet<string>(context.Constraints.SelectMany(c => c.ShapeIds));
            for (int i = 1; i <= (int)slide.Shapes.Count; i++)
            {
                dynamic shape = slide.Shapes[i];
                string id = ReadId((object)shape);
                if (id == null || !required.Contains(id)) continue;
                ValidateShape((object)shape);
                if (context.Shapes.ContainsKey(id)) throw new InvalidOperationException("矩形标识重复，请移除复制对象的约束后重新建立。");
                context.Shapes.Add(id, (object)shape);
                context.Geometry.Add(Geometry((object)shape, id));
            }
            if (required.Any(id => !context.Shapes.ContainsKey(id)))
                throw new InvalidOperationException("约束引用的矩形已删除或不在此页，请移除失效约束。");
            return context;
        }

        public static string ReadId(object shapeObject)
        {
            dynamic shape = shapeObject;
            string value = (string)shape.Tags[TagName];
            Guid guid;
            return Guid.TryParse(value, out guid) && guid != Guid.Empty ? guid.ToString("D").ToUpperInvariant() : null;
        }

        public static ShapeGeometry Geometry(object shapeObject, string id)
        {
            dynamic shape = shapeObject;
            return new ShapeGeometry(id, (double)shape.Left, (double)shape.Top, (double)shape.Width, (double)shape.Height);
        }

        public static List<ShapeGeometry> Snapshot(LayoutContext context)
        {
            return context.Shapes.Select(pair => Geometry(pair.Value, pair.Key)).ToList();
        }

        public static void Apply(LayoutContext context, IEnumerable<ShapeGeometry> values)
        {
            var target = values.ToList();
            var before = Snapshot(context);
            try { Write(context, target); }
            catch
            {
                // Roll back the complete geometry if Office rejects any write.
                Write(context, before);
                throw;
            }
        }

        private static void Write(LayoutContext context, IEnumerable<ShapeGeometry> values)
        {
            foreach (var value in values)
            {
                object shapeObject;
                if (!context.Shapes.TryGetValue(value.Id, out shapeObject)) throw new InvalidOperationException("几何写回找不到目标矩形。");
                dynamic shape = shapeObject;
                ShapeGeometry current = Geometry(shapeObject, value.Id);
                if (Math.Abs(current.Left-value.Left)<=0.0001 && Math.Abs(current.Top-value.Top)<=0.0001 &&
                    Math.Abs(current.Width-value.Width)<=0.0001 && Math.Abs(current.Height-value.Height)<=0.0001) continue;
                int aspect = (int)shape.LockAspectRatio;
                bool sizeChanged = Math.Abs(current.Width-value.Width)>0.0001 || Math.Abs(current.Height-value.Height)>0.0001;
                if (sizeChanged && aspect != 0) shape.LockAspectRatio = 0;
                try
                {
                    if (Math.Abs(current.Width-value.Width)>0.0001) shape.Width = (float)value.Width;
                    if (Math.Abs(current.Height-value.Height)>0.0001) shape.Height = (float)value.Height;
                    if (Math.Abs(current.Left-value.Left)>0.0001) shape.Left = (float)value.Left;
                    if (Math.Abs(current.Top-value.Top)>0.0001) shape.Top = (float)value.Top;
                }
                finally { if (sizeChanged && aspect != 0) shape.LockAspectRatio = aspect; }
            }
        }
    }
}
