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
        public double PageWidth,PageHeight;
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
            ShapeCapabilities.Validate(shapeObject);
        }

        public static LayoutContext Read(object applicationObject)
        {
            dynamic application = applicationObject;
            if ((int)application.Presentations.Count == 0) throw new InvalidOperationException("请先打开演示文稿。");
            return ReadWindow((object)application.ActiveWindow);
        }

        public static LayoutContext ReadWindow(object windowObject)
        {
            dynamic window = windowObject;
            if (window == null || (int)window.ViewType != 9) throw new InvalidOperationException("请切换到普通幻灯片编辑视图。");
            dynamic slide = window.View.Slide;
            object presentation = (object)window.Presentation;
            return ReadSlide(presentation,(object)slide,windowObject,ReadDocument(presentation));
        }

        internal static LayoutContext NewContext(object presentation,object slide,object window,ConstraintDocument document)
        {
            dynamic page=((dynamic)presentation).PageSetup;
            var context=new LayoutContext {Window=window,Presentation=presentation,Slide=slide,
                SlideId=(int)((dynamic)slide).SlideID,Document=document,PageWidth=(double)page.SlideWidth,PageHeight=(double)page.SlideHeight};
            context.Constraints=BindRelations(context,document);
            return context;
        }

        internal static List<LayoutConstraint> BindRelations(LayoutContext context,ConstraintDocument document)
        { return BindRelations(context.SlideId,document,context.PageWidth,context.PageHeight); }

        internal static List<LayoutConstraint> BindRelations(int slideId,ConstraintDocument document,double width,double height)
        { return document.Constraints.Where(c=>c.SlideId==slideId && c.Enabled).Select(c=>c.BindPageSize(width,height)).ToList(); }

        internal static LayoutContext ReadSlide(object presentation,object slideObject,object window,ConstraintDocument document)
        {
            var context=NewContext(presentation,slideObject,window,document);
            dynamic slide=slideObject;
            var required = new HashSet<string>(context.Constraints.SelectMany(c => c.ShapeIds));
            for (int i = 1; i <= (int)slide.Shapes.Count; i++)
            {
                dynamic shape = slide.Shapes[i];
                string id = ReadId((object)shape);
                if (id == null || !required.Contains(id)) continue;
                ValidateShape((object)shape);
                if (context.Shapes.ContainsKey(id)) throw new InvalidOperationException("对象标识重复，请移除复制对象的约束后重新建立。");
                context.Shapes.Add(id, (object)shape);
                context.Geometry.Add(Geometry((object)shape, id));
            }
            if (required.Any(id => !context.Shapes.ContainsKey(id)))
                throw new InvalidOperationException("约束引用的对象已删除或不在此页，请移除失效约束。");
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
                if (!context.Shapes.TryGetValue(value.Id, out shapeObject)) throw new InvalidOperationException("几何写回找不到目标对象。");
                dynamic shape = shapeObject;
                ShapeGeometry current = Geometry(shapeObject, value.Id);
                if (Math.Abs(current.Left-value.Left)<=0.0001 && Math.Abs(current.Top-value.Top)<=0.0001 &&
                    Math.Abs(current.Width-value.Width)<=0.0001 && Math.Abs(current.Height-value.Height)<=0.0001) continue;
                if (ShapeCapabilities.IsLocked(shapeObject))
                {
                    if (!LayoutSolver.NeedsUpdate(current, value)) continue;
                    throw new InvalidOperationException("此约束编辑需要移动或调整已锁定的对象。");
                }
                int aspect = (int)shape.LockAspectRatio;
                bool sizeChanged = Math.Abs(current.Width-value.Width)>0.0001 || Math.Abs(current.Height-value.Height)>0.0001;
                if (sizeChanged && aspect != 0) shape.LockAspectRatio = 0;
                try
                {
                    if (Math.Abs(current.Width-value.Width)>0.0001) shape.Width = (float)value.Width;
                    // Text AutoSize and group resizing can change other bounds
                    // as a side effect. Compare against Office's current values.
                    if (Math.Abs((double)shape.Height-value.Height)>0.0001) shape.Height = (float)value.Height;
                    if (Math.Abs((double)shape.Left-value.Left)>0.0001) shape.Left = (float)value.Left;
                    if (Math.Abs((double)shape.Top-value.Top)>0.0001) shape.Top = (float)value.Top;
                }
                finally { if (sizeChanged && aspect != 0) shape.LockAspectRatio = aspect; }
                if (LayoutSolver.NeedsUpdate(Geometry(shapeObject,value.Id),value))
                    throw new InvalidOperationException("PowerPoint 未保留请求的对象位置或尺寸，操作已取消。");
            }
        }
    }
}
