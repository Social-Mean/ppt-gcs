using System;
using System.Collections.Generic;
using System.Linq;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    internal static class PageResize
    {
        public const int MinimumPoints=72,MaximumPoints=4032;

        public static void Apply(object applicationObject,double width,double height)
        {
            if(!(width>=MinimumPoints && width<=MaximumPoints && height>=MinimumPoints && height<=MaximumPoints))
                throw new ArgumentOutOfRangeException("size","幻灯片的宽和高必须在 72 到 4032 磅之间。");
            dynamic application=applicationObject;
            if((int)application.Presentations.Count==0)throw new InvalidOperationException("请先打开演示文稿。");
            dynamic presentation=application.ActivePresentation;
            dynamic page=presentation.PageSetup;
            double oldWidth=(double)page.SlideWidth,oldHeight=(double)page.SlideHeight;
            int oldSize=(int)page.SlideSize;
            double newWidth=(float)width,newHeight=(float)height;
            var document=PowerPointLayout.ReadDocument((object)presentation);
            var affected=new HashSet<int>(document.Constraints.Where(item=>item.Enabled).Select(item=>item.SlideId));
            var contexts=new List<LayoutContext>();
            bool changing=newWidth!=oldWidth || newHeight!=oldHeight;
            bool reflow=false;
            for(int i=1;i<=(int)presentation.Slides.Count;i++)
            {
                dynamic slide=presentation.Slides[i];
                if(!affected.Contains((int)slide.SlideID))continue;
                var context=PowerPointLayout.ReadSlide((object)presentation,(object)slide,null,document);
                contexts.Add(context);
                // The native setters recenter old content about the new page.
                // Preflight this geometry before touching the document.
                var predicted=context.Geometry.Select(item=>new ShapeGeometry(item.Id,
                    item.Left+(newWidth-oldWidth)/2,item.Top+(newHeight-oldHeight)/2,item.Width,item.Height)).ToList();
                var result=Solve(context,predicted,newWidth,newHeight);
                reflow|=predicted.Zip(result,LayoutSolver.NeedsUpdate).Any(changed=>changed);
            }
            if(!changing && !reflow)return;

            application.StartNewUndoEntry();
            try
            {
                SetSize((object)page,newWidth,newHeight);
                foreach(var context in contexts)
                {
                    // Native readback is authoritative, including group bounds.
                    var after=PowerPointLayout.Snapshot(context);
                    PowerPointLayout.Apply(context,Solve(context,after,newWidth,newHeight));
                    var bound=PowerPointLayout.BindRelations(context.SlideId,document,newWidth,newHeight);
                    if(LayoutSolver.MaxResidual(PowerPointLayout.Snapshot(context),bound)>LayoutSolver.PositionTolerance)
                        throw new InvalidOperationException("页面调整后的几何未满足约束。");
                }
            }
            catch(Exception failure)
            {
                var errors=new List<Exception> {failure};
                try
                {
                    if((int)page.SlideSize!=oldSize && oldSize!=7)page.SlideSize=oldSize;
                    SetSize((object)page,oldWidth,oldHeight);
                }
                catch(Exception error){errors.Add(error);}
                foreach(var context in contexts)
                    try{PowerPointLayout.Apply(context,context.Geometry);}catch(Exception error){errors.Add(error);}
                if(errors.Count>1)throw new AggregateException("页面调整失败，恢复时也发生错误。",errors);
                throw;
            }
            finally {application.StartNewUndoEntry();}
        }

        private static IReadOnlyList<ShapeGeometry> Solve(LayoutContext context,IReadOnlyList<ShapeGeometry> geometry,double width,double height)
        {
            var bound=PowerPointLayout.BindRelations(context.SlideId,context.Document,width,height);
            try{return LayoutSolver.Normalize(geometry,bound);}
            catch(ConstraintConflictException error)
            {
                var witness=LayoutDiagnostics.ExplainConflict(geometry,bound);
                int page=(int)((dynamic)context.Slide).SlideIndex;
                string description=String.Join(Environment.NewLine,witness.Select(item=>ConstraintCommands.DescribeRelation(context,item)));
                throw new ConstraintConflictException("第 "+page+" 页的约束无法满足所选幻灯片大小。"+Environment.NewLine+description,witness,error);
            }
        }

        private static void SetSize(object pageObject,double width,double height)
        {
            dynamic page=pageObject;
            if((double)page.SlideWidth!=width)page.SlideWidth=(float)width;
            if((double)page.SlideHeight!=height)page.SlideHeight=(float)height;
            if(Math.Abs((double)page.SlideWidth-width)>LayoutSolver.PositionTolerance ||
                Math.Abs((double)page.SlideHeight-height)>LayoutSolver.PositionTolerance)
                throw new InvalidOperationException("PowerPoint 未保留请求的页面尺寸。");
        }
    }
}
