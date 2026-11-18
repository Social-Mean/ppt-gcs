using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Linq;
using System.Runtime.InteropServices;
using PptGcs.AddIn;
using PptGcs.Core;

internal static class CanvasConstraintTests
{
    private static int passed;
    private static readonly Size Viewport = new Size(900, 500);

    [STAThread] private static int Main()
    {
        try
        {
            FilterAndValues();
            LabelFits();
            ProjectionAndVisibility();
            PlacementAndState();
            BadgeInteractionBounds();
            RenderAlpha();
            Console.WriteLine("PASS: " + passed + " canvas constraint tests (no Office, windows or input)");
            return 0;
        }
        catch(Exception error){Console.Error.WriteLine(error);return 1;}
    }

    private static void BadgeInteractionBounds()
    {
        var model = Model();
        model.Relations.Add(Relation("BC", "同高", "B", "C"));
        var projection = new CanvasConstraintProjection { ScaleX=1, ScaleY=1 };
        var layout = CanvasConstraintRenderer.Build(model,CanvasConstraintDisplay.All,projection,Viewport,"AB");
        var target = layout.Badges.Single(badge => badge.RelationId=="AB");
        var anchor = Point.Round(new PointF(target.Bounds.Left+target.Bounds.Width/2, target.Bounds.Top+target.Bounds.Height/2));
        Check(layout.BadgeAt(anchor)==target && layout.Badges.Count(badge=>badge.Highlighted)==1 && target.Highlighted,
            "only the addressed relation owns its badge point and highlight");
        var line = layout.Lines.First();
        var linePoint = Point.Round(new PointF((line.Start.X+line.End.X)/2,(line.Start.Y+line.End.Y)/2));
        Check(layout.BadgeAt(linePoint)==null && layout.BadgeAt(Point.Empty)==null,
            "leader lines and canvas outside the badge remain unclaimed");
        var limited=CanvasConstraintRenderer.Build(model,CanvasConstraintDisplay.Selection,projection,Viewport,null,new[]{"BC"});
        Check(limited.Badges.Any(badge=>badge.RelationId=="BC" && badge.Limited && badge.Color.R>badge.Color.B) &&
            limited.Badges.All(badge=>badge.RelationId=="BC" || !badge.Limited),
            "an indirect blocker is visible and distinguished in selected-object mode");
    }

    private static void LabelFits()
    {
        using(var image=new Bitmap(1,1))
        using(var graphics=Graphics.FromImage(image))
        using(var font=new Font(SystemFonts.MessageBoxFont.FontFamily,12,FontStyle.Regular,GraphicsUnit.Pixel))
        using(var format=new StringFormat {Alignment=StringAlignment.Near,LineAlignment=StringAlignment.Center,
            Trimming=StringTrimming.EllipsisCharacter,FormatFlags=StringFormatFlags.NoWrap})
        {
            graphics.TextRenderingHint=TextRenderingHint.AntiAliasGridFit;
            Check(new[]{"同宽","顶端对齐","水平等间距"}.All(label=>{
                var model=Model();model.Relations[0].Label=label;
                var badge=Build(model).Badges.Single();
                return badge.Bounds.Width-28>=graphics.MeasureString(label,font,PointF.Empty,format).Width;
            }),"common Chinese labels fit the actual text area after reserving the icon and all padding");
        }
    }

    private static void FilterAndValues()
    {
        var model=Model();
        model.Relations.Add(Relation("BC", "同高", "B", "C"));
        model.Relations.Add(new ConstraintPaneRelation {Id="Fixed",Kind="FixedPosition",Label="固定位置",
            MemberIds=new[]{"A"},Value=125.5,Second=-20,Status="已满足"});
        var selected=Build(model);
        Check(selected.Badges.Select(b=>b.RelationId).SequenceEqual(new[]{"AB","Fixed"}),
            "selection mode includes only relations connected through object UI keys");
        Check(Build(model,CanvasConstraintDisplay.All).Badges.Count==3 &&
            Build(model,CanvasConstraintDisplay.Hidden).Badges.Count==0,
            "all and hidden modes display all placeable relations or no annotations");
        Check(selected.Badges.Single(b=>b.RelationId=="Fixed").Label.Contains("125.5, -20"),
            "both fixed parameters appear beside the relation name");
        model.Objects[0].Selected=false;
        Check(Build(model).Badges.Count==0 && selected.Badges.Count==2,
            "clearing selection hides badges without changing an existing display snapshot");
    }

    private static void ProjectionAndVisibility()
    {
        var model=Model();
        var projection=new CanvasConstraintProjection {OriginX=10,OriginY=20,ScaleX=2,ScaleY=3,DpiScale=2};
        var result=CanvasConstraintRenderer.Build(model,CanvasConstraintDisplay.Selection,projection,new Size(1400,1000));
        var badge=result.Badges.Single();
        Check(Near(badge.Bounds.Left,10+(60+80)*2+8*2) && Near(badge.Bounds.Top,20+80*3) && Near(badge.Bounds.Height,24*2),
            "shape coordinates use the projection once while marker clearance and height follow DPI");

        model.Objects[0].Geometry=Geometry(2000,80,80,50);
        model.Objects[1].Geometry=Geometry(2400,80,80,50);
        Check(Build(model).Badges.Count==0 && Build(model).Lines.Count==0,
            "fully offscreen relations do not accumulate clamped badges at viewport corners");
        model.Objects[0].Geometry=Geometry(-600,80,80,50);
        model.Objects[1].Geometry=Geometry(320,80,80,50);
        result=Build(model);
        Check(result.Badges.Count==1 && Near(result.Badges[0].Bounds.Left,320+80+8),
            "a visible member becomes the anchor when the selected member is offscreen");
        Check(result.Lines.Count==2 && result.Lines.All(line=>Inside(line.Start,Viewport) && Inside(line.End,Viewport)) &&
            result.Lines.Any(line=>Near(line.Start.X,0) || Near(line.End.X,0)),
            "leaders to offscreen members stop on the viewport boundary");

        model.Objects[0].Geometry=null;
        result=Build(model);
        Check(result.Badges.Count==1 && IsRed(result.Badges[0].Color),
            "a missing member still reports a red relation at its known visible partner");
        model.Objects[1].Geometry=null;
        Check(Build(model).Badges.Count==0,"relations with no known geometric anchor are skipped");
    }

    private static void PlacementAndState()
    {
        var model=Model();model.Relations.Clear();
        for(int index=0;index<8;index++)model.Relations.Add(Relation("R"+index,"同宽","A","B","C"));
        var result=Build(model);
        Check(result.Badges.Count==8 && result.Lines.Count==24,
            "multiple relations stack independently and retain leaders to every member");
        Check(result.Badges.All(b=>Within(b.Bounds,Viewport)) && result.Badges.SelectMany((first,index)=>
            result.Badges.Skip(index+1).Select(second=>!first.Bounds.IntersectsWith(second.Bounds))).All(value=>value),
            "stacked badges remain fully in the viewport without overlapping each other");
        var protectedShape=new RectangleF(60,80,80,50);protectedShape.Inflate(8,8);
        Check(result.Badges.All(b=>!b.Bounds.IntersectsWith(protectedShape)),
            "badge placement leaves the object text and selection-handle clearance unobscured");
        Check(result.Badges.Select(b=>b.Bounds).SequenceEqual(Build(model).Badges.Select(b=>b.Bounds)),
            "identical snapshots keep a stable badge arrangement");

        model.Relations[0].Enabled=false;
        model.Relations[1].Status="未满足";
        result=Build(model);
        Check(IsGray(result.Badges[0].Color) && IsRed(result.Badges[1].Color) &&
            result.Badges[2].Color.B>result.Badges[2].Color.R && result.Lines.All(line=>line.Color.A<160),
            "disabled, unsatisfied and satisfied badges have distinct states and subdued leaders");

        var crowded=new ConstraintPaneModel();crowded.Objects.Add(Object("A",0,0,180,100,true));
        crowded.Relations.Add(Relation("Large","固定宽度","A"));
        Check(CanvasConstraintRenderer.Build(crowded,CanvasConstraintDisplay.All,Projection(),new Size(180,100)).Badges.Count==0,
            "an annotation is omitted when it cannot fit outside the object instead of covering its content");

        model=Model();model.Objects[0].Geometry=Geometry(800,420,80,50);
        result=Build(model);
        Check(result.Badges.Count==1 && Within(result.Badges[0].Bounds,Viewport) && result.Badges[0].Bounds.Right<=800-8,
            "objects near the bottom-right edge place their badges safely on the other side");
    }

    private static void RenderAlpha()
    {
        var layout=Build(Model());
        using(var bitmap=new Bitmap(Viewport.Width,Viewport.Height,PixelFormat.Format32bppPArgb))
        {
            using(var graphics=Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Transparent);
                CanvasConstraintRenderer.Draw(graphics,layout,new Dictionary<string,Image>(),1);
            }
            Check(bitmap.GetPixel(0,0).A==0 && bitmap.GetPixel(Viewport.Width-1,Viewport.Height-1).A==0 &&
                bitmap.GetPixel((int)layout.Badges[0].Bounds.Left+2,(int)layout.Badges[0].Bounds.Top+2).A>0,
                "drawing changes annotation pixels while the rest of the HUD stays transparent");
            bool valid=true;int translucent=0;
            var bits=bitmap.LockBits(new Rectangle(Point.Empty,bitmap.Size),ImageLockMode.ReadOnly,PixelFormat.Format32bppPArgb);
            try
            {
                var row=new byte[bitmap.Width*4];
                for(int y=0;y<bitmap.Height;y++)
                {
                    Marshal.Copy(IntPtr.Add(bits.Scan0,y*bits.Stride),row,0,row.Length);
                    for(int x=0;x<bitmap.Width;x++)
                    {
                        int offset=x*4,alpha=row[offset+3];if(alpha>0 && alpha<255)translucent++;
                        valid&=row[offset]<=alpha && row[offset+1]<=alpha && row[offset+2]<=alpha;
                    }
                }
            }
            finally{bitmap.UnlockBits(bits);}
            Check(valid && translucent>0,"labels and leaders produce valid premultiplied pixels for the layered surface");
        }
    }

    private static ConstraintPaneModel Model()
    {
        var model=new ConstraintPaneModel();
        model.Objects.Add(Object("A",60,80,80,50,true));
        model.Objects.Add(Object("B",320,80,80,50,false));
        model.Objects.Add(Object("C",580,80,80,50,false));
        model.Relations.Add(Relation("AB","左对齐","A","B"));return model;
    }
    private static ConstraintPaneObject Object(string id,double x,double y,double width,double height,bool selected)
    {return new ConstraintPaneObject {Id=id,Name=id,Selected=selected,Geometry=Geometry(x,y,width,height)};}
    private static ShapeGeometry Geometry(double x,double y,double width,double height)
    {return new ShapeGeometry(Guid.NewGuid().ToString("D"),x,y,width,height);}
    private static ConstraintPaneRelation Relation(string id,string label,params string[] members)
    {return new ConstraintPaneRelation {Id=id,Kind="SameWidth",Label=label,MemberIds=members,Status="已满足"};}
    private static CanvasConstraintProjection Projection()
    {return new CanvasConstraintProjection {ScaleX=1,ScaleY=1};}
    private static CanvasConstraintLayout Build(ConstraintPaneModel model,CanvasConstraintDisplay display=CanvasConstraintDisplay.Selection)
    {return CanvasConstraintRenderer.Build(model,display,Projection(),Viewport);}
    private static bool Near(float value,double expected){return Math.Abs(value-expected)<0.02;}
    private static bool Inside(PointF point,Size size){return point.X>=0 && point.Y>=0 && point.X<=size.Width && point.Y<=size.Height;}
    private static bool Within(RectangleF bounds,Size size){return bounds.Left>=0 && bounds.Top>=0 && bounds.Right<=size.Width && bounds.Bottom<=size.Height;}
    private static bool IsRed(Color color){return color.R>color.G*1.5 && color.R>color.B*1.5;}
    private static bool IsGray(Color color){return Math.Max(color.R,Math.Max(color.G,color.B))-Math.Min(color.R,Math.Min(color.G,color.B))<25;}
    private static void Check(bool condition,string description)
    {if(!condition)throw new InvalidOperationException(description);passed++;Console.WriteLine("PASS "+description);}
}
