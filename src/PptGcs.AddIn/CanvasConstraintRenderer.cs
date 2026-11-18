using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;
using System.Linq;

namespace PptGcs.AddIn
{
    public enum CanvasConstraintDisplay { Selection, All, Hidden }

    public sealed class CanvasConstraintProjection
    {
        // Origin is relative to the HUD; the scale already includes slide zoom and DPI.
        public double OriginX, OriginY, ScaleX, ScaleY;
        public float DpiScale = 1;
    }

    public sealed class CanvasConstraintLayout
    {
        public List<CanvasConstraintBadge> Badges { get; } = new List<CanvasConstraintBadge>();
        public List<CanvasConstraintLine> Lines { get; } = new List<CanvasConstraintLine>();

        public CanvasConstraintBadge BadgeAt(Point point)
        { return Badges.LastOrDefault(badge => badge.Bounds.Contains(point)); }
    }

    public sealed class CanvasConstraintBadge
    {
        public string RelationId, Kind, Label;
        public RectangleF Bounds;
        public Color Color;
        public bool Highlighted;
        public bool Limited;
    }

    public sealed class CanvasConstraintLine
    {
        public PointF Start, End;
        public Color Color;
    }

    public static class CanvasConstraintRenderer
    {
        private static readonly Color Satisfied = Color.FromArgb(36, 111, 205);
        private static readonly Color Disabled = Color.FromArgb(116, 122, 132);
        private static readonly Color Unsatisfied = Color.FromArgb(199, 57, 57);
        private static readonly Color Limiting = Color.FromArgb(168, 83, 10);

        public static CanvasConstraintLayout Build(ConstraintPaneModel model, CanvasConstraintDisplay display,
            CanvasConstraintProjection projection, Size viewport, string highlightedRelationId = null,
            IReadOnlyCollection<string> limitingRelationIds = null)
        {
            if (model == null) throw new ArgumentNullException("model");
            if (projection == null) throw new ArgumentNullException("projection");
            if (!Enum.IsDefined(typeof(CanvasConstraintDisplay), display)) throw new ArgumentOutOfRangeException("display");
            var layout = new CanvasConstraintLayout();
            if (display == CanvasConstraintDisplay.Hidden || viewport.Width <= 0 || viewport.Height <= 0) return layout;
            if (!Finite(projection.OriginX) || !Finite(projection.OriginY) || !Finite(projection.ScaleX) ||
                !Finite(projection.ScaleY) || projection.ScaleX <= 0 || projection.ScaleY <= 0 ||
                !Finite(projection.DpiScale) || projection.DpiScale <= 0)
                throw new ArgumentOutOfRangeException("projection", "A finite, positive canvas projection is required.");

            var objects = model.Objects.Where(item => item != null && !String.IsNullOrEmpty(item.Id))
                .ToDictionary(item => item.Id, StringComparer.Ordinal);
            var bounds = new Dictionary<string, RectangleF>(StringComparer.Ordinal);
            foreach (var item in objects.Values)
            {
                RectangleF projected;
                if (TryProject(item, projection, out projected)) bounds.Add(item.Id, projected);
            }
            var clip = new RectangleF(PointF.Empty, viewport);
            float dpi = projection.DpiScale, clearance = 8 * dpi;
            var otherBounds = objects.Values.Where(item => item.CanvasBounds.HasValue).Select(item => {
                var value = item.CanvasBounds.Value;
                return new RectangleF((float)(projection.OriginX+value.Left*projection.ScaleX),
                    (float)(projection.OriginY+value.Top*projection.ScaleY),(float)(value.Width*projection.ScaleX),
                    (float)(value.Height*projection.ScaleY));
            });
            var obstacles = bounds.Values.Concat(otherBounds).Select(rectangle => {
                rectangle.Inflate(clearance, clearance); return rectangle;
            }).Where(rectangle => rectangle.IntersectsWith(clip)).ToList();
            using (var measurementImage = new Bitmap(1, 1))
            using (var measurement = Graphics.FromImage(measurementImage))
            using (var font = CreateFont(dpi))
            using (var format = CreateFormat())
            {
                measurement.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                foreach (var relation in model.Relations)
                {
                    if (relation == null || relation.MemberIds == null) continue;
                    var members = relation.MemberIds.Where(id => id != null).Distinct(StringComparer.Ordinal).ToList();
                    bool limited = limitingRelationIds != null && limitingRelationIds.Contains(relation.Id);
                    if (display == CanvasConstraintDisplay.Selection && !limited && !members.Any(id => objects.ContainsKey(id) && objects[id].Selected)) continue;
                    var known = members.Where(bounds.ContainsKey).ToList();
                    var visible = known.Where(id => bounds[id].IntersectsWith(clip)).ToList();
                    if (visible.Count == 0) continue;
                    string anchor = visible.FirstOrDefault(id => objects[id].Selected) ?? visible[0];
                    string label = RelationLabel(relation);
                    float height = 24 * dpi, margin = 4 * dpi;
                    // Reserve the icon and all three padding gaps before measuring the label.
                    float width = 28 * dpi + (float)Math.Ceiling(measurement.MeasureString(label, font, PointF.Empty, format).Width);
                    width = Math.Min(Math.Min(width, 220 * dpi), viewport.Width - 2 * margin);
                    if (width < 24 * dpi || height > viewport.Height - 2 * margin) continue;
                    RectangleF badgeBounds;
                    if (!Place(bounds[anchor], new SizeF(width, height), viewport, dpi, obstacles, layout.Badges, out badgeBounds)) continue;
                    bool missing = known.Count != members.Count;
                    Color color = !relation.Enabled ? Disabled : missing ||
                        (!String.IsNullOrEmpty(relation.Status) && (relation.Status.Contains("未满足") ||
                            relation.Status.Contains("缺失") || relation.Status.Contains("不支持"))) ? Unsatisfied : Satisfied;
                    if(limited) color=Limiting;
                    bool highlighted = relation.Id == highlightedRelationId && highlightedRelationId != null;
                    layout.Badges.Add(new CanvasConstraintBadge {RelationId=relation.Id, Kind=relation.Kind,
                        Label=label, Bounds=badgeBounds, Color=color, Highlighted=highlighted, Limited=limited});
                    foreach (string member in known)
                    {
                        PointF from = EdgeToward(bounds[member], Center(badgeBounds));
                        PointF to = EdgeToward(badgeBounds, Center(bounds[member]));
                        if (ClipLine(ref from, ref to, viewport) && DistanceSquared(from, to) > 1)
                            layout.Lines.Add(new CanvasConstraintLine {Start=from, End=to, Color=Color.FromArgb(highlighted || limited ? 200 : 88, color)});
                    }
                }
            }
            return layout;
        }

        public static void Draw(Graphics graphics, CanvasConstraintLayout layout,
            IReadOnlyDictionary<string, Image> icons, float dpiScale)
        {
            if (graphics == null) throw new ArgumentNullException("graphics");
            if (layout == null) throw new ArgumentNullException("layout");
            if (!Finite(dpiScale) || dpiScale <= 0) throw new ArgumentOutOfRangeException("dpiScale");
            GraphicsState saved = graphics.Save();
            try
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.CompositingMode = CompositingMode.SourceOver;
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                foreach (var line in layout.Lines)
                    using (var pen = new Pen(line.Color, Math.Max(1, dpiScale))) graphics.DrawLine(pen, line.Start, line.End);
                using (var background = new SolidBrush(Color.FromArgb(240, 255, 255, 255)))
                using (var highlightedBackground = new SolidBrush(Color.FromArgb(250, 224, 238, 255)))
                using (var limitedBackground = new SolidBrush(Color.FromArgb(252, 255, 231, 198)))
                using (var font = CreateFont(dpiScale))
                using (var format = CreateFormat())
                foreach (var badge in layout.Badges)
                {
                    graphics.FillRectangle(badge.Limited ? limitedBackground : badge.Highlighted ? highlightedBackground : background, badge.Bounds);
                    bool emphasized=badge.Limited || badge.Highlighted;
                    using (var pen = new Pen(Color.FromArgb(emphasized ? 255 : 190, badge.Color), Math.Max(1, dpiScale)*(emphasized ? 2 : 1)))
                        graphics.DrawRectangle(pen, badge.Bounds.X, badge.Bounds.Y, badge.Bounds.Width, badge.Bounds.Height);
                    float padding = 4 * dpiScale;
                    int iconSize = Math.Max(1, (int)Math.Round(16 * dpiScale));
                    var iconBounds = new Rectangle((int)Math.Round(badge.Bounds.Left+padding),
                        (int)Math.Round(badge.Bounds.Top+(badge.Bounds.Height-iconSize)/2), iconSize, iconSize);
                    Image icon;
                    if (icons != null && badge.Kind != null && icons.TryGetValue(badge.Kind, out icon) && icon != null)
                    {
                        if (icon.Size == iconBounds.Size) graphics.DrawImageUnscaled(icon, iconBounds.Location);
                        else graphics.DrawImage(icon, iconBounds);
                    }
                    var text = RectangleF.FromLTRB(badge.Bounds.Left+24*dpiScale, badge.Bounds.Top,
                        badge.Bounds.Right-padding, badge.Bounds.Bottom);
                    if (text.Width > 0)
                        using (var brush = new SolidBrush(badge.Color)) graphics.DrawString(badge.Label ?? String.Empty, font, brush, text, format);
                }
            }
            finally { graphics.Restore(saved); }
        }

        private static Font CreateFont(float dpiScale)
        { return new Font(SystemFonts.MessageBoxFont.FontFamily, 12 * dpiScale, FontStyle.Regular, GraphicsUnit.Pixel); }

        private static StringFormat CreateFormat()
        { return new StringFormat {Alignment=StringAlignment.Near, LineAlignment=StringAlignment.Center,
            Trimming=StringTrimming.EllipsisCharacter, FormatFlags=StringFormatFlags.NoWrap}; }

        private static bool TryProject(ConstraintPaneObject item, CanvasConstraintProjection projection, out RectangleF bounds)
        {
            bounds = RectangleF.Empty;
            if (item.Geometry == null) return false;
            var geometry = item.Geometry;
            double x = projection.OriginX + geometry.Left * projection.ScaleX,
                y = projection.OriginY + geometry.Top * projection.ScaleY,
                width = geometry.Width * projection.ScaleX, height = geometry.Height * projection.ScaleY;
            if (new[] {x, y, width, height, x+width, y+height}.Any(value => !Finite(value) || Math.Abs(value) > Single.MaxValue) || width <= 0 || height <= 0) return false;
            bounds = new RectangleF((float)x, (float)y, (float)width, (float)height);
            return true;
        }

        private static string RelationLabel(ConstraintPaneRelation relation)
        {
            string label = relation.Label ?? relation.Kind ?? "约束";
            if (relation.Value.HasValue) label += " " + relation.Value.Value.ToString("G6", CultureInfo.InvariantCulture);
            if (relation.Second.HasValue) label += (relation.Value.HasValue ? ", " : " ") + relation.Second.Value.ToString("G6", CultureInfo.InvariantCulture);
            return label;
        }

        private static bool Place(RectangleF anchor, SizeF size, Size viewport, float dpi,
            IList<RectangleF> obstacles, IList<CanvasConstraintBadge> badges, out RectangleF result)
        {
            float margin=4*dpi, gap=8*dpi, rowStep=size.Height+4*dpi, columnStep=size.Width+4*dpi;
            float maximumX=viewport.Width-margin-size.Width, maximumY=viewport.Height-margin-size.Height;
            float firstY=Clamp(anchor.Top,margin,maximumY), firstX=Clamp(anchor.Left,margin,maximumX);
            foreach (float x in new[] {anchor.Right+gap, anchor.Left-gap-size.Width})
                foreach (float y in Positions(firstY, margin, maximumY, rowStep))
                    if (Accept(new RectangleF(x,y,size.Width,size.Height), viewport, margin, dpi, obstacles, badges, out result)) return true;
            foreach (float y in new[] {anchor.Top-gap-size.Height, anchor.Bottom+gap})
                foreach (float x in Positions(firstX, margin, maximumX, columnStep))
                    if (Accept(new RectangleF(x,y,size.Width,size.Height), viewport, margin, dpi, obstacles, badges, out result)) return true;
            result=RectangleF.Empty;return false;
        }

        private static IEnumerable<float> Positions(float first, float minimum, float maximum, float step)
        {
            for (float value=first;value<=maximum;value+=step) yield return value;
            for (float value=first-step;value>=minimum;value-=step) yield return value;
        }

        private static bool Accept(RectangleF candidate, Size viewport, float margin, float dpi,
            IList<RectangleF> obstacles, IList<CanvasConstraintBadge> badges, out RectangleF result)
        {
            result=RectangleF.Empty;
            if (candidate.Left<margin || candidate.Top<margin || candidate.Right>viewport.Width-margin || candidate.Bottom>viewport.Height-margin ||
                obstacles.Any(rectangle=>rectangle.IntersectsWith(candidate))) return false;
            var padded=candidate;padded.Inflate(2*dpi,2*dpi);
            if (badges.Any(badge=>badge.Bounds.IntersectsWith(padded))) return false;
            result=candidate;return true;
        }

        private static PointF Center(RectangleF bounds) { return new PointF(bounds.Left+bounds.Width/2,bounds.Top+bounds.Height/2); }
        private static PointF EdgeToward(RectangleF bounds, PointF target)
        {
            PointF center=Center(bounds);float dx=target.X-center.X,dy=target.Y-center.Y;
            if (dx==0 && dy==0) return center;
            float x=dx==0?Single.PositiveInfinity:bounds.Width/2/Math.Abs(dx),
                y=dy==0?Single.PositiveInfinity:bounds.Height/2/Math.Abs(dy),scale=Math.Min(x,y);
            return new PointF(center.X+dx*scale,center.Y+dy*scale);
        }

        // Liang-Barsky clipping keeps the line direction instead of clamping
        // its endpoints independently when a related member is offscreen.
        private static bool ClipLine(ref PointF start, ref PointF end, Size viewport)
        {
            double dx=(double)end.X-start.X,dy=(double)end.Y-start.Y,low=0,high=1;
            if (!Clip(-dx,start.X,ref low,ref high) || !Clip(dx,viewport.Width-start.X,ref low,ref high) ||
                !Clip(-dy,start.Y,ref low,ref high) || !Clip(dy,viewport.Height-start.Y,ref low,ref high)) return false;
            PointF original=start;
            start=new PointF(Clamp((float)(original.X+low*dx),0,viewport.Width),Clamp((float)(original.Y+low*dy),0,viewport.Height));
            end=new PointF(Clamp((float)(original.X+high*dx),0,viewport.Width),Clamp((float)(original.Y+high*dy),0,viewport.Height));
            return true;
        }
        private static bool Clip(double p,double q,ref double low,ref double high)
        {
            if(p==0)return q>=0;
            double ratio=q/p;
            if(p<0){if(ratio>high)return false;if(ratio>low)low=ratio;}
            else {if(ratio<low)return false;if(ratio<high)high=ratio;}
            return true;
        }
        private static double DistanceSquared(PointF first,PointF second)
        {double x=(double)first.X-second.X,y=(double)first.Y-second.Y;return x*x+y*y;}
        private static float Clamp(float value,float minimum,float maximum) {return Math.Max(minimum,Math.Min(maximum,value));}
        private static bool Finite(double value) {return !Double.IsNaN(value) && !Double.IsInfinity(value);}
    }
}
