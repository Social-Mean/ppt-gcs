using System;
using System.Collections.Generic;
using System.Linq;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    public static class ConstraintCatalog
    {
        public static IList<ConstraintPaneOption> Options(IList<ShapeGeometry> selected)
        {
            ShapeGeometry a = selected.FirstOrDefault(), b = selected.Skip(1).FirstOrDefault();
            var horizontal = selected.OrderBy(s => s.Left).ToList();
            var vertical = selected.OrderBy(s => s.Top).ToList();
            return new List<ConstraintPaneOption> {
                Choice(ConstraintKind.SameX,"左对齐"), Choice(ConstraintKind.CenterXAligned,"水平居中"),
                Choice(ConstraintKind.RightAligned,"右对齐"), Choice(ConstraintKind.SameY,"顶端对齐"),
                Choice(ConstraintKind.CenterYAligned,"垂直居中"), Choice(ConstraintKind.BottomAligned,"底端对齐"),
                Choice(ConstraintKind.CentersCoincident,"中心重合"),
                Choice(ConstraintKind.SameWidth,"同宽"), Choice(ConstraintKind.SameHeight,"同高"),
                Choice(ConstraintKind.EqualHorizontalGap,"水平等间距",3), Choice(ConstraintKind.EqualVerticalGap,"垂直等间距",3),
                Choice(ConstraintKind.FixedWidth,"固定宽度",1,1,"宽度 (pt)",a == null ? 120 : a.Width,null,null,1),
                Choice(ConstraintKind.FixedHeight,"固定高度",1,1,"高度 (pt)",a == null ? 80 : a.Height,null,null,1),
                Choice(ConstraintKind.FixedPosition,"固定位置",1,1,"左侧 X (pt)",a == null ? 0 : a.Left,"顶部 Y (pt)",a == null ? 0 : a.Top),
                Choice(ConstraintKind.AspectRatio,"固定宽高比",1,1,"宽 / 高",a == null ? 1 : a.Width/a.Height,null,null,0.001),
                Choice(ConstraintKind.FixedHorizontalGap,"固定水平间距",2,0,"边缘间距 (pt)",horizontal.Count < 2 ? 20 : horizontal[1].Left-horizontal[0].Left-horizontal[0].Width),
                Choice(ConstraintKind.FixedVerticalGap,"固定垂直间距",2,0,"边缘间距 (pt)",vertical.Count < 2 ? 20 : vertical[1].Top-vertical[0].Top-vertical[0].Height),
                Choice(ConstraintKind.Offset,"相对偏移",2,2,"第二个相对第一个 X (pt)",a == null || b == null ? 0 : b.Left-a.Left,"相对 Y (pt)",a == null || b == null ? 0 : b.Top-a.Top),
                Choice(ConstraintKind.SymmetricVertical,"关于垂直轴对称",2,2,"对称轴 X (pt)",a == null || b == null ? 480 : (a.Left+a.Width/2+b.Left+b.Width/2)/2),
                Choice(ConstraintKind.SymmetricHorizontal,"关于水平轴对称",2,2,"对称轴 Y (pt)",a == null || b == null ? 270 : (a.Top+a.Height/2+b.Top+b.Height/2)/2)
            };
        }

        private static ConstraintPaneOption Choice(ConstraintKind kind, string label, int minimum=2, int maximum=0,
            string valueLabel=null, double? value=null, string secondLabel=null, double? second=null, double minValue=-1000000)
        {
            return new ConstraintPaneOption { Kind=kind.ToString(),Label=label,MinimumMembers=minimum,MaximumMembers=maximum,
                ValueLabel=valueLabel,SecondLabel=secondLabel,DefaultValue=value,DefaultSecond=second,MinimumValue=minValue,
                Description=kind == ConstraintKind.EqualHorizontalGap || kind == ConstraintKind.EqualVerticalGap
                    ? "按建立时的顺序保持边缘间距相等，间距值可以变化。"
                    : kind == ConstraintKind.FixedHorizontalGap || kind == ConstraintKind.FixedVerticalGap
                    ? "按建立时的顺序保持指定边缘间距；负值表示重叠。"
                    : kind == ConstraintKind.Offset ? "成员顺序按选择范围排列，偏移为第二个对象减第一个对象。"
                    : kind == ConstraintKind.SymmetricVertical || kind == ConstraintKind.SymmetricHorizontal
                    ? "两个对象的外接框关于固定轴对称，尺寸相等；内容方向不翻转。" : "单位为磅 (pt)，宽高比为无单位比值。" };
        }

        public static string Label(ConstraintKind kind)
        { return Options(new List<ShapeGeometry>()).Single(o => o.Kind == kind.ToString()).Label; }

        public static bool IsHorizontalGap(ConstraintKind kind)
        { return kind == ConstraintKind.EqualHorizontalGap || kind == ConstraintKind.FixedHorizontalGap; }
        public static bool IsVerticalGap(ConstraintKind kind)
        { return kind == ConstraintKind.EqualVerticalGap || kind == ConstraintKind.FixedVerticalGap; }
    }
}
