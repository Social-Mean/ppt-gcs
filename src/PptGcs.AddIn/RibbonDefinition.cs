using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    // Every entry uses the same catalog, icon map and command route as the pane.
    public static class RibbonDefinition
    {
        private static readonly XNamespace Ns = "http://schemas.microsoft.com/office/2009/07/customui";

        public static string CreateXml()
        {
            var options = ConstraintCatalog.Options(new List<ShapeGeometry>()).ToDictionary(option => ParseKind(option.Kind));
            var tab = new XElement(Ns + "tab", new XAttribute("id", "PptGcsTab"), new XAttribute("label", "PPT GCS"),
                Group(options, "Alignment", "位置约束", ConstraintKind.SameX, ConstraintKind.CenterXAligned,
                    ConstraintKind.RightAligned, ConstraintKind.SameY, ConstraintKind.CenterYAligned,
                    ConstraintKind.BottomAligned, ConstraintKind.CentersCoincident, ConstraintKind.FixedPosition),
                Group(options, "Dimensions", "尺寸约束", ConstraintKind.SameWidth, ConstraintKind.SameHeight,
                    ConstraintKind.FixedWidth, ConstraintKind.FixedHeight, ConstraintKind.AspectRatio),
                Group(options, "Spacing", "间距约束", ConstraintKind.EqualHorizontalGap, ConstraintKind.EqualVerticalGap,
                    ConstraintKind.FixedHorizontalGap, ConstraintKind.FixedVerticalGap),
                Group(options, "Relations", "偏移与对称", ConstraintKind.Offset,
                    ConstraintKind.SymmetricVertical, ConstraintKind.SymmetricHorizontal),
                new XElement(Ns + "group", new XAttribute("id", "Manage"), new XAttribute("label", "管理"),
                    ManageButton("ConstraintsPane", "约束窗格", "view", "ShowConstraintPane"),
                    ManageButton("Remove", "移除选中对象的约束", "remove", "UnlinkSelected"),
                    ManageButton("Demo", "新建六组演示", "demo", "CreateDemo"),
                    ManageButton("Status", "状态", "status", "ShowStatus")));
            return new XElement(Ns + "customUI", new XElement(Ns + "ribbon", new XElement(Ns + "tabs", tab)))
                .ToString(SaveOptions.DisableFormatting);
        }

        public static ConstraintKind ParseKind(string tag)
        {
            ConstraintKind kind;
            if (!Enum.TryParse(tag, false, out kind) || !Enum.IsDefined(typeof(ConstraintKind), kind) || kind.ToString() != tag)
                throw new ArgumentException("未知约束类型。", "tag");
            return kind;
        }

        public static bool RequiresParameterEditor(ConstraintKind kind)
        { return LayoutConstraint.GetParameterCount(kind) != 0; }

        private static XElement Group(IDictionary<ConstraintKind, ConstraintPaneOption> options,
            string id, string label, params ConstraintKind[] kinds)
        {
            return new XElement(Ns + "group", new XAttribute("id", id), new XAttribute("label", label),
                kinds.Select(kind => ConstraintButton(kind, options[kind])));
        }

        private static XElement ConstraintButton(ConstraintKind kind, ConstraintPaneOption option)
        {
            string members = option.MinimumMembers == option.MaximumMembers
                ? "选中 " + option.MinimumMembers + " 个对象。" : "至少选中 " + option.MinimumMembers + " 个对象。";
            string action = RequiresParameterEditor(kind) ? "在约束窗格中设置参数并确认创建。"
                : "创建后拖动任一成员会保持该关系。";
            return new XElement(Ns + "button", new XAttribute("id", "Constraint" + option.Kind),
                new XAttribute("tag", option.Kind), new XAttribute("label", option.Label),
                new XAttribute("imageMso", OfficeIconProvider.ImageMso(option.Kind)),
                new XAttribute("onAction", "CreateConstraint"), new XAttribute("screentip", option.Label + "约束"),
                new XAttribute("supertip", members + action + Environment.NewLine + option.Description));
        }

        private static XElement ManageButton(string id, string label, string icon, string action)
        {
            return new XElement(Ns + "button", new XAttribute("id", id), new XAttribute("label", label),
                new XAttribute("imageMso", OfficeIconProvider.ImageMso(icon)), new XAttribute("onAction", action));
        }
    }
}
