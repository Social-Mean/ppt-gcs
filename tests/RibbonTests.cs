using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using PptGcs.AddIn;
using PptGcs.Core;

internal static class RibbonTests
{
    private static int passed;
    private static void Main()
    {
        XNamespace ns = "http://schemas.microsoft.com/office/2009/07/customui";
        XElement root = XElement.Parse(RibbonDefinition.CreateXml());
        var tab = root.Element(ns + "ribbon").Element(ns + "tabs").Elements(ns + "tab").Single();
        var buttons = tab.Descendants(ns + "button").ToList();
        var constraints = buttons.Where(button => button.Attribute("tag") != null).ToList();
        var options = ConstraintCatalog.Options(new List<ShapeGeometry>());
        var kinds = Enum.GetValues(typeof(ConstraintKind)).Cast<ConstraintKind>().ToArray();

        Check(root.Name == ns + "customUI" && (string)tab.Attribute("label") == "PPT GCS",
            "ribbon is valid namespaced Office customUI with one product tab");
        Check(constraints.Count == kinds.Length && constraints.Select(button => (string)button.Attribute("tag")).Distinct().Count() == kinds.Length &&
            new HashSet<string>(constraints.Select(button => (string)button.Attribute("tag"))).SetEquals(kinds.Select(kind => kind.ToString())),
            "every constraint kind appears exactly once in the ribbon");
        Check(constraints.All(button => button.Parent.Name == ns + "group") && !tab.Descendants(ns + "menu").Any(),
            "all twenty constraints are directly accessible group buttons");
        Check(root.DescendantsAndSelf().Attributes("id").Select(attribute => attribute.Value).Distinct().Count() == root.DescendantsAndSelf().Attributes("id").Count(),
            "all custom ribbon control identifiers are unique");
        Check(constraints.All(button => (string)button.Attribute("onAction") == "CreateConstraint" &&
            RibbonDefinition.ParseKind((string)button.Attribute("tag")).ToString() == (string)button.Attribute("tag")),
            "all constraint buttons dispatch through one canonical kind tag");
        Check(constraints.All(button => (string)button.Attribute("label") == options.Single(option => option.Kind == (string)button.Attribute("tag")).Label),
            "ribbon and pane share every constraint label");
        Check(constraints.All(button => OfficeIconProvider.UsesFramedImage((string)button.Attribute("tag"))
            ? (string)button.Attribute("getImage")=="GetCentersCoincidentImage" && button.Attribute("imageMso")==null
            : (string)button.Attribute("imageMso")==OfficeIconProvider.ImageMso((string)button.Attribute("tag")) && !String.IsNullOrWhiteSpace((string)button.Attribute("imageMso"))),
            "all icon routes use shared resources and the framed center icon uses its Ribbon image callback");
        string[] nativeLabels = {"左对齐", "水平居中", "右对齐", "顶端对齐", "垂直居中", "底端对齐"};
        Check(options.Take(6).Select(option => option.Label).SequenceEqual(nativeLabels),
            "alignment labels and ordering match the native PowerPoint menu");
        var parameterKinds = new HashSet<ConstraintKind> {ConstraintKind.FixedWidth, ConstraintKind.FixedHeight,
            ConstraintKind.FixedPosition, ConstraintKind.AspectRatio, ConstraintKind.FixedHorizontalGap,
            ConstraintKind.FixedVerticalGap, ConstraintKind.Offset, ConstraintKind.SymmetricVertical, ConstraintKind.SymmetricHorizontal};
        Check(kinds.All(kind => RibbonDefinition.RequiresParameterEditor(kind) == parameterKinds.Contains(kind)),
            "all nine numeric constraints require the parameter editor and eleven do not");
        Check(constraints.Where(button => parameterKinds.Contains(RibbonDefinition.ParseKind((string)button.Attribute("tag"))))
            .All(button => ((string)button.Attribute("supertip")).Contains("设置参数并确认创建")),
            "numeric constraint tooltips explain explicit parameter confirmation");
        var groups = tab.Elements(ns + "group").ToList();
        Check(groups.Select(group => (string)group.Attribute("label")).SequenceEqual(new[] {"位置约束","尺寸约束","间距约束","偏移与对称","管理"}) &&
            groups.Select(group => group.Elements(ns + "button").Count()).SequenceEqual(new[] {8,5,4,3,4}),
            "constraints are grouped as position, dimensions, spacing and relations with management separate");
        var paneButton = buttons.Single(button => (string)button.Attribute("id") == "ConstraintsPane");
        Check((string)paneButton.Attribute("label") == "约束窗格" && (string)paneButton.Attribute("onAction") == "ShowConstraintPane",
            "ribbon opens the renamed constraint pane");
        Check(buttons.Where(button => button.Attribute("tag") == null).Select(button => (string)button.Attribute("onAction"))
            .SequenceEqual(new[] {"ShowConstraintPane","UnlinkSelected","CreateDemo","ShowStatus"}),
            "existing management actions remain explicit and complete");
        Check(new[] {null,"","0","samex","SameX ","unknown"}.All(Rejected),
            "malformed or numeric callback tags cannot select an unintended constraint");
        bool invalidKind = false;
        try { RibbonDefinition.RequiresParameterEditor((ConstraintKind)999); }
        catch (ArgumentOutOfRangeException) { invalidKind = true; }
        Check(invalidKind, "undefined kinds cannot bypass parameter routing");
        Console.WriteLine("PASS: " + passed + " ribbon tests");
    }

    private static bool Rejected(string value)
    {
        try { RibbonDefinition.ParseKind(value); return false; }
        catch (ArgumentException) { return true; }
    }
    private static void Check(bool condition, string description)
    {
        if (!condition) throw new Exception(description);
        passed++;
        Console.WriteLine("PASS " + description);
    }
}
