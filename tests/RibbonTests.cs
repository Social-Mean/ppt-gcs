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
        var referenceMenus=tab.Descendants(ns+"menu").ToList();
        Check(referenceMenus.Select(menu=>(string)menu.Attribute("id")).SequenceEqual(new[]{"PageConstraints","ContainerConstraints"}) &&
            referenceMenus.All(menu=>menu.Elements(ns+"button").Count()==6) &&
            constraints.All(button=>{
                var anchor=RelativeAnchor.For(RibbonDefinition.ParseKind((string)button.Attribute("tag")));
                return anchor==null?button.Parent.Name==ns+"group":
                    (string)button.Parent.Attribute("id")== (anchor.OnPage?"PageConstraints":"ContainerConstraints");
            }), "reference constraints have their own two menus and original commands remain direct buttons");
        Check(root.DescendantsAndSelf().Attributes("id").Select(attribute => attribute.Value).Distinct().Count() == root.DescendantsAndSelf().Attributes("id").Count(),
            "all custom ribbon control identifiers are unique");
        Check(constraints.All(button => (string)button.Attribute("onAction") == "CreateConstraint" &&
            RibbonDefinition.ParseKind((string)button.Attribute("tag")).ToString() == (string)button.Attribute("tag")),
            "all constraint buttons dispatch through one canonical kind tag");
        Check(constraints.All(button => (string)button.Attribute("label") == options.Single(option => option.Kind == (string)button.Attribute("tag")).Label),
            "ribbon and pane share every constraint label");
        Check(constraints.All(button => (string)button.Attribute("imageMso")==OfficeIconProvider.ImageMso((string)button.Attribute("tag")) &&
            !String.IsNullOrWhiteSpace((string)button.Attribute("imageMso")) && button.Attribute("getImage")==null),
            "all constraint buttons use the shared native Office icon resources");
        string[] nativeLabels = {"左对齐", "水平居中", "右对齐", "顶端对齐", "垂直居中", "底端对齐"};
        Check(options.Take(6).Select(option => option.Label).SequenceEqual(nativeLabels),
            "alignment labels and ordering match the native PowerPoint menu");
        var parameterKinds = new HashSet<ConstraintKind> {ConstraintKind.FixedWidth, ConstraintKind.FixedHeight,
            ConstraintKind.FixedPosition, ConstraintKind.AspectRatio, ConstraintKind.FixedHorizontalGap,
            ConstraintKind.FixedVerticalGap, ConstraintKind.Offset, ConstraintKind.SymmetricVertical, ConstraintKind.SymmetricHorizontal};
        parameterKinds.UnionWith(kinds.Where(kind=>RelativeAnchor.For(kind)!=null));
        Check(kinds.All(kind => RibbonDefinition.RequiresParameterEditor(kind) == parameterKinds.Contains(kind)),
            "all numeric constraints, including page and container anchors, require the parameter editor");
        Check(constraints.Where(button => parameterKinds.Contains(RibbonDefinition.ParseKind((string)button.Attribute("tag"))))
            .All(button => ((string)button.Attribute("supertip")).Contains("设置参数并确认创建")),
            "numeric constraint tooltips explain explicit parameter confirmation");
        var groups = tab.Elements(ns + "group").ToList();
        Check(groups.Select(group => (string)group.Attribute("label")).SequenceEqual(new[] {"位置约束","尺寸约束","间距约束","偏移与对称","参考框约束","管理","组件与参数","画布标记"}) &&
            groups.Select(group => group.Elements(ns + "button").Count()).SequenceEqual(new[] {8,5,4,3,0,6,2,0}),
            "constraint creation, management and canvas display controls are grouped separately");
        Check(buttons.Any(button => (string)button.Attribute("id") == "LayersPane" &&
            (string)button.Attribute("label") == "图层窗格" && (string)button.Attribute("onAction") == "ShowLayerPane"),
            "the layer tab can be opened from the Ribbon");
        var displayButtons=tab.Descendants(ns+"toggleButton").ToList();
        Check(displayButtons.Select(button=>(string)button.Attribute("tag")).SequenceEqual(new[]{"Selection","All","Hidden"}) &&
            displayButtons.All(button=>(string)button.Attribute("getPressed")=="GetCanvasConstraintDisplay" && (string)button.Attribute("onAction")=="ChangeCanvasConstraintDisplay"),
            "canvas display exposes mutually exclusive selection, all and hidden modes");
        Check((string)root.Attribute("onLoad")=="OnRibbonLoad","Office supplies the ribbon instance used to refresh display toggle state");
        var paneButton = buttons.Single(button => (string)button.Attribute("id") == "ConstraintsPane");
        Check((string)paneButton.Attribute("label") == "约束窗格" && (string)paneButton.Attribute("onAction") == "ShowConstraintPane",
            "ribbon opens the renamed constraint pane");
        Check(buttons.Where(button => button.Attribute("tag") == null).Select(button => (string)button.Attribute("onAction"))
            .SequenceEqual(new[] {"ShowConstraintPane","ShowLayerPane","ShowPageResize","UnlinkSelected","CreateDemo","ShowStatus","ShowComponentCopy","ShowNamedParameters"}),
            "management includes the page size, copy and named parameter commands");
        var arrange = root.Element(ns + "commands").Elements(ns + "command").ToArray();
        Check(arrange.Select(command => (string)command.Attribute("idMso")).SequenceEqual(new[] {
            "ObjectBringToFront", "ObjectBringForward", "ObjectSendBackward", "ObjectSendToBack"
        }) && arrange.All(command => (string)command.Attribute("onAction") == "OnLayerArrange"),
            "native arrange commands share the layer-aware route");
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
