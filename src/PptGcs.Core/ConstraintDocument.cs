using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace PptGcs.Core
{
    public static class ShapeIdentity
    {
        public static string Normalize(string value)
        {
            Guid parsed;
            if (!Guid.TryParse(value, out parsed) || parsed == Guid.Empty)
                throw new ArgumentException("A non-empty GUID is required.", "value");
            return parsed.ToString("D").ToUpperInvariant();
        }
    }

    public enum ConstraintKind
    {
        SameX,
        SameY,
        SameWidth,
        SameHeight,
        EqualHorizontalGap,
        EqualVerticalGap,
        RightAligned,
        BottomAligned,
        CenterXAligned,
        CenterYAligned,
        CentersCoincident,
        FixedWidth,
        FixedHeight,
        FixedPosition,
        AspectRatio,
        FixedHorizontalGap,
        FixedVerticalGap,
        Offset,
        SymmetricVertical,
        SymmetricHorizontal,
        PageLeft, PageCenterX, PageRight, PageTop, PageCenterY, PageBottom,
        ContainerLeft, ContainerCenterX, ContainerRight, ContainerTop, ContainerCenterY, ContainerBottom
    }

    public sealed class LayoutConstraint
    {
        public string Id { get; private set; }
        public int SlideId { get; private set; }
        public ConstraintKind Kind { get; private set; }
        public ReadOnlyCollection<string> ShapeIds { get; private set; }
        public double? Value { get; private set; }
        public double? SecondaryValue { get; private set; }
        public string ParameterId { get; private set; }
        public bool Enabled { get; private set; }
        internal double? PageLength { get; private set; }

        public LayoutConstraint(string id, int slideId, ConstraintKind kind, IEnumerable<string> shapeIds,
            double? value = null, double? secondaryValue = null, bool enabled = true, string parameterId = null)
        {
            Id = ShapeIdentity.Normalize(id);
            if (slideId <= 0) throw new ArgumentOutOfRangeException("slideId");
            if (!Enum.IsDefined(typeof(ConstraintKind), kind)) throw new ArgumentOutOfRangeException("kind");
            if (shapeIds == null) throw new ArgumentNullException("shapeIds");
            var ids = shapeIds.Select(ShapeIdentity.Normalize).ToList();
            int minimum = GetMinimumShapeCount(kind), maximum = GetMaximumShapeCount(kind);
            if (ids.Count < minimum || ids.Count > maximum)
                throw new ArgumentException(minimum == maximum ? "This constraint requires exactly " + minimum + " shapes." :
                    "This constraint requires at least " + minimum + " shapes.", "shapeIds");
            if (ids.Distinct(StringComparer.Ordinal).Count() != ids.Count)
                throw new ArgumentException("A constraint cannot reference a shape more than once.", "shapeIds");
            int parameters = GetParameterCount(kind);
            if (value.HasValue != (parameters >= 1) || secondaryValue.HasValue != (parameters == 2))
                throw new ArgumentException("This constraint requires " + parameters + " numeric parameters.", "value");
            if ((value.HasValue && !Finite(value.Value)) || (secondaryValue.HasValue && !Finite(secondaryValue.Value)))
                throw new ArgumentOutOfRangeException("value", "Constraint parameters must be finite.");
            if ((kind == ConstraintKind.FixedWidth || kind == ConstraintKind.FixedHeight) && value.Value < LayoutSolver.MinimumSize)
                throw new ArgumentOutOfRangeException("value", "A fixed dimension must be at least 1 point.");
            if (kind == ConstraintKind.AspectRatio && value.Value <= 0)
                throw new ArgumentOutOfRangeException("value", "An aspect ratio must be positive.");
            if (parameterId != null && !SupportsNamedParameter(kind))
                throw new ArgumentException("此约束类型不支持命名尺寸参数。", nameof(parameterId));
            SlideId = slideId;
            Kind = kind;
            ShapeIds = ids.AsReadOnly();
            Value = value; SecondaryValue = secondaryValue; Enabled = enabled;
            ParameterId = parameterId == null ? null : ShapeIdentity.Normalize(parameterId);
        }

        public static bool SupportsNamedParameter(ConstraintKind kind)
        {
            return kind == ConstraintKind.FixedWidth || kind == ConstraintKind.FixedHeight ||
                kind == ConstraintKind.FixedHorizontalGap || kind == ConstraintKind.FixedVerticalGap || RelativeAnchor.For(kind) != null;
        }

        public LayoutConstraint WithEnabled(bool enabled)
        { return new LayoutConstraint(Id, SlideId, Kind, ShapeIds, Value, SecondaryValue, enabled, ParameterId); }

        public static int GetMinimumShapeCount(ConstraintKind kind)
        {
            if (!Enum.IsDefined(typeof(ConstraintKind), kind)) throw new ArgumentOutOfRangeException("kind");
            var anchor=RelativeAnchor.For(kind);
            if(anchor!=null)return anchor.OnPage?1:2;
            if (kind == ConstraintKind.FixedWidth || kind == ConstraintKind.FixedHeight ||
                kind == ConstraintKind.FixedPosition || kind == ConstraintKind.AspectRatio) return 1;
            return kind == ConstraintKind.EqualHorizontalGap || kind == ConstraintKind.EqualVerticalGap ? 3 : 2;
        }

        public static int GetMaximumShapeCount(ConstraintKind kind)
        {
            int minimum = GetMinimumShapeCount(kind);
            if(RelativeAnchor.For(kind)!=null)return Int32.MaxValue;
            if (minimum == 1) return 1;
            return kind == ConstraintKind.Offset || kind == ConstraintKind.SymmetricVertical ||
                kind == ConstraintKind.SymmetricHorizontal ? 2 : Int32.MaxValue;
        }

        public static int GetParameterCount(ConstraintKind kind)
        {
            GetMinimumShapeCount(kind);
            if(RelativeAnchor.For(kind)!=null)return 1;
            if (kind == ConstraintKind.FixedPosition || kind == ConstraintKind.Offset) return 2;
            return kind == ConstraintKind.FixedWidth || kind == ConstraintKind.FixedHeight || kind == ConstraintKind.AspectRatio ||
                kind == ConstraintKind.FixedHorizontalGap || kind == ConstraintKind.FixedVerticalGap ||
                kind == ConstraintKind.SymmetricVertical || kind == ConstraintKind.SymmetricHorizontal ? 1 : 0;
        }

        public static bool PreservesOrder(ConstraintKind kind)
        {
            return kind == ConstraintKind.EqualHorizontalGap || kind == ConstraintKind.EqualVerticalGap ||
                kind == ConstraintKind.FixedHorizontalGap || kind == ConstraintKind.FixedVerticalGap || kind == ConstraintKind.Offset || RelativeAnchor.UsesContainer(kind);
        }

        public static string MembershipKey(ConstraintKind kind,IEnumerable<string> members)
        {
            if(RelativeAnchor.UsesContainer(kind))
                return String.Join(",",members.Take(1).Concat(members.Skip(1).OrderBy(id=>id,StringComparer.Ordinal)));
            return String.Join(",",PreservesOrder(kind)?members:members.OrderBy(id=>id,StringComparer.Ordinal));
        }

        // Page dimensions are read from the host for each operation and never serialized.
        public LayoutConstraint BindPageSize(double width,double height)
        {
            var anchor=RelativeAnchor.For(Kind);
            if(anchor==null || !anchor.OnPage)return this;
            double length=anchor.Horizontal?width:height;
            if(!Finite(length) || length<LayoutSolver.MinimumSize)throw new ArgumentOutOfRangeException("pageSize");
            return new LayoutConstraint(Id,SlideId,Kind,ShapeIds,Value,SecondaryValue,Enabled,ParameterId) {PageLength=length};
        }

        private static bool Finite(double value) { return !Double.IsNaN(value) && !Double.IsInfinity(value); }
    }

    public sealed class ConstraintDocument
    {
        public const string NamespaceUri = "urn:ppt-gcs:v2";
        private static readonly XNamespace Ns = NamespaceUri;
        public ReadOnlyCollection<LayoutConstraint> Constraints { get; private set; }
        public ReadOnlyCollection<LayoutParameter> Parameters { get; private set; }

        public ConstraintDocument(IEnumerable<LayoutConstraint> constraints, IEnumerable<LayoutParameter> parameters = null)
        {
            if (constraints == null) throw new ArgumentNullException("constraints");
            var library = (parameters ?? Enumerable.Empty<LayoutParameter>()).ToList();
            var parameterById = IndexParameters(library);
            var list = new List<LayoutConstraint>();
            var ids = new HashSet<string>(StringComparer.Ordinal);
            var relations = new HashSet<string>(StringComparer.Ordinal);
            foreach (var item in constraints)
            {
                if (item == null) throw new ArgumentException("A constraint cannot be null.", "constraints");
                if (!ids.Add(item.Id)) throw new ArgumentException("Duplicate constraint ID: " + item.Id);
                string key = item.SlideId.ToString(CultureInfo.InvariantCulture) + ":" + item.Kind + ":" + LayoutConstraint.MembershipKey(item.Kind,item.ShapeIds);
                if (!relations.Add(key)) throw new ArgumentException("Duplicate " + item.Kind + " relation on slide " + item.SlideId + ".");
                if (item.ParameterId == null) { list.Add(item); continue; }
                LayoutParameter parameter;
                if (!parameterById.TryGetValue(item.ParameterId, out parameter))
                    throw new ArgumentException("约束引用了不存在的命名参数。", nameof(constraints));
                list.Add(item.Value == parameter.Value ? item : new LayoutConstraint(item.Id, item.SlideId, item.Kind,
                    item.ShapeIds, parameter.Value, item.SecondaryValue, item.Enabled, parameter.Id));
            }
            Constraints = list.AsReadOnly();
            Parameters = library.AsReadOnly();
        }

        private static Dictionary<string, LayoutParameter> IndexParameters(IEnumerable<LayoutParameter> parameters)
        {
            var index = new Dictionary<string, LayoutParameter>(StringComparer.Ordinal);
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var parameter in parameters)
            {
                if (parameter == null) throw new ArgumentException("命名参数不能为空。", nameof(parameters));
                if (index.ContainsKey(parameter.Id)) throw new ArgumentException("命名参数编号重复。", nameof(parameters));
                if (!names.Add(parameter.Name)) throw new ArgumentException("已有同名参数“" + parameter.Name + "”。", nameof(parameters));
                index.Add(parameter.Id, parameter);
            }
            return index;
        }

        public ConstraintDocument WithConstraints(IEnumerable<LayoutConstraint> constraints)
        { return new ConstraintDocument(constraints, Parameters); }

        public string ToXml()
        {
            return new XElement(Ns + "constraints", Parameters.Select(item =>
                new XElement(Ns + "parameter", new XAttribute("id", item.Id), new XAttribute("name", item.Name),
                    new XAttribute("value", item.Value.ToString("R", CultureInfo.InvariantCulture)))), Constraints.Select(item =>
                new XElement(Ns + "constraint",
                    new XAttribute("id", item.Id),
                    new XAttribute("slideId", item.SlideId.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("kind", item.Kind.ToString()),
                    item.ParameterId != null ? new XAttribute("parameterId", item.ParameterId) : null,
                    item.ParameterId == null && item.Value.HasValue ? new XAttribute("value", item.Value.Value.ToString("R", CultureInfo.InvariantCulture)) : null,
                    item.SecondaryValue.HasValue ? new XAttribute("secondaryValue", item.SecondaryValue.Value.ToString("R", CultureInfo.InvariantCulture)) : null,
                    item.Enabled ? null : new XAttribute("enabled", "false"),
                    item.ShapeIds.Select(id => new XElement(Ns + "shape", new XAttribute("id", id))))))
                .ToString(SaveOptions.DisableFormatting);
        }

        public static ConstraintDocument Parse(string xml)
        {
            if (String.IsNullOrWhiteSpace(xml)) throw new FormatException("Constraint XML is empty.");
            try
            {
                var settings = new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    MaxCharactersInDocument = 1024 * 1024
                };
                XDocument document;
                using (var input = new StringReader(xml))
                using (var reader = XmlReader.Create(input, settings))
                    document = XDocument.Load(reader);
                var root = document.Root;
                if (root == null || root.Name != Ns + "constraints")
                    throw new FormatException("Unsupported constraint XML root or namespace.");
                CheckAttributes(root, new string[0]);
                CheckText(root);
                var parameters = root.Elements(Ns + "parameter").Select(ParseParameter).ToList();
                var parameterById = IndexParameters(parameters);
                var constraints = new List<LayoutConstraint>();
                foreach (var element in root.Elements().Where(item => item.Name != Ns + "parameter"))
                {
                    if (element.Name != Ns + "constraint") throw new FormatException("Unsupported constraint element: " + element.Name);
                    CheckAttributes(element, new[] { "id", "slideId", "kind", "value", "secondaryValue", "enabled", "parameterId" });
                    CheckText(element);
                    int slideId;
                    if (!Int32.TryParse((string)element.Attribute("slideId"), NumberStyles.None, CultureInfo.InvariantCulture, out slideId))
                        throw new FormatException("A positive numeric slideId is required.");
                    ConstraintKind kind;
                    string text = (string)element.Attribute("kind");
                    if (!Enum.TryParse(text, false, out kind) || !Enum.IsDefined(typeof(ConstraintKind), kind) || kind.ToString() != text)
                        throw new FormatException("Unsupported constraint kind.");
                    var shapeIds = new List<string>();
                    foreach (var shape in element.Elements())
                    {
                        if (shape.Name != Ns + "shape" || shape.HasElements) throw new FormatException("Unsupported shape reference.");
                        CheckAttributes(shape, new[] { "id" });
                        CheckText(shape);
                        shapeIds.Add((string)shape.Attribute("id"));
                    }
                    string enabled = (string)element.Attribute("enabled");
                    if (enabled != null && enabled != "true" && enabled != "false") throw new FormatException("Invalid enabled flag.");
                    double? value = ReadNumber(element, "value");
                    string parameterId = (string)element.Attribute("parameterId");
                    if (parameterId != null)
                    {
                        LayoutParameter parameter;
                        if (value.HasValue || !parameterById.TryGetValue(ShapeIdentity.Normalize(parameterId), out parameter))
                            throw new FormatException("A named binding must refer to an existing parameter and have no literal value.");
                        value = parameter.Value;
                    }
                    constraints.Add(new LayoutConstraint((string)element.Attribute("id"), slideId, kind, shapeIds,
                        value, ReadNumber(element, "secondaryValue"), enabled != "false", parameterId));
                }
                return new ConstraintDocument(constraints, parameters);
            }
            catch (XmlException exception) { throw new FormatException("Invalid constraint XML.", exception); }
            catch (ArgumentException exception) { throw new FormatException("Invalid constraint data: " + exception.Message, exception); }
        }

        private static LayoutParameter ParseParameter(XElement element)
        {
            CheckAttributes(element, new[] { "id", "name", "value" });
            CheckText(element);
            double? value = ReadNumber(element, "value");
            if (element.HasElements || !value.HasValue) throw new FormatException("A named parameter must have one numeric value.");
            return new LayoutParameter((string)element.Attribute("id"), (string)element.Attribute("name"), value.Value);
        }

        private static double? ReadNumber(XElement element, string name)
        {
            string value = (string)element.Attribute(name);
            if (value == null) return null;
            double number;
            if (!Double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number) ||
                Double.IsNaN(number) || Double.IsInfinity(number)) throw new FormatException("Invalid numeric parameter: " + name);
            return number;
        }

        private static void CheckAttributes(XElement element, string[] allowed)
        {
            foreach (var attribute in element.Attributes())
            {
                if (attribute.IsNamespaceDeclaration) continue;
                if (attribute.Name.Namespace != XNamespace.None || !allowed.Contains(attribute.Name.LocalName))
                    throw new FormatException("Unsupported attribute: " + attribute.Name);
            }
        }

        private static void CheckText(XElement element)
        {
            if (element.Nodes().OfType<XText>().Any(text => !String.IsNullOrWhiteSpace(text.Value)))
                throw new FormatException("Constraint elements cannot contain text.");
        }
    }
}
