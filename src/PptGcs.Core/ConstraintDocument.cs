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
        EqualVerticalGap
    }

    public sealed class LayoutConstraint
    {
        public string Id { get; private set; }
        public int SlideId { get; private set; }
        public ConstraintKind Kind { get; private set; }
        public ReadOnlyCollection<string> ShapeIds { get; private set; }

        public LayoutConstraint(string id, int slideId, ConstraintKind kind, IEnumerable<string> shapeIds)
        {
            Id = ShapeIdentity.Normalize(id);
            if (slideId <= 0) throw new ArgumentOutOfRangeException("slideId");
            if (!Enum.IsDefined(typeof(ConstraintKind), kind)) throw new ArgumentOutOfRangeException("kind");
            if (shapeIds == null) throw new ArgumentNullException("shapeIds");
            var ids = shapeIds.Select(ShapeIdentity.Normalize).ToList();
            int minimum = kind == ConstraintKind.EqualHorizontalGap || kind == ConstraintKind.EqualVerticalGap ? 3 : 2;
            if (ids.Count < minimum) throw new ArgumentException("This constraint requires at least " + minimum + " shapes.", "shapeIds");
            if (ids.Distinct(StringComparer.Ordinal).Count() != ids.Count)
                throw new ArgumentException("A constraint cannot reference a shape more than once.", "shapeIds");
            SlideId = slideId;
            Kind = kind;
            ShapeIds = ids.AsReadOnly();
        }
    }

    public sealed class ConstraintDocument
    {
        public const string NamespaceUri = "urn:ppt-gcs:v2";
        private static readonly XNamespace Ns = NamespaceUri;
        public ReadOnlyCollection<LayoutConstraint> Constraints { get; private set; }

        public ConstraintDocument(IEnumerable<LayoutConstraint> constraints)
        {
            if (constraints == null) throw new ArgumentNullException("constraints");
            var list = constraints.ToList();
            var ids = new HashSet<string>(StringComparer.Ordinal);
            var relations = new HashSet<string>(StringComparer.Ordinal);
            foreach (var item in list)
            {
                if (item == null) throw new ArgumentException("A constraint cannot be null.", "constraints");
                if (!ids.Add(item.Id)) throw new ArgumentException("Duplicate constraint ID: " + item.Id);
                bool ordered = item.Kind == ConstraintKind.EqualHorizontalGap || item.Kind == ConstraintKind.EqualVerticalGap;
                IEnumerable<string> members = ordered ? item.ShapeIds.AsEnumerable() : item.ShapeIds.OrderBy(id => id, StringComparer.Ordinal);
                string key = item.SlideId.ToString(CultureInfo.InvariantCulture) + ":" + item.Kind + ":" + String.Join(",", members);
                if (!relations.Add(key)) throw new ArgumentException("Duplicate " + item.Kind + " relation on slide " + item.SlideId + ".");
            }
            Constraints = list.AsReadOnly();
        }

        public string ToXml()
        {
            return new XElement(Ns + "constraints", Constraints.Select(item =>
                new XElement(Ns + "constraint",
                    new XAttribute("id", item.Id),
                    new XAttribute("slideId", item.SlideId.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("kind", item.Kind.ToString()),
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
                var constraints = new List<LayoutConstraint>();
                foreach (var element in root.Elements())
                {
                    if (element.Name != Ns + "constraint") throw new FormatException("Unsupported constraint element: " + element.Name);
                    CheckAttributes(element, new[] { "id", "slideId", "kind" });
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
                    constraints.Add(new LayoutConstraint((string)element.Attribute("id"), slideId, kind, shapeIds));
                }
                return new ConstraintDocument(constraints);
            }
            catch (XmlException exception) { throw new FormatException("Invalid constraint XML.", exception); }
            catch (ArgumentException exception) { throw new FormatException("Invalid constraint data: " + exception.Message, exception); }
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
