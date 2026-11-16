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
        // PowerPoint Tags uppercases stored values. Normalize before comparisons.
        public static string Normalize(string value)
        {
            Guid parsed;
            if (!Guid.TryParse(value, out parsed) || parsed == Guid.Empty)
                throw new ArgumentException("A non-empty GUID is required.", "value");
            return parsed.ToString("D").ToUpperInvariant();
        }
    }

    public sealed class SameTopConstraint
    {
        public string Id { get; private set; }
        public int SlideId { get; private set; }
        public string ShapeA { get; private set; }
        public string ShapeB { get; private set; }

        public SameTopConstraint(string id, int slideId, string shapeA, string shapeB)
        {
            Id = ShapeIdentity.Normalize(id);
            if (slideId <= 0) throw new ArgumentOutOfRangeException("slideId");
            SlideId = slideId;
            ShapeA = ShapeIdentity.Normalize(shapeA);
            ShapeB = ShapeIdentity.Normalize(shapeB);
            if (ShapeA == ShapeB)
                throw new ArgumentException("A same-top constraint requires two different shapes.");
        }
    }

    public sealed class ConstraintDocument
    {
        public const string NamespaceUri = "urn:ppt-gcs:v1";
        private static readonly XNamespace Ns = NamespaceUri;
        public ReadOnlyCollection<SameTopConstraint> Constraints { get; private set; }

        public ConstraintDocument(IEnumerable<SameTopConstraint> constraints)
        {
            if (constraints == null) throw new ArgumentNullException("constraints");
            var list = constraints.ToList();
            var ids = new HashSet<string>(StringComparer.Ordinal);
            var pairs = new HashSet<string>(StringComparer.Ordinal);
            foreach (var item in list)
            {
                if (item == null) throw new ArgumentException("A constraint cannot be null.", "constraints");
                if (!ids.Add(item.Id)) throw new ArgumentException("Duplicate constraint ID: " + item.Id);
                string first = String.CompareOrdinal(item.ShapeA, item.ShapeB) < 0 ? item.ShapeA : item.ShapeB;
                string second = first == item.ShapeA ? item.ShapeB : item.ShapeA;
                string key = item.SlideId.ToString(CultureInfo.InvariantCulture) + ":" + first + ":" + second;
                if (!pairs.Add(key)) throw new ArgumentException("Duplicate same-top relation on slide " + item.SlideId + ".");
            }
            Constraints = list.AsReadOnly();
        }

        public string ToXml()
        {
            return new XElement(Ns + "constraints", Constraints.Select(item =>
                new XElement(Ns + "sameTop",
                    new XAttribute("id", item.Id),
                    new XAttribute("slideId", item.SlideId.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("shapeA", item.ShapeA),
                    new XAttribute("shapeB", item.ShapeB)))).ToString(SaveOptions.DisableFormatting);
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
                var constraints = new List<SameTopConstraint>();
                foreach (var element in root.Elements())
                {
                    if (element.Name != Ns + "sameTop" || element.HasElements)
                        throw new FormatException("Unsupported constraint element: " + element.Name);
                    CheckAttributes(element, new[] { "id", "slideId", "shapeA", "shapeB" });
                    CheckText(element);
                    int slideId;
                    if (!Int32.TryParse((string)element.Attribute("slideId"), NumberStyles.None,
                        CultureInfo.InvariantCulture, out slideId))
                        throw new FormatException("A positive numeric slideId is required.");
                    constraints.Add(new SameTopConstraint((string)element.Attribute("id"), slideId,
                        (string)element.Attribute("shapeA"), (string)element.Attribute("shapeB")));
                }
                return new ConstraintDocument(constraints);
            }
            catch (XmlException exception)
            {
                throw new FormatException("Invalid constraint XML.", exception);
            }
            catch (ArgumentException exception)
            {
                throw new FormatException("Invalid constraint data: " + exception.Message, exception);
            }
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
