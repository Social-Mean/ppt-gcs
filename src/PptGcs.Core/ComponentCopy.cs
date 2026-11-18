using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PptGcs.Core
{
    public static class ComponentCopy
    {
        public static ConstraintDocument Plan(ConstraintDocument document, int sourceSlideId, int targetSlideId,
            IReadOnlyDictionary<string, string> shapeIds, bool shareParameters)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (shapeIds == null) throw new ArgumentNullException(nameof(shapeIds));
            if (sourceSlideId <= 0 || targetSlideId <= 0) throw new ArgumentOutOfRangeException(nameof(targetSlideId));
            var mapping = shapeIds.ToDictionary(pair => ShapeIdentity.Normalize(pair.Key),
                pair => ShapeIdentity.Normalize(pair.Value), StringComparer.Ordinal);
            var existing = new HashSet<string>(document.Constraints.SelectMany(item => item.ShapeIds).Concat(mapping.Keys), StringComparer.Ordinal);
            if (mapping.Values.Any(id => !existing.Add(id)))
                throw new ArgumentException("副本必须使用不同于原件和现有对象的唯一编号。", nameof(shapeIds));

            var internalRelations = document.Constraints.Where(item => item.SlideId == sourceSlideId &&
                item.ShapeIds.All(mapping.ContainsKey)).ToList();
            var library = document.Parameters.ToList();
            var replacements = new Dictionary<string, string>(StringComparer.Ordinal);
            if (!shareParameters)
            {
                var used = new HashSet<string>(internalRelations.Select(item => item.ParameterId).Where(id => id != null), StringComparer.Ordinal);
                var names = new HashSet<string>(library.Select(parameter => parameter.Name), StringComparer.OrdinalIgnoreCase);
                foreach (var parameter in document.Parameters.Where(item => used.Contains(item.Id)))
                {
                    var copy = new LayoutParameter(Guid.NewGuid().ToString("D"), NextName(parameter.Name, names), parameter.Value);
                    library.Add(copy);
                    replacements.Add(parameter.Id, copy.Id);
                }
            }

            var copied = internalRelations.Select(item => {
                string parameterId = item.ParameterId;
                if (!shareParameters && parameterId != null) parameterId = replacements[parameterId];
                return new LayoutConstraint(Guid.NewGuid().ToString("D"), targetSlideId, item.Kind,
                    item.ShapeIds.Select(id => mapping[id]), item.Value, item.SecondaryValue, item.Enabled, parameterId);
            });
            return new ConstraintDocument(document.Constraints.Concat(copied), library);
        }

        private static string NextName(string source, HashSet<string> names)
        {
            for (int i = 1; ; i++)
            {
                string suffix = i == 1 ? "（副本）" : "（副本 " + i.ToString(CultureInfo.InvariantCulture) + "）";
                int length = Math.Min(source.Length, LayoutParameter.MaximumNameLength - suffix.Length);
                if (length > 0 && Char.IsHighSurrogate(source[length - 1])) length--;
                string name = source.Substring(0, length) + suffix;
                if (names.Add(name)) return name;
            }
        }
    }
}
