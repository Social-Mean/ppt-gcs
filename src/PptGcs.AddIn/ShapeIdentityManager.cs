using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    // A copied shape carries its source's tags, but its native owner changes.
    // Preserve the source GUID and give each copy a fresh identity. Constraint
    // XML and shape geometry are deliberately outside this component's writes.
    public sealed class ShapeIdentityManager
    {
        public const string OwnerTagName = "PPTGCS_OWNER";
        public const string DocumentTagName = "PPTGCS_DOCUMENT";

        // Returns the number of shapes assigned a new GUID. Merely recording
        // the ownership of an existing unique GUID does not count as a copy.
        public int Normalize(object presentationObject)
        {
            if (presentationObject == null) throw new ArgumentNullException("presentationObject");
            dynamic presentation = presentationObject;
            var shapes = new List<IdentityShape>();
            var physicalKeys = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 1; i <= (int)presentation.Slides.Count; i++)
            {
                dynamic slide = presentation.Slides[i];
                int slideId = (int)slide.SlideID;
                for (int j = 1; j <= (int)slide.Shapes.Count; j++)
                {
                    dynamic shape = slide.Shapes[j];
                    string id = PowerPointLayout.ReadId((object)shape);
                    if (id == null) continue;
                    int nativeId = (int)shape.Id;
                    if (slideId <= 0 || nativeId <= 0)
                        throw new InvalidOperationException("对象的原生标识无效，未修改任何编号。");
                    string physicalKey = slideId.ToString(CultureInfo.InvariantCulture) + "/" +
                        nativeId.ToString(CultureInfo.InvariantCulture);
                    if (!physicalKeys.Add(physicalKey))
                        throw new InvalidOperationException("对象的原生标识重复，未修改任何编号。");
                    shapes.Add(new IdentityShape
                    {
                        Shape = (object)shape, Id = id, SlideId = slideId,
                        PhysicalKey = physicalKey,
                        Owner = (string)shape.Tags[OwnerTagName]
                    });
                }
            }
            if (shapes.Count == 0) return 0;

            string oldDocumentId = (string)presentation.Tags[DocumentTagName];
            string documentId = String.IsNullOrEmpty(oldDocumentId)
                ? Guid.NewGuid().ToString("D").ToUpperInvariant()
                : ShapeIdentity.Normalize(oldDocumentId);
            ConstraintDocument document = PowerPointLayout.ReadDocument(presentationObject);
            var references = document.Constraints
                .SelectMany(c => c.ShapeIds.Select(id => new { Id = id, c.SlideId }))
                .GroupBy(r => r.Id, StringComparer.Ordinal)
                .ToDictionary(g => g.Key, g => new HashSet<int>(g.Select(r => r.SlideId)), StringComparer.Ordinal);
            var usedIds = new HashSet<string>(shapes.Select(s => s.Id), StringComparer.Ordinal);
            var changes = new List<TagChange>();
            if (!String.Equals(oldDocumentId, documentId, StringComparison.Ordinal))
                changes.Add(new TagChange(presentationObject, DocumentTagName, oldDocumentId, documentId));

            int reassigned = 0;
            foreach (var group in shapes.GroupBy(s => s.Id, StringComparer.Ordinal))
            {
                var members = group.ToList();
                var owned = members.Where(s => String.Equals(s.Owner, documentId + "/" + s.PhysicalKey,
                    StringComparison.Ordinal)).ToList();
                HashSet<int> referenceSlides;
                if (!references.TryGetValue(group.Key, out referenceSlides)) referenceSlides = new HashSet<int>();
                if (referenceSlides.Count > 1)
                    throw new InvalidOperationException("同一对象编号已被多页约束引用，无法自动确定归属，未修改任何编号。");
                if (owned.Count > 1)
                    throw new InvalidOperationException("多个对象声明了同一编号的归属，未修改任何编号。");

                IdentityShape original = owned.SingleOrDefault();
                if (original == null)
                {
                    // A missing owner can be adopted when the existing XML
                    // identifies exactly one candidate on its referenced slide.
                    // Without any references either candidate is safe to keep:
                    // no original constraint can be redirected by that choice.
                    var unowned = members.Where(s => String.IsNullOrEmpty(s.Owner)).ToList();
                    if (referenceSlides.Count == 0) original = unowned.FirstOrDefault();
                    else
                    {
                        int referencedSlide = referenceSlides.Single();
                        var candidates = unowned.Where(s => s.SlideId == referencedSlide).ToList();
                        if (candidates.Count != 1)
                            throw new InvalidOperationException("约束引用的对象编号重复或归属不明，无法安全识别原对象，未修改任何编号。");
                        original = candidates[0];
                    }
                }
                if (referenceSlides.Count != 0 && (original == null || !referenceSlides.Contains(original.SlideId)))
                    throw new InvalidOperationException("约束与对象的归属不一致，未修改任何编号。");

                foreach (IdentityShape member in members)
                {
                    if (!Object.ReferenceEquals(member, original))
                    {
                        string replacement;
                        do { replacement = Guid.NewGuid().ToString("D").ToUpperInvariant(); }
                        while (!usedIds.Add(replacement));
                        string oldId = (string)((dynamic)member.Shape).Tags[PowerPointLayout.TagName];
                        changes.Add(new TagChange(member.Shape, PowerPointLayout.TagName, oldId, replacement));
                        reassigned++;
                    }
                    string newOwner = documentId + "/" + member.PhysicalKey;
                    if (!String.Equals(member.Owner, newOwner, StringComparison.Ordinal))
                        changes.Add(new TagChange(member.Shape, OwnerTagName, member.Owner, newOwner));
                }
            }

            // Plan every change first. A rejected or ambiguous group leaves the
            // entire presentation untouched, including its document token.
            var attempted = new List<TagChange>();
            try
            {
                foreach (TagChange change in changes)
                {
                    attempted.Add(change);
                    ((dynamic)change.Target).Tags.Add(change.Name, change.After);
                }
            }
            catch (Exception writeError)
            {
                var errors = new List<Exception> { writeError };
                for (int i = attempted.Count - 1; i >= 0; i--)
                {
                    TagChange change = attempted[i];
                    try
                    {
                        if (String.IsNullOrEmpty(change.Before)) ((dynamic)change.Target).Tags.Delete(change.Name);
                        else ((dynamic)change.Target).Tags.Add(change.Name, change.Before);
                    }
                    catch (Exception restoreError) { errors.Add(restoreError); }
                }
                if (errors.Count > 1) throw new AggregateException("对象编号写入失败，部分标签无法恢复。", errors);
                throw;
            }
            return reassigned;
        }

        private sealed class IdentityShape
        {
            public object Shape;
            public string Id, PhysicalKey, Owner;
            public int SlideId;
        }

        private sealed class TagChange
        {
            public readonly object Target;
            public readonly string Name, Before, After;
            public TagChange(object target, string name, string before, string after)
            { Target = target; Name = name; Before = before; After = after; }
        }
    }
}
