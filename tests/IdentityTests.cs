using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using PptGcs.AddIn;
using PptGcs.Core;

internal static class IdentityTests
{
    private const string A = "AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA";
    private const string B = "BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB";
    private static int passed;

    private static void Main()
    {
        Run("copy gets a fresh GUID without changing original constraints or geometry", CopyRetainsOriginal);
        Run("multiple copies and copying a copy stay unique", MultipleCopies);
        Run("z-order changes do not transfer original identity to a copy", ReorderedCopies);
        Run("copy to another slide gets a fresh GUID", AcrossSlides);
        Run("copy to another document gets a fresh GUID despite identical native IDs", AcrossDocuments);
        Run("unbound duplicate GUIDs can be repaired without changing XML", UnboundDuplicates);
        Run("referenced slide disambiguates an unowned cross-slide copy", ReferencedSlide);
        Run("ambiguous constrained duplicates fail before any tag writes", AmbiguousDuplicates);
        Run("failed tag writes roll back all earlier identity changes", FailedWriteRollback);
        Run("repeated normalization and a new manager do not write tags", Idempotent);
        Run("ordinary untagged shapes do not dirty a presentation", UntaggedShapes);
        Console.WriteLine("PASS: " + passed + " identity tests");
    }

    private static IdentityPresentation Fixture(bool constrained = true)
    {
        string xml = new ConstraintDocument(constrained
            ? new[] { new LayoutConstraint("11111111-1111-1111-1111-111111111111", 256, ConstraintKind.SameY, new[] { A, B }) }
            : new LayoutConstraint[0]).ToXml();
        var presentation = new IdentityPresentation(xml);
        var slide = new IdentitySlide(256);
        slide.Shapes.Add(new IdentityShape(4, A));
        slide.Shapes.Add(new IdentityShape(5, B));
        presentation.Slides.Add(slide);
        return presentation;
    }

    private static void CopyRetainsOriginal()
    {
        var p = Fixture();
        var manager = new ShapeIdentityManager();
        manager.Normalize(p);
        string xml = p.CustomXMLParts[1].XML;
        var original = p.Slides[1].Shapes[1];
        var other = p.Slides[1].Shapes[2];
        string originalOwner = original.Tags[ShapeIdentityManager.OwnerTagName];
        var copy = original.Copy(6);
        p.Slides[1].Shapes.Add(copy);
        string geometry = Geometry(p);
        Assert(manager.Normalize(p) == 1, "Exactly one copy must receive a new GUID.");
        Assert(Id(original) == A && Id(other) == B, "An original GUID changed.");
        Assert(original.Tags[ShapeIdentityManager.OwnerTagName] == originalOwner, "Original ownership changed.");
        Assert(Id(copy) != A && Id(copy) != B, "Copy retained an existing GUID.");
        Assert(p.CustomXMLParts[1].XML == xml, "Constraint XML changed.");
        Assert(Geometry(p) == geometry, "Identity repair changed shape geometry.");
        AssertOwner(p, p.Slides[1], copy);
        Unique(p);
    }

    private static void MultipleCopies()
    {
        var p = Fixture();
        var manager = new ShapeIdentityManager();
        manager.Normalize(p);
        var shapes = p.Slides[1].Shapes;
        var original = shapes[1];
        shapes.Add(original.Copy(6));
        shapes.Add(original.Copy(7));
        Assert(manager.Normalize(p) == 2);
        var firstCopy = shapes[3];
        string copyId = Id(firstCopy);
        shapes.Add(firstCopy.Copy(8));
        Assert(manager.Normalize(p) == 1);
        Assert(Id(original) == A && Id(firstCopy) == copyId);
        Unique(p);
    }

    private static void ReorderedCopies()
    {
        var p = Fixture();
        var manager = new ShapeIdentityManager();
        manager.Normalize(p);
        var original = p.Slides[1].Shapes[1];
        var copy = original.Copy(6);
        p.Slides[1].Shapes.InsertFirst(copy);
        manager.Normalize(p);
        Assert(Id(original) == A && Id(copy) != A, "Collection order stole original identity.");
        Unique(p);
    }

    private static void AcrossSlides()
    {
        var p = Fixture();
        var manager = new ShapeIdentityManager();
        manager.Normalize(p);
        var original = p.Slides[1].Shapes[1];
        var copy = original.Copy(original.Id);
        var destination = new IdentitySlide(257);
        destination.Shapes.Add(copy);
        p.Slides.InsertFirst(destination);
        Assert(manager.Normalize(p) == 1);
        Assert(Id(original) == A && Id(copy) != A);
        AssertOwner(p, destination, copy);
        Unique(p);
    }

    private static void AcrossDocuments()
    {
        var source = Fixture();
        var manager = new ShapeIdentityManager();
        manager.Normalize(source);
        string sourceState = Tags(source);
        var original = source.Slides[1].Shapes[1];
        var destination = new IdentityPresentation(new ConstraintDocument(new LayoutConstraint[0]).ToXml());
        var slide = new IdentitySlide(source.Slides[1].SlideID);
        var copy = original.Copy(original.Id);
        slide.Shapes.Add(copy);
        destination.Slides.Add(slide);
        Assert(manager.Normalize(destination) == 1);
        Assert(Id(copy) != A && Tags(source) == sourceState);
        Assert(source.Tags[ShapeIdentityManager.DocumentTagName] != destination.Tags[ShapeIdentityManager.DocumentTagName]);
        AssertOwner(destination, slide, copy);
    }

    private static void UnboundDuplicates()
    {
        var p = Fixture(false);
        var original = p.Slides[1].Shapes[1];
        p.Slides[1].Shapes.Add(original.Copy(6));
        p.Slides[1].Shapes.Add(original.Copy(7));
        string xml = p.CustomXMLParts[1].XML;
        Assert(new ShapeIdentityManager().Normalize(p) == 2);
        Assert(Id(original) == A && p.CustomXMLParts[1].XML == xml);
        Unique(p);
    }

    private static void ReferencedSlide()
    {
        var p = Fixture();
        var original = p.Slides[1].Shapes[1];
        var otherSlide = new IdentitySlide(257);
        var copy = original.Copy(4);
        otherSlide.Shapes.Add(copy);
        p.Slides.InsertFirst(otherSlide);
        Assert(new ShapeIdentityManager().Normalize(p) == 1);
        Assert(Id(original) == A && Id(copy) != A);
    }

    private static void AmbiguousDuplicates()
    {
        var p = Fixture();
        p.Slides[1].Shapes.Add(p.Slides[1].Shapes[1].Copy(6));
        string before = Tags(p);
        int writes = Writes(p);
        Throws(() => new ShapeIdentityManager().Normalize(p));
        Assert(Tags(p) == before && Writes(p) == writes, "An ambiguous plan mutated document tags.");
    }

    private static void FailedWriteRollback()
    {
        foreach (bool afterWrite in new[] { false, true })
        {
            var p = Fixture(false);
            var copy = p.Slides[1].Shapes[1].Copy(6);
            p.Slides[1].Shapes.Add(copy);
            string before = Tags(p);
            string xml = p.CustomXMLParts[1].XML;
            copy.Tags.FailNextAddName = ShapeIdentityManager.OwnerTagName;
            copy.Tags.FailAfterAdd = afterWrite;
            Throws(() => new ShapeIdentityManager().Normalize(p));
            Assert(Tags(p) == before, "Partial identity changes survived a failed tag write.");
            Assert(p.CustomXMLParts[1].XML == xml);
            Assert(new ShapeIdentityManager().Normalize(p) == 1, "Repair was not retryable after rollback.");
            Unique(p);
        }
    }

    private static void Idempotent()
    {
        var p = Fixture();
        var manager = new ShapeIdentityManager();
        manager.Normalize(p);
        string before = Tags(p);
        int writes = Writes(p);
        Assert(manager.Normalize(p) == 0 && new ShapeIdentityManager().Normalize(p) == 0);
        Assert(Tags(p) == before && Writes(p) == writes, "Stable identities caused extra document writes.");
    }

    private static void UntaggedShapes()
    {
        var p = new IdentityPresentation(new ConstraintDocument(new LayoutConstraint[0]).ToXml());
        var slide = new IdentitySlide(256);
        slide.Shapes.Add(new IdentityShape(4, null));
        p.Slides.Add(slide);
        Assert(new ShapeIdentityManager().Normalize(p) == 0 && Writes(p) == 0);
    }

    private static string Id(IdentityShape shape) { return shape.Tags[PowerPointLayout.TagName]; }
    private static void Unique(IdentityPresentation presentation)
    {
        var ids = presentation.Slides.SelectMany(s => s.Shapes).Select(Id).Where(id => id.Length != 0).ToList();
        Assert(ids.All(id => Guid.TryParse(id, out Guid value) && value != Guid.Empty), "Invalid generated GUID.");
        Assert(ids.Distinct(StringComparer.OrdinalIgnoreCase).Count() == ids.Count, "Duplicate GUID remains.");
    }
    private static void AssertOwner(IdentityPresentation p, IdentitySlide slide, IdentityShape shape)
    {
        string expected = p.Tags[ShapeIdentityManager.DocumentTagName] + "/" + slide.SlideID + "/" + shape.Id;
        Assert(shape.Tags[ShapeIdentityManager.OwnerTagName] == expected, "Owner does not identify the current physical shape.");
    }
    private static string Tags(IdentityPresentation p)
    {
        return p.Tags.Snapshot() + "|" + String.Join("|", p.Slides.SelectMany(s => s.Shapes.Select(sh =>
            s.SlideID + "/" + sh.Id + ":" + sh.Tags.Snapshot())));
    }
    private static int Writes(IdentityPresentation p) { return p.Tags.Writes + p.Slides.Sum(s => s.Shapes.Sum(sh => sh.Tags.Writes)); }
    private static string Geometry(IdentityPresentation p)
    {
        return String.Join("|", p.Slides.SelectMany(s => s.Shapes.Select(sh => String.Join(",", new[]
            { sh.Left, sh.Top, sh.Width, sh.Height }.Select(v => v.ToString("R", CultureInfo.InvariantCulture))))));
    }
    private static void Run(string name, Action action) { action(); passed++; Console.WriteLine("PASS " + name); }
    private static void Assert(bool condition, string message = "Identity assertion failed.") { if (!condition) throw new Exception(message); }
    private static void Throws(Action action)
    {
        try { action(); }
        catch (InvalidOperationException) { return; }
        throw new Exception("Expected identity normalization to reject this operation.");
    }
}

// Public types let the production DLL's dynamic binder use the same surface as
// the Office object model. The fixture intentionally exposes no XML mutation API.
public sealed class IdentityPresentation
{
    public IdentityTags Tags { get; private set; } = new IdentityTags();
    public IdentityItems<IdentitySlide> Slides { get; private set; } = new IdentityItems<IdentitySlide>();
    public IdentityXmlParts CustomXMLParts { get; private set; }
    public IdentityPresentation(string xml) { CustomXMLParts = new IdentityXmlParts(xml); }
}

public sealed class IdentitySlide
{
    public int SlideID { get; private set; }
    public IdentityItems<IdentityShape> Shapes { get; private set; } = new IdentityItems<IdentityShape>();
    public IdentitySlide(int id) { SlideID = id; }
}

public sealed class IdentityShape
{
    public int Id { get; private set; }
    public IdentityTags Tags { get; private set; } = new IdentityTags();
    public double Left { get; set; } = 100;
    public double Top { get; set; } = 200;
    public double Width { get; set; } = 120;
    public double Height { get; set; } = 80;
    public IdentityShape(int id, string guid)
    {
        Id = id;
        if (guid != null) Tags.Add(PowerPointLayout.TagName, guid);
    }
    public IdentityShape Copy(int id)
    {
        return new IdentityShape(id, null) { Tags = Tags.Clone(), Left = Left, Top = Top, Width = Width, Height = Height };
    }
}

public sealed class IdentityTags
{
    private readonly Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    public int Writes { get; private set; }
    public string FailNextAddName { get; set; }
    public bool FailAfterAdd { get; set; }
    public string this[string name] { get { return values.TryGetValue(name, out string value) ? value : ""; } }
    public void Add(string name, string value)
    {
        bool fail = String.Equals(name, FailNextAddName, StringComparison.OrdinalIgnoreCase);
        if (fail) FailNextAddName = null;
        if (fail && !FailAfterAdd)
        {
            throw new InvalidOperationException("Injected one-shot Office tag-write failure.");
        }
        Writes++;
        values[name.ToUpperInvariant()] = value.ToUpperInvariant();
        if (fail) throw new InvalidOperationException("Injected failure after Office accepted the tag write.");
    }
    public void Delete(string name) { Writes++; values.Remove(name); }
    public string Snapshot() { return String.Join(";", values.OrderBy(p => p.Key, StringComparer.Ordinal).Select(p => p.Key + "=" + p.Value)); }
    public IdentityTags Clone()
    {
        var result = new IdentityTags();
        foreach (var item in values) result.values.Add(item.Key, item.Value);
        return result;
    }
}

public sealed class IdentityItems<T> : IEnumerable<T>
{
    private readonly List<T> values = new List<T>();
    public int Count { get { return values.Count; } }
    public T this[int index] { get { return values[index-1]; } }
    public void Add(T value) { values.Add(value); }
    public void InsertFirst(T value) { values.Insert(0, value); }
    public IEnumerator<T> GetEnumerator() { return values.GetEnumerator(); }
    IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
}

public sealed class IdentityXmlPart
{
    public string XML { get; private set; }
    public IdentityXmlPart(string xml) { XML = xml; }
}

public sealed class IdentityXmlParts
{
    private readonly IdentityXmlPart part;
    public int Count { get { return 1; } }
    public IdentityXmlPart this[int index] { get { if (index != 1) throw new IndexOutOfRangeException(); return part; } }
    public IdentityXmlParts(string xml) { part = new IdentityXmlPart(xml); }
    public IdentityXmlParts SelectByNamespace(string value)
    {
        if (value != ConstraintDocument.NamespaceUri) throw new InvalidOperationException("Unexpected XML namespace.");
        return this;
    }
}
