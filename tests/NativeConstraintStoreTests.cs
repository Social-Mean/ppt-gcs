using System;
using System.Collections.Generic;
using PptGcs.AddIn;
using PptGcs.Core;

internal static class NativeConstraintStoreTests
{
    private static int passed;
    private static ConstraintDocument Document()
    {
        return new ConstraintDocument(new[] { new LayoutConstraint("11111111-1111-1111-1111-111111111111",
            256, ConstraintKind.SameY, new[] { "AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA", "BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB" }) });
    }
    private static void Main()
    {
        Run("uppercase-safe Unicode codec preserves exact XML bytes", UnicodeCodec);
        Run("constraint document writes and reads through native tags", DocumentRoundTrip);
        Run("empty documents remain explicit undoable native tag values", EmptyDocument);
        Run("reading an unmarked document is empty and read-only", MissingTag);
        Run("obsolete CustomXML storage is rejected without fallback", RejectOldStorage);
        Run("invalid hexadecimal and UTF-8 fail closed", InvalidEncoding);
        Run("one MiB UTF-8 boundary is enforced before writing", SizeLimit);
        Run("decoded XML still requires the strict constraint schema", InvalidXml);
        Run("identical write and repeated reads do not modify tags", ReadOnlyAndNoop);
        Run("failed initial and replacement writes restore original tags", FailedWrite);
        Run("truncated readback is rejected and rolled back", BadReadback);
        Console.WriteLine("PASS: " + passed + " native constraint store tests");
    }
    private static void UnicodeCodec()
    {
        const string xml = "<data Name=\"MiXeD\">中文 &amp; café 😀</data>";
        string encoded = NativeConstraintStore.EncodeXml(xml);
        Assert(encoded == encoded.ToUpperInvariant());
        Assert(NativeConstraintStore.DecodeXml(encoded) == xml);
        Throws<FormatException>(() => NativeConstraintStore.EncodeXml("\uD800"));
    }
    private static void DocumentRoundTrip()
    {
        var p = new StorePresentation();
        var document = Document();
        NativeConstraintStore.Write(p, document);
        Assert(p.Tags.Writes == 1);
        Assert(NativeConstraintStore.Read(p).ToXml() == document.ToXml());
        Assert(p.CustomXMLParts.Reads == 0, "Native tagged documents should not read old storage.");
    }
    private static void EmptyDocument()
    {
        var p = new StorePresentation();
        NativeConstraintStore.Write(p, Document());
        var empty = new ConstraintDocument(new LayoutConstraint[0]);
        NativeConstraintStore.Write(p, empty);
        Assert(p.Tags[NativeConstraintStore.TagName].Length > 0);
        Assert(NativeConstraintStore.Read(p).Constraints.Count == 0 && p.Tags.Deletes == 0);
    }
    private static void MissingTag()
    {
        var p = new StorePresentation();
        Assert(NativeConstraintStore.Read(p).Constraints.Count == 0);
        Assert(p.Tags.Writes == 0 && p.Tags.Deletes == 0);
    }
    private static void RejectOldStorage()
    {
        var p = new StorePresentation { CustomXMLParts = new StoreXmlParts(1) };
        Throws<InvalidOperationException>(() => NativeConstraintStore.Read(p));
        Assert(p.Tags.Writes == 0 && p.Tags.Deletes == 0);
    }
    private static void InvalidEncoding()
    {
        foreach (string value in new[] { "", "A", "ZZ", "a0", "C0AF", "FF" })
            Throws<FormatException>(() => NativeConstraintStore.DecodeXml(value));
        Throws<ArgumentNullException>(() => NativeConstraintStore.DecodeXml(null));
        Throws<ArgumentNullException>(() => NativeConstraintStore.EncodeXml(null));
    }
    private static void SizeLimit()
    {
        string boundary = new string('A', NativeConstraintStore.MaxXmlBytes);
        Assert(NativeConstraintStore.DecodeXml(NativeConstraintStore.EncodeXml(boundary)) == boundary);
        Throws<FormatException>(() => NativeConstraintStore.EncodeXml(boundary + "A"));
        Throws<FormatException>(() => NativeConstraintStore.EncodeXml(new string('é', NativeConstraintStore.MaxXmlBytes / 2 + 1)));
        Throws<FormatException>(() => NativeConstraintStore.DecodeXml(new string('0', NativeConstraintStore.MaxXmlBytes * 2 + 2)));
    }
    private static void InvalidXml()
    {
        var p = new StorePresentation();
        p.Tags.Seed(NativeConstraintStore.TagName, NativeConstraintStore.EncodeXml("<bogus/>"));
        Throws<FormatException>(() => NativeConstraintStore.Read(p));
        p.Tags.Seed(NativeConstraintStore.TagName, "ZZ");
        Throws<FormatException>(() => NativeConstraintStore.Read(p));
        Assert(p.Tags.Writes == 0);
    }
    private static void ReadOnlyAndNoop()
    {
        var p = new StorePresentation();
        var document = Document();
        NativeConstraintStore.Write(p, document);
        int writes = p.Tags.Writes;
        NativeConstraintStore.Write(p, document);
        NativeConstraintStore.Read(p); NativeConstraintStore.Read(p);
        Assert(p.Tags.Writes == writes && p.Tags.Deletes == 0);
    }
    private static void FailedWrite()
    {
        foreach (bool existing in new[] { false, true })
        foreach (bool afterWrite in new[] { false, true })
        {
            var p = new StorePresentation();
            if (existing) NativeConstraintStore.Write(p, new ConstraintDocument(new LayoutConstraint[0]));
            string before = p.Tags[NativeConstraintStore.TagName];
            p.Tags.FailNext = true; p.Tags.FailAfterWrite = afterWrite;
            Throws<InvalidOperationException>(() => NativeConstraintStore.Write(p, Document()));
            Assert(p.Tags[NativeConstraintStore.TagName] == before, "Failed write did not restore original tag.");
        }
    }
    private static void BadReadback()
    {
        var p = new StorePresentation();
        NativeConstraintStore.Write(p, new ConstraintDocument(new LayoutConstraint[0]));
        string before = p.Tags[NativeConstraintStore.TagName];
        p.Tags.TruncateNext = true;
        Throws<InvalidOperationException>(() => NativeConstraintStore.Write(p, Document()));
        Assert(p.Tags[NativeConstraintStore.TagName] == before);
    }
    private static void Run(string name, Action test) { test(); passed++; Console.WriteLine("PASS " + name); }
    private static void Assert(bool condition, string message = "Store assertion failed.") { if (!condition) throw new Exception(message); }
    private static void Throws<T>(Action action) where T : Exception
    {
        try { action(); } catch (T) { return; }
        throw new Exception("Expected " + typeof(T).Name);
    }
}

public sealed class StorePresentation
{
    public StoreTags Tags { get; private set; } = new StoreTags();
    public StoreXmlParts CustomXMLParts { get; set; } = new StoreXmlParts(0);
}
public sealed class StoreTags
{
    private readonly Dictionary<string, string> values = new Dictionary<string, string>();
    public int Writes { get; private set; }
    public int Deletes { get; private set; }
    public bool FailNext, FailAfterWrite, TruncateNext;
    public string this[string name] { get { return values.TryGetValue(name, out string value) ? value : String.Empty; } }
    public void Seed(string name, string value) { values[name] = value; }
    public void Add(string name, string value)
    {
        bool fail = FailNext; FailNext = false;
        if (fail && !FailAfterWrite) throw new InvalidOperationException("Injected Office failure.");
        Writes++;
        values[name] = TruncateNext ? value.Substring(0, value.Length / 2).ToUpperInvariant() : value.ToUpperInvariant();
        TruncateNext = false;
        if (fail) throw new InvalidOperationException("Injected failure after write.");
    }
    public void Delete(string name) { Deletes++; values.Remove(name); }
}
public sealed class StoreXmlParts
{
    public int Count { get; private set; }
    public int Reads { get; private set; }
    public StoreXmlParts(int count) { Count = count; }
    public StoreXmlParts SelectByNamespace(string value)
    {
        if (value != "urn:ppt-gcs:v2") throw new Exception("Unexpected legacy namespace.");
        Reads++; return this;
    }
}
