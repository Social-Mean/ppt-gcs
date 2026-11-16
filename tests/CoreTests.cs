using System;
using System.Globalization;
using System.Threading;
using PptGcs.Core;

internal static class CoreTests
{
    private const string A = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa";
    private const string B = "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb";
    private const string C = "cccccccc-cccc-cccc-cccc-cccccccccccc";
    private const string Id = "11111111-1111-1111-1111-111111111111";
    private const string Id2 = "22222222-2222-2222-2222-222222222222";
    private static int passed;

    private static void Main()
    {
        Run("round-trip independent of culture", RoundTrip);
        Run("drive either shape without coordinate drift", Bidirectional);
        Run("shape identity survives PowerPoint tag casing", TagCasing);
        Run("reject missing or same shape identities", InvalidIdentities);
        Run("reject duplicate constraint IDs and reversed pairs", Duplicates);
        Run("reject unsupported or malformed documents", InvalidXml);
        Run("reject DTD before resolving external entities", NoDtd);
        Run("reject non-finite coordinates and unknown driver", InvalidEdit);
        Run("ignore float noise but preserve meaningful edits", Tolerance);
        Console.WriteLine("PASS: " + passed + " core tests");
    }

    private static SameTopConstraint Pair()
    {
        return new SameTopConstraint(Id, 256, A, B);
    }

    private static void Run(string name, Action action)
    {
        action();
        passed++;
        Console.WriteLine("PASS " + name);
    }

    private static void Assert(bool condition)
    {
        if (!condition) throw new Exception("Assertion failed.");
    }

    private static void Throws<T>(Action action) where T : Exception
    {
        try { action(); }
        catch (T) { return; }
        throw new Exception("Expected " + typeof(T).Name);
    }

    private static void RoundTrip()
    {
        var culture = Thread.CurrentThread.CurrentCulture;
        try
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");
            var source = new ConstraintDocument(new[] { Pair() });
            var result = ConstraintDocument.Parse(source.ToXml());
            Assert(result.Constraints.Count == 1);
            Assert(result.Constraints[0].SlideId == 256);
            Assert(result.Constraints[0].ShapeA == A.ToUpperInvariant());
            Assert(result.ToXml() == source.ToXml());
            Assert(ConstraintDocument.Parse(new ConstraintDocument(new SameTopConstraint[0]).ToXml()).Constraints.Count == 0);
        }
        finally { Thread.CurrentThread.CurrentCulture = culture; }
    }

    private static void Bidirectional()
    {
        var pair = Pair();
        var fromA = SameTopSolver.Solve(pair, A, 123.125);
        Assert(fromA.ShapeId == B.ToUpperInvariant() && fromA.Top == 123.125);
        var fromB = SameTopSolver.Solve(pair, B, -15.75);
        Assert(fromB.ShapeId == A.ToUpperInvariant() && fromB.Top == -15.75);
        for (int i = 0; i < 1000; i++)
        {
            var result = SameTopSolver.Solve(pair, i % 2 == 0 ? A : B, 123.125);
            Assert(result.Top == 123.125);
        }
    }

    private static void TagCasing()
    {
        Assert(ShapeIdentity.Normalize("{" + A.ToUpperInvariant() + "}") == A.ToUpperInvariant());
        Assert(SameTopSolver.Solve(Pair(), A.ToUpperInvariant(), 10).ShapeId == B.ToUpperInvariant());
    }

    private static void InvalidIdentities()
    {
        Throws<ArgumentException>(() => new SameTopConstraint(Id, 256, A, A.ToUpperInvariant()));
        Throws<ArgumentException>(() => new SameTopConstraint(Id, 256, "", B));
        Throws<ArgumentException>(() => new SameTopConstraint(Id, 256, Guid.Empty.ToString(), B));
        Throws<ArgumentOutOfRangeException>(() => new SameTopConstraint(Id, 0, A, B));
    }

    private static void Duplicates()
    {
        Throws<ArgumentException>(() => new ConstraintDocument(new[] { Pair(), new SameTopConstraint(Id, 256, B, C) }));
        Throws<ArgumentException>(() => new ConstraintDocument(new[] { Pair(), new SameTopConstraint(Id2, 256, B, A) }));
    }

    private static void InvalidXml()
    {
        Throws<FormatException>(() => ConstraintDocument.Parse(""));
        Throws<FormatException>(() => ConstraintDocument.Parse("<constraints>"));
        Throws<FormatException>(() => ConstraintDocument.Parse("<constraints xmlns='urn:other'/>"));
        Throws<FormatException>(() => ConstraintDocument.Parse("<constraints xmlns='urn:ppt-gcs:v1'><sameWidth/></constraints>"));
        Throws<FormatException>(() => ConstraintDocument.Parse("<constraints xmlns='urn:ppt-gcs:v1'><sameTop/></constraints>"));
        var xml = new ConstraintDocument(new[] { Pair() }).ToXml();
        Throws<FormatException>(() => ConstraintDocument.Parse(xml.Replace("slideId=\"256\"", "slideId=\"-1\"")));
        Throws<FormatException>(() => ConstraintDocument.Parse(xml.Replace("shapeB=\"" + B.ToUpperInvariant(), "shapeB=\"" + A.ToUpperInvariant())));
        Throws<FormatException>(() => ConstraintDocument.Parse(xml.Replace("<sameTop ", "<sameTop other=\"x\" ")));
    }

    private static void NoDtd()
    {
        Throws<FormatException>(() => ConstraintDocument.Parse("<!DOCTYPE constraints [<!ENTITY external SYSTEM 'file:///not-read'>]><constraints xmlns='urn:ppt-gcs:v1'>&external;</constraints>"));
    }

    private static void InvalidEdit()
    {
        Throws<ArgumentException>(() => SameTopSolver.Solve(Pair(), C, 10));
        Throws<ArgumentOutOfRangeException>(() => SameTopSolver.Solve(Pair(), A, Double.NaN));
        Throws<ArgumentOutOfRangeException>(() => SameTopSolver.Solve(Pair(), B, Double.PositiveInfinity));
        Throws<ArgumentOutOfRangeException>(() => SameTopSolver.NeedsUpdate(Double.NegativeInfinity, 10));
    }

    private static void Tolerance()
    {
        Assert(!SameTopSolver.NeedsUpdate(123.000001, 123.0));
        Assert(SameTopSolver.NeedsUpdate(123.0, 123.02));
        Assert(SameTopSolver.NeedsUpdate(123.0, 122.98));
    }
}
