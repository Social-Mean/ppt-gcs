using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using PptGcs.Core;

internal static class CoreTests
{
    private const string A = "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa";
    private const string B = "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb";
    private const string C = "cccccccc-cccc-cccc-cccc-cccccccccccc";
    private const string D = "dddddddd-dddd-dddd-dddd-dddddddddddd";
    private const string E = "eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee";
    private static int passed;
    private static int serial;

    private static void Main()
    {
        Run("round-trip every constraint and preserve gap order", RoundTrip);
        Run("reject obsolete v1 without migration", RejectV1);
        Run("reject malformed and unsupported XML", InvalidXml);
        Run("reject DTD before resolving external entities", NoDtd);
        Run("validate IDs, kinds, group sizes and duplicates", InvalidConstraints);
        Run("reject duplicate relations but allow different kinds", DuplicateRelations);
        Run("drive same X from either member", () => Bidirectional(ConstraintKind.SameX));
        Run("drive same Y from either member", () => Bidirectional(ConstraintKind.SameY));
        Run("resize same width from either member", () => Bidirectional(ConstraintKind.SameWidth));
        Run("resize same height from either member", () => Bidirectional(ConstraintKind.SameHeight));
        Run("normalize unequal dimensions with selected driver", InitialDimensions);
        Run("propagate multiple dimension constraints through chains", SizeChains);
        Run("preserve disconnected geometry and constraints", Disconnected);
        Run("three unequal-width horizontal gaps use edges", HorizontalGaps);
        Run("three unequal-height vertical gaps use edges", VerticalGaps);
        Run("four-shape gaps preserve order with any driver", FourShapeGaps);
        Run("resize inside equal-gap layout preserves driver anchors", GapResize);
        Run("combine alignment, size and gap constraints", Combined);
        Run("minimum-movement solution preserves free coordinates", MinimumMovement);
        Run("redundant compatible equations remain solvable", Redundant);
        Run("reject conflicting required equations atomically", Conflicting);
        Run("large translated coordinates remain stable", TranslationStability);
        Run("repeated edits do not drift", Repeated);
        Run("validate finite and positive geometry", InvalidGeometry);
        Run("reject missing, duplicate and mismatched edit identities", InvalidEdit);
        Run("no-constraint edit affects only its driver", FreeEdit);
        Run("readback residual detects every constraint kind", Residual);
        Run("ignore float writeback noise but retain actual changes", Tolerance);
        Console.WriteLine("PASS: " + passed + " core tests");
    }

    private static ShapeGeometry Shape(string id, double x = 0, double y = 0, double w = 20, double h = 30)
    { return new ShapeGeometry(id, x, y, w, h); }

    private static LayoutConstraint Constraint(ConstraintKind kind, params string[] shapes)
    {
        serial++;
        return new LayoutConstraint(serial.ToString("x8") + "-0000-0000-0000-000000000001", 256, kind, shapes);
    }

    private static Dictionary<string, ShapeGeometry> Solve(IEnumerable<ShapeGeometry> input,
        IEnumerable<LayoutConstraint> constraints, ShapeGeometry target)
    { return LayoutSolver.Solve(input, constraints, target.Id, target).ToDictionary(s => s.Id); }

    private static ShapeGeometry Get(IDictionary<string, ShapeGeometry> state, string id)
    { return state[ShapeIdentity.Normalize(id)]; }

    private static void Run(string name, Action action)
    { action(); passed++; Console.WriteLine("PASS " + name); }

    private static void Assert(bool condition, string message = "Assertion failed.")
    { if (!condition) throw new Exception(message); }

    private static void Near(double actual, double expected, double epsilon = 0.000001)
    { Assert(!Double.IsNaN(actual) && Math.Abs(actual - expected) <= epsilon, "Expected " + expected + ", got " + actual); }

    private static void Same(ShapeGeometry a, ShapeGeometry b)
    { Assert(a.Id == b.Id); Near(a.Left, b.Left); Near(a.Top, b.Top); Near(a.Width, b.Width); Near(a.Height, b.Height); }

    private static void Throws<T>(Action action) where T : Exception
    {
        try { action(); } catch (T) { return; }
        throw new Exception("Expected " + typeof(T).Name);
    }

    private static void RoundTrip()
    {
        var prior = Thread.CurrentThread.CurrentCulture;
        try
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("fr-FR");
            var all = Enum.GetValues(typeof(ConstraintKind)).Cast<ConstraintKind>()
                .Select(kind => Constraint(kind, C, A, B)).ToArray();
            var source = new ConstraintDocument(all);
            var result = ConstraintDocument.Parse(source.ToXml());
            Assert(result.Constraints.Count == 6);
            Assert(result.Constraints[5].ShapeIds.SequenceEqual(new[] { C, A, B }.Select(ShapeIdentity.Normalize)));
            Assert(source.ToXml() == result.ToXml());
            Assert(ConstraintDocument.Parse(new ConstraintDocument(new LayoutConstraint[0]).ToXml()).Constraints.Count == 0);
        }
        finally { Thread.CurrentThread.CurrentCulture = prior; }
    }

    private static void RejectV1()
    {
        Throws<FormatException>(() => ConstraintDocument.Parse("<constraints xmlns='urn:ppt-gcs:v1'/>"));
        Throws<FormatException>(() => ConstraintDocument.Parse("<constraints xmlns='urn:ppt-gcs:v1'><sameTop/></constraints>"));
    }

    private static void InvalidXml()
    {
        string xml = new ConstraintDocument(new[] { Constraint(ConstraintKind.SameX, A, B) }).ToXml();
        foreach (string invalid in new[] { "", "<constraints>", "<constraints xmlns='other'/>",
            "<constraints xmlns='urn:ppt-gcs:v2'><sameX/></constraints>",
            xml.Replace("kind=\"SameX\"", "kind=\"0\""),
            xml.Replace("kind=\"SameX\"", "kind=\"sameX\""),
            xml.Replace("slideId=\"256\"", "slideId=\"-1\""),
            xml.Replace("<constraint ", "<constraint other=\"1\" "),
            xml.Replace("<shape ", "<shape other=\"1\" "),
            xml.Replace("/><shape", ">unexpected</shape><shape") })
            Throws<FormatException>(() => ConstraintDocument.Parse(invalid));
    }

    private static void NoDtd()
    {
        Throws<FormatException>(() => ConstraintDocument.Parse("<!DOCTYPE constraints [<!ENTITY x SYSTEM 'file:///not-read'>]><constraints xmlns='urn:ppt-gcs:v2'>&x;</constraints>"));
    }

    private static void InvalidConstraints()
    {
        Throws<ArgumentException>(() => Constraint(ConstraintKind.SameX, A, A.ToUpperInvariant()));
        Throws<ArgumentException>(() => Constraint(ConstraintKind.SameX, A));
        Throws<ArgumentException>(() => Constraint(ConstraintKind.EqualHorizontalGap, A, B));
        Throws<ArgumentException>(() => Constraint(ConstraintKind.EqualVerticalGap, A, B, C, A));
        Throws<ArgumentException>(() => Constraint(ConstraintKind.SameY, A, Guid.Empty.ToString()));
        Throws<ArgumentException>(() => Constraint(ConstraintKind.SameWidth, A, "invalid"));
        Throws<ArgumentOutOfRangeException>(() => new LayoutConstraint(Guid.NewGuid().ToString(), 0, ConstraintKind.SameY, new[] { A, B }));
        Throws<ArgumentOutOfRangeException>(() => Constraint((ConstraintKind)99, A, B));
        Assert(ShapeIdentity.Normalize("{" + A + "}") == A.ToUpperInvariant());
    }

    private static void DuplicateRelations()
    {
        var first = Constraint(ConstraintKind.SameX, A, B, C);
        Throws<ArgumentException>(() => new ConstraintDocument(new[] { first, new LayoutConstraint(first.Id, 256, ConstraintKind.SameY, new[] { A, B }) }));
        Throws<ArgumentException>(() => new ConstraintDocument(new[] { first, Constraint(ConstraintKind.SameX, C, B, A) }));
        Assert(new ConstraintDocument(new[] { first, Constraint(ConstraintKind.SameY, A, B, C) }).Constraints.Count == 2);
        var ordered = Constraint(ConstraintKind.EqualHorizontalGap, A, B, C);
        Throws<ArgumentException>(() => new ConstraintDocument(new[] { ordered, Constraint(ConstraintKind.EqualHorizontalGap, A, B, C) }));
        Assert(new ConstraintDocument(new[] { ordered, Constraint(ConstraintKind.EqualHorizontalGap, C, B, A) }).Constraints.Count == 2);
    }

    private static void Bidirectional(ConstraintKind kind)
    {
        foreach (string driver in new[] { A, B })
        {
            var initial = new[] { Shape(A, 10, 20, 40, 30), Shape(B, 100, 80, 40, 30) };
            var target = Shape(driver, driver == A ? 25 : 105, driver == A ? 22 : 95,
                kind == ConstraintKind.SameWidth ? 73 : 40, kind == ConstraintKind.SameHeight ? 59 : 30);
            var relation = Constraint(kind, A, B);
            var result = Solve(initial, new[] { relation }, target);
            Same(Get(result, driver), target);
            var follower = Get(result, driver == A ? B : A);
            var oldFollower = initial[driver == A ? 1 : 0];
            Near(follower.Left, kind == ConstraintKind.SameX ? target.Left : oldFollower.Left);
            Near(follower.Top, kind == ConstraintKind.SameY ? target.Top : oldFollower.Top);
            Near(follower.Width, kind == ConstraintKind.SameWidth ? target.Width : oldFollower.Width);
            Near(follower.Height, kind == ConstraintKind.SameHeight ? target.Height : oldFollower.Height);
            Near(LayoutSolver.MaxResidual(result.Values, new[] { relation }), 0);
        }
    }

    private static void InitialDimensions()
    {
        var initial = new[] { Shape(A, 10, 20, 40, 30), Shape(B, 100, 80, 90, 70) };
        var relations = new[] { Constraint(ConstraintKind.SameWidth, A, B), Constraint(ConstraintKind.SameHeight, A, B) };
        var result = Solve(initial, relations, initial[1]);
        Near(Get(result, A).Width, 90); Near(Get(result, A).Height, 70);
        Near(Get(result, A).Left, 10); Near(Get(result, A).Top, 20);
    }

    private static void SizeChains()
    {
        var initial = new[] { Shape(A, 0), Shape(B, 100), Shape(C, 200), Shape(D, 300) };
        var relations = new[] { Constraint(ConstraintKind.SameWidth, A, B), Constraint(ConstraintKind.SameWidth, B, C),
            Constraint(ConstraintKind.SameHeight, B, C), Constraint(ConstraintKind.SameY, C, D) };
        var result = Solve(initial, relations, Shape(C, 210, 50, 55, 75));
        foreach (string id in new[] { A, B, C }) Near(Get(result, id).Width, 55);
        Near(Get(result, A).Height, 30); Near(Get(result, B).Height, 75); Near(Get(result, C).Height, 75);
        Near(Get(result, D).Width, 20); Near(Get(result, D).Height, 30); Near(Get(result, D).Top, 50);
        Near(Get(result, B).Top, 0);
    }

    private static void Disconnected()
    {
        var initial = new[] { Shape(A), Shape(B, 50), Shape(C, 100, 70), Shape(D, 150, 70), Shape(E, 400, 300) };
        var relations = new[] { Constraint(ConstraintKind.SameX, A, B), Constraint(ConstraintKind.SameY, C, D) };
        var result = Solve(initial, relations, Shape(A, 25));
        foreach (var shape in initial.Skip(2)) Same(Get(result, shape.Id), shape);
    }

    private static ShapeGeometry[] Row()
    { return new[] { Shape(A, 0, 10, 20, 20), Shape(B, 50, 60, 30, 30), Shape(C, 110, 120, 40, 40) }; }

    private static void HorizontalGaps()
    {
        var initial = Row();
        var relation = Constraint(ConstraintKind.EqualHorizontalGap, A, B, C);
        var end = Solve(initial, new[] { relation }, Shape(A, 20, 10, 20, 20));
        Near(Get(end, B).Left, 58); Near(Get(end, C).Left, 106);
        foreach (var shape in initial) Near(Get(end, shape.Id).Top, shape.Top);
        var middle = Solve(initial, new[] { relation }, Shape(B, 70, 60, 30, 30));
        Near(Get(middle, A).Left, 20); Near(Get(middle, C).Left, 130);
        Near(LayoutSolver.MaxResidual(end.Values, new[] { relation }), 0);
        Near(LayoutSolver.MaxResidual(middle.Values, new[] { relation }), 0);
    }

    private static void VerticalGaps()
    {
        var initial = new[] { Shape(A, 10, 0, 20, 20), Shape(B, 60, 50, 30, 30), Shape(C, 120, 110, 40, 40) };
        var relation = Constraint(ConstraintKind.EqualVerticalGap, A, B, C);
        var result = Solve(initial, new[] { relation }, Shape(C, 120, 90, 40, 40));
        Near(Get(result, A).Top, 4); Near(Get(result, B).Top, 42);
        foreach (var shape in initial) Near(Get(result, shape.Id).Left, shape.Left);
        Near(LayoutSolver.MaxResidual(result.Values, new[] { relation }), 0);
    }

    private static void FourShapeGaps()
    {
        var initial = new[] { Shape(A, 0, 0, 20), Shape(B, 50, 20, 30), Shape(C, 110, 40, 40), Shape(D, 180, 60, 50) };
        var relation = Constraint(ConstraintKind.EqualHorizontalGap, A, B, C, D);
        foreach (var driver in initial)
        {
            var target = Shape(driver.Id, driver.Left + 35, driver.Top, driver.Width, driver.Height);
            var result = Solve(initial, new[] { relation }, target);
            Same(Get(result, driver.Id), target);
            Near(LayoutSolver.MaxResidual(result.Values, new[] { relation }), 0);
            foreach (var shape in initial) { Near(Get(result, shape.Id).Width, shape.Width); Near(Get(result, shape.Id).Top, shape.Top); }
        }
    }

    private static void GapResize()
    {
        var initial = Row();
        var relation = Constraint(ConstraintKind.EqualHorizontalGap, A, B, C);
        var target = Shape(B, 30, 60, 50, 30); // left handle: opposite edge remains x=80.
        var result = Solve(initial, new[] { relation }, target);
        Same(Get(result, B), target); Near(Get(result, B).Left + Get(result, B).Width, 80);
        Near(Get(result, A).Width, 20); Near(Get(result, C).Width, 40);
        Near(LayoutSolver.MaxResidual(result.Values, new[] { relation }), 0);
    }

    private static void Combined()
    {
        var initial = new[] { Shape(A, 0, 20, 20, 30), Shape(B, 50, 20, 20, 30), Shape(C, 100, 20, 20, 30) };
        var relations = new[] { Constraint(ConstraintKind.SameY, A, B, C), Constraint(ConstraintKind.SameWidth, A, B, C),
            Constraint(ConstraintKind.SameHeight, A, B, C), Constraint(ConstraintKind.EqualHorizontalGap, A, B, C) };
        var target = Shape(B, 60, 35, 35, 45);
        var result = Solve(initial, relations, target);
        Same(Get(result, B), target);
        Near(LayoutSolver.MaxResidual(result.Values, relations), 0);
        foreach (var shape in result.Values) { Near(shape.Top, 35); Near(shape.Width, 35); Near(shape.Height, 45); }
    }

    private static void MinimumMovement()
    {
        var initial = Row();
        var relation = Constraint(ConstraintKind.EqualHorizontalGap, A, B, C);
        var result = Solve(initial, new[] { relation }, Shape(A, 20, 10, 20, 20));
        double b = Get(result, B).Left, c = Get(result, C).Left;
        double cost = Math.Pow(b - 50, 2) + Math.Pow(c - 110, 2);
        // Every feasible change in the remaining free direction is (t, 2t).
        foreach (double t in new[] { -100.0, -1.0, -0.001, 0.001, 1, 100 })
            Assert(cost < Math.Pow(b + t - 50, 2) + Math.Pow(c + 2 * t - 110, 2));
    }

    private static void Redundant()
    {
        var relations = new[] { Constraint(ConstraintKind.SameX, A, B), Constraint(ConstraintKind.SameX, B, C), Constraint(ConstraintKind.SameX, A, C) };
        var result = Solve(Row(), relations, Shape(C, 77, 120, 40, 40));
        foreach (var shape in result.Values) Near(shape.Left, 77);
    }

    private static void Conflicting()
    {
        var initial = Row();
        var relations = new[] { Constraint(ConstraintKind.SameX, A, B, C), Constraint(ConstraintKind.EqualHorizontalGap, A, B, C) };
        Throws<ConstraintConflictException>(() => Solve(initial, relations, initial[0]));
        Near(initial[0].Left, 0); Near(initial[1].Left, 50); Near(initial[2].Left, 110);
    }

    private static void TranslationStability()
    {
        var initial = Row();
        var relation = Constraint(ConstraintKind.EqualHorizontalGap, A, B, C);
        var normal = Solve(initial, new[] { relation }, Shape(A, 20, 10, 20, 20));
        const double shift = 100000000;
        var shifted = initial.Select(s => Shape(s.Id, s.Left + shift, s.Top - shift, s.Width, s.Height)).ToArray();
        var result = Solve(shifted, new[] { relation }, Shape(A, shift + 20, 10 - shift, 20, 20));
        foreach (var shape in normal.Values)
        {
            Near(Get(result, shape.Id).Left - shift, shape.Left, 0.000001);
            Near(Get(result, shape.Id).Top + shift, shape.Top, 0.000001);
        }
    }

    private static void Repeated()
    {
        var relations = new[] { Constraint(ConstraintKind.SameX, A, B), Constraint(ConstraintKind.SameWidth, A, B) };
        var state = new[] { Shape(A, 12.125), Shape(B, 12.125) }.AsEnumerable();
        for (int i = 0; i < 500; i++)
            state = LayoutSolver.Solve(state, relations, i % 2 == 0 ? A : B, Shape(i % 2 == 0 ? A : B, 12.125, 0, 20, 30));
        foreach (var shape in state) { Near(shape.Left, 12.125); Near(shape.Width, 20); }
    }

    private static void InvalidGeometry()
    {
        Throws<ArgumentOutOfRangeException>(() => Shape(A, Double.NaN));
        Throws<ArgumentOutOfRangeException>(() => Shape(A, 0, Double.PositiveInfinity));
        Throws<ArgumentOutOfRangeException>(() => Shape(A, 0, 0, Double.NegativeInfinity));
        Throws<ArgumentOutOfRangeException>(() => Shape(A, 0, 0, 0));
        Throws<ArgumentOutOfRangeException>(() => Shape(A, 0, 0, 20, 0.5));
        Shape(A, -100, -200, 1, 1);
    }

    private static void InvalidEdit()
    {
        var relation = Constraint(ConstraintKind.SameX, A, B);
        Throws<ArgumentException>(() => LayoutSolver.Solve(new[] { Shape(A), Shape(B) }, new[] { relation }, C, Shape(C)));
        Throws<ArgumentException>(() => LayoutSolver.Solve(new[] { Shape(A), Shape(B) }, new[] { relation }, A, Shape(B)));
        Throws<ArgumentException>(() => LayoutSolver.Solve(new[] { Shape(A) }, new[] { relation }, A, Shape(A)));
        Throws<ArgumentException>(() => LayoutSolver.Solve(new[] { Shape(A), Shape(A.ToUpperInvariant()) }, new LayoutConstraint[0], A, Shape(A)));
    }

    private static void FreeEdit()
    {
        var initial = Row();
        var target = Shape(B, 25, 70, 50, 60);
        var result = Solve(initial, new LayoutConstraint[0], target);
        Same(Get(result, A), initial[0]); Same(Get(result, B), target); Same(Get(result, C), initial[2]);
    }

    private static void Residual()
    {
        var initial = Row();
        foreach (ConstraintKind kind in new[] { ConstraintKind.SameX, ConstraintKind.SameY, ConstraintKind.SameWidth, ConstraintKind.SameHeight })
            Assert(LayoutSolver.MaxResidual(initial, new[] { Constraint(kind, A, B) }) > 0);
        Near(LayoutSolver.MaxResidual(initial, new[] { Constraint(ConstraintKind.EqualHorizontalGap, A, B, C) }), 0);
        Near(LayoutSolver.MaxResidual(new[] { Shape(A, 0, 0, 20, 20), Shape(B, 0, 50, 20, 30), Shape(C, 0, 111, 20, 40) },
            new[] { Constraint(ConstraintKind.EqualVerticalGap, A, B, C) }), 1);
    }

    private static void Tolerance()
    {
        Assert(!LayoutSolver.NeedsUpdate(Shape(A, 123.000001), Shape(A, 123)));
        Assert(LayoutSolver.NeedsUpdate(Shape(A, 123), Shape(A, 123.02)));
        Assert(LayoutSolver.NeedsUpdate(Shape(A), Shape(A, 0, 0, 20.02)));
    }
}
