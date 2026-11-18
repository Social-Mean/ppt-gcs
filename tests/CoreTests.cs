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
        Run("align right bottom and centers without resizing followers", EdgeAndCenterAlignment);
        Run("normalize fixed dimensions and propagate proportional components", FixedDimensions);
        Run("fixed position normalization and driver conflict", FixedPosition);
        Run("aspect ratio side handles preserve opposite edge and orthogonal center", RatioSides);
        Run("aspect ratio corners project pointer dimensions and preserve opposite corner", RatioCorners);
        Run("aspect ratio bounds enforce the minimum on every linked dimension", RatioMinimum);
        Run("mixed ratio and equal-size chains propagate and detect inconsistent cycles", RatioChains);
        Run("fixed signed edge gaps keep unequal sizes and drive either end", FixedGaps);
        Run("ordered two-axis offsets are bidirectional", Offsets);
        Run("vertical and horizontal reflection preserve full bounding-box symmetry", Symmetry);
        Run("disabled missing-object relations do not block unrelated edits", DisabledRelations);
        Run("parameter and member validation covers every new kind", ParameterValidation);
        Run("normalization without a driver minimizes dimensions and positions", NeutralNormalization);
        Run("new relation residuals detect changed constrained geometry", NewResiduals);
        Run("freedom separates indirect position locks from free axes and fixed dimensions", PositionFreedom);
        Run("freedom reports proportional dimensions and the linked minimum", SizeFreedomAndMinimum);
        Run("conflict witness excludes unrelated and compatible relations", ConflictWitness);
        Run("edit restriction explains a rejected direction and a projected corner", EditRestriction);
        Run("page anchors use the current dimensions and preserve both margins", PageAnchors);
        Run("moving and resizing a container preserves its members' insets", ContainerAnchors);
        Run("named dimensions store one value and resolve every reference", NamedParameters);
        Run("component copies remap internal relations and can share or clone parameters", CopyComponent);
        Console.WriteLine("PASS: " + passed + " core tests");
    }

    private static ShapeGeometry Shape(string id, double x = 0, double y = 0, double w = 20, double h = 30)
    { return new ShapeGeometry(id, x, y, w, h); }

    private static void PositionFreedom()
    {
        var shapes=new[]{Shape(A,0,0,20,10),Shape(B,0,30,20,10)};
        var left=Constraint(ConstraintKind.SameX,A,B);
        var pin=Parameter(ConstraintKind.FixedPosition,0,30,B);
        var width=Parameter(ConstraintKind.FixedWidth,20,null,A);
        var freedom=LayoutDiagnostics.AnalyzeShape(shapes,new[]{left,pin,width},A);
        Assert(!freedom.Horizontal && freedom.Vertical && freedom.Width==SizeFreedom.Fixed && freedom.Height==SizeFreedom.IncreaseAndDecrease);
        var unlocked=LayoutDiagnostics.AnalyzeShape(shapes,new[]{left,width},A);
        Assert(unlocked.Horizontal && unlocked.Vertical);
    }

    private static void SizeFreedomAndMinimum()
    {
        var shapes=new[]{Shape(A,0,0,2,1),Shape(B,80,0,2,10)};
        var ratio=Parameter(ConstraintKind.AspectRatio,2,null,A);
        var sameWidth=Constraint(ConstraintKind.SameWidth,A,B);
        var freedom=LayoutDiagnostics.AnalyzeShape(shapes,new[]{ratio,sameWidth},A);
        Assert(freedom.Proportional && freedom.Width==SizeFreedom.Increase && freedom.Height==SizeFreedom.Increase);
        var linked=LayoutDiagnostics.AnalyzeShape(shapes,new[]{ratio,sameWidth},B);
        Assert(linked.Width==SizeFreedom.Increase && linked.Height==SizeFreedom.IncreaseAndDecrease && !linked.Proportional);
    }

    private static void ConflictWitness()
    {
        var shapes=new[]{Shape(A,0,0,20,10),Shape(B,0,0,30,10),Shape(C,60,0,10,10)};
        var a=Parameter(ConstraintKind.FixedWidth,20,null,A);
        var b=Parameter(ConstraintKind.FixedWidth,30,null,B);
        var attempted=Constraint(ConstraintKind.SameWidth,A,B);
        var compatible=Constraint(ConstraintKind.SameY,A,B);
        var unrelated=Parameter(ConstraintKind.FixedWidth,10,null,C);
        var input=new[]{unrelated,a,compatible,attempted,b};
        var witness=LayoutDiagnostics.ExplainConflict(shapes,input,attempted);
        Assert(new HashSet<string>(witness.Select(item=>item.Id)).SetEquals(new[]{a.Id,b.Id,attempted.Id}));
    }

    private static void EditRestriction()
    {
        var shapes=new[]{Shape(A,0,0,20,10),Shape(B,0,30,30,10)};
        var x=Constraint(ConstraintKind.SameX,A,B);
        var pin=Parameter(ConstraintKind.FixedPosition,0,30,B);
        var ratio=Parameter(ConstraintKind.AspectRatio,2,null,A);
        var move=new[]{new GeometryEdit(Shape(A,3,0,20,10),GeometryEditMode.Move)};
        var blocked=LayoutDiagnostics.ExplainEdit(shapes,new[]{x,pin,ratio},move);
        Assert(new HashSet<string>(blocked.Select(item=>item.Id)).SetEquals(new[]{x.Id,pin.Id}));
        var allowed=new[]{new GeometryEdit(Shape(A,0,3,20,10),GeometryEditMode.Move)};
        Assert(LayoutDiagnostics.ExplainEdit(shapes,new[]{x,pin,ratio},allowed,blocked).Count==0);
        var corner=new[]{new GeometryEdit(Shape(A,0,0,26,10),GeometryEditMode.ResizeSE)};
        Assert(LayoutDiagnostics.ExplainEdit(shapes,new[]{x,pin,ratio},corner).SequenceEqual(new[]{ratio}));
    }

    private static LayoutConstraint Constraint(ConstraintKind kind, params string[] shapes)
    {
        serial++;
        return new LayoutConstraint(serial.ToString("x8") + "-0000-0000-0000-000000000001", 256, kind, shapes);
    }

    private static LayoutConstraint Parameter(ConstraintKind kind, double value, double? second, params string[] shapes)
    {
        serial++;
        return new LayoutConstraint(serial.ToString("x8") + "-0000-0000-0000-000000000001", 256, kind, shapes, value, second);
    }

    private static Dictionary<string, ShapeGeometry> Normalize(IEnumerable<ShapeGeometry> input, params LayoutConstraint[] relations)
    { return LayoutSolver.Normalize(input, relations).ToDictionary(s => s.Id); }

    private static Dictionary<string, ShapeGeometry> Edit(IEnumerable<ShapeGeometry> input,
        IEnumerable<LayoutConstraint> relations, ShapeGeometry target, GeometryEditMode mode)
    { return LayoutSolver.Solve(input, relations, target.Id, target, mode).ToDictionary(s => s.Id); }

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
            var all = Enum.GetValues(typeof(ConstraintKind)).Cast<ConstraintKind>().Select(kind => {
                int maximum = LayoutConstraint.GetMaximumShapeCount(kind), parameters = LayoutConstraint.GetParameterCount(kind);
                string[] members = new[] { C, A, B }.Take(Math.Min(maximum, 3)).ToArray();
                return parameters == 0 ? Constraint(kind, members) : Parameter(kind, 12.375, parameters == 2 ? (double?)-8.25 : null, members);
            }).ToArray();
            var source = new ConstraintDocument(all);
            var result = ConstraintDocument.Parse(source.ToXml());
            Assert(result.Constraints.Select(item => item.Kind).SequenceEqual(all.Select(item => item.Kind)));
            Assert(result.Constraints[5].ShapeIds.SequenceEqual(new[] { C, A, B }.Select(ShapeIdentity.Normalize)));
            Assert(source.ToXml() == result.ToXml());
            Assert(ConstraintDocument.Parse(new ConstraintDocument(new LayoutConstraint[0]).ToXml()).Constraints.Count == 0);
        }
        finally { Thread.CurrentThread.CurrentCulture = prior; }
    }

    private static void PageAnchors()
    {
        var input = new[] { Shape(A, 20, 40, 80, 50), Shape(B, 120, 140, 90, 60), Shape(C, 720, 420, 30, 40) };
        var document = new ConstraintDocument(new[] {
            Parameter(ConstraintKind.PageLeft, 20, null, A, B),
            Parameter(ConstraintKind.PageRight, 30, null, A, B),
            Parameter(ConstraintKind.PageCenterY, -10, null, A),
            Parameter(ConstraintKind.PageBottom, 15, null, B)
        });
        string xml = document.ToXml();
        var reloaded = ConstraintDocument.Parse(xml);
        foreach (var page in new[] { new[] { 960.0, 540.0 }, new[] { 1080.0, 610.0 } })
        {
            double width = page[0], height = page[1];
            var relations = reloaded.Constraints.Select(item => item.BindPageSize(width, height)).ToArray();
            var result = Normalize(input, relations);
            Same(Get(result, A), Shape(A, 20, height / 2 - 35, width - 50, 50));
            Same(Get(result, B), Shape(B, 20, height - 75, width - 50, 60));
            Same(Get(result, C), input[2]);
            Near(LayoutSolver.MaxResidual(result.Values, relations), 0);
            var impossible = relations.Concat(new[] { Parameter(ConstraintKind.FixedWidth, 80, null, A) }).ToArray();
            Throws<ConstraintConflictException>(() => Normalize(input, impossible));
        }
        Assert(reloaded.ToXml() == xml);
    }

    private static void ContainerAnchors()
    {
        var input = new[] { Shape(A, 40, 60, 200, 100), Shape(B, 70, 90, 60, 20), Shape(C, 80, 120, 60, 30) };
        var left = Parameter(ConstraintKind.ContainerLeft, 15, null, A, B, C);
        var right = Parameter(ConstraintKind.ContainerRight, 20, null, A, B, C);
        var top = Parameter(ConstraintKind.ContainerTop, 8, null, A, B, C);
        var relations = new[] { left, right, top };
        var initial = Normalize(input, relations);
        Same(Get(initial, A), input[0]);
        Same(Get(initial, B), Shape(B, 55, 68, 165, 20));
        Same(Get(initial, C), Shape(C, 55, 68, 165, 30));
        var moved = Edit(initial.Values, relations, Shape(A, 80, 130, 200, 100), GeometryEditMode.Move);
        Same(Get(moved, B), Shape(B, 95, 138, 165, 20));
        Same(Get(moved, C), Shape(C, 95, 138, 165, 30));
        var resized = Edit(moved.Values, relations, Shape(A, 80, 130, 250, 100), GeometryEditMode.ResizeE);
        Same(Get(resized, B), Shape(B, 95, 138, 215, 20));
        Same(Get(resized, C), Shape(C, 95, 138, 215, 30));
        Near(LayoutSolver.MaxResidual(resized.Values, relations), 0);
        Throws<ConstraintConflictException>(() => Edit(resized.Values, relations, Shape(B, 96, 138, 215, 20), GeometryEditMode.Move));
        Throws<ArgumentException>(() => new ConstraintDocument(new[] { left, Parameter(ConstraintKind.ContainerLeft, 15, null, A, C, B) }));
        Throws<ConstraintConflictException>(() => Normalize(input, left, Parameter(ConstraintKind.ContainerLeft, 15, null, B, A)));
    }

    private static void NamedParameters()
    {
        var parameter = new LayoutParameter(D, "  卡片间距  ", 32.5);
        var width = new LayoutConstraint(A, 256, ConstraintKind.FixedWidth, new[] { A }, 12, parameterId: D);
        var gap = new LayoutConstraint(B, 257, ConstraintKind.FixedHorizontalGap, new[] { B, C }, 1, parameterId: D);
        var doc = new ConstraintDocument(new[] { width, gap }, new[] { parameter });
        var xml = doc.ToXml();
        var reopened = ConstraintDocument.Parse(xml);
        Assert(reopened.ToXml() == xml && reopened.Parameters.Single().Name == "卡片间距");
        Assert(reopened.Constraints.All(item => item.ParameterId == parameter.Id && item.Value == 32.5));
        var updated = new ConstraintDocument(reopened.Constraints, new[] { new LayoutParameter(D, "宽度和间距", 42.75) });
        Assert(updated.Constraints.All(item => item.Value == 42.75));
        Assert(updated.Constraints[0].WithEnabled(false).ParameterId == parameter.Id);
        Near(Get(Normalize(new[] { Shape(A) }, updated.Constraints[0]), A).Width, 42.75);
        Near(LayoutSolver.MaxResidual(new[] { Shape(B), Shape(C, 62.75) }, new[] { updated.Constraints[1] }), 0);
        Throws<ArgumentException>(() => new ConstraintDocument(reopened.Constraints));
        Throws<ArgumentException>(() => new ConstraintDocument(reopened.Constraints, new[] { parameter, new LayoutParameter(E, parameter.Name, 21) }));
        Throws<ArgumentException>(() => new LayoutConstraint(E, 257, ConstraintKind.AspectRatio, new[] { A }, 2, parameterId: D));
        Throws<ArgumentOutOfRangeException>(() => new ConstraintDocument(reopened.Constraints, new[] { new LayoutParameter(D, parameter.Name, 0) }));
        Throws<FormatException>(() => ConstraintDocument.Parse(xml.Replace(" parameterId=", " value=\"1\" parameterId=")));
        Throws<FormatException>(() => ConstraintDocument.Parse(xml.Replace("parameterId=\"" + parameter.Id, "parameterId=\"" + E)));
    }

    private static void CopyComponent()
    {
        var parameter = new LayoutParameter(D, "Step", 30);
        var occupied = new LayoutParameter(E, "Step（副本）", 9);
        var source = new[] {
            new LayoutConstraint(A, 256, ConstraintKind.FixedWidth, new[] { A }, 30, parameterId: D),
            new LayoutConstraint(B, 256, ConstraintKind.FixedHorizontalGap, new[] { B, A }, 30, enabled: false, parameterId: D),
            Constraint(ConstraintKind.SameY, A, C)
        };
        var doc = new ConstraintDocument(source, new[] { parameter, occupied });
        string newA = Guid.NewGuid().ToString("D").ToUpperInvariant(), newB = Guid.NewGuid().ToString("D").ToUpperInvariant();
        var mapping = new Dictionary<string, string> { { A, newA }, { B, newB } };
        foreach (bool share in new[] { false, true })
        {
            var result = ComponentCopy.Plan(doc, 256, 257, mapping, share);
            Assert(result.Constraints.Count == 5 && result.Constraints.Take(3).SequenceEqual(doc.Constraints));
            var copied = result.Constraints.Skip(3).ToArray();
            Assert(copied.All(item => item.SlideId == 257 && !source.Select(old => old.Id).Contains(item.Id)));
            Assert(copied[0].ShapeIds.SequenceEqual(new[] { newA }));
            Assert(copied[1].ShapeIds.SequenceEqual(new[] { newB, newA }) && !copied[1].Enabled);
            Assert(copied[0].ParameterId == copied[1].ParameterId && copied[0].Value == 30);
            if (share) Assert(result.Parameters.Count == 2 && copied[0].ParameterId == parameter.Id);
            else Assert(result.Parameters.Count == 3 && copied[0].ParameterId != parameter.Id && result.Parameters.Last().Name == "Step（副本 2）");
            var updated = new ConstraintDocument(result.Constraints, result.Parameters.Select(item => item.Id == parameter.Id ?
                new LayoutParameter(item.Id, item.Name, 40) : item));
            Assert(updated.Constraints[0].Value == 40 && updated.Constraints[3].Value == (share ? 40 : 30));
            Assert(ConstraintDocument.Parse(result.ToXml()).ToXml() == result.ToXml());
        }
        Throws<ArgumentException>(() => ComponentCopy.Plan(doc, 256, 256, new Dictionary<string, string> { { A, C } }, false));
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

    private static void EdgeAndCenterAlignment()
    {
        var initial = new[] { Shape(A, 10, 20, 40, 30), Shape(B, 100, 80, 70, 50) };
        foreach (var kind in new[] { ConstraintKind.RightAligned, ConstraintKind.BottomAligned,
            ConstraintKind.CenterXAligned, ConstraintKind.CenterYAligned, ConstraintKind.CentersCoincident })
        foreach (var driver in initial)
        {
            var target = Shape(driver.Id, 65, 95, driver.Width + 12, driver.Height + 8);
            var relation = Constraint(kind, A, B);
            var result = Solve(initial, new[] { relation }, target);
            Same(Get(result, driver.Id), target);
            var follower = Get(result, driver.Id == initial[0].Id ? B : A);
            var old = initial.Single(s => s.Id == follower.Id);
            Near(follower.Width, old.Width); Near(follower.Height, old.Height);
            if (kind == ConstraintKind.RightAligned) Near(follower.Left + follower.Width, target.Left + target.Width);
            if (kind == ConstraintKind.BottomAligned) Near(follower.Top + follower.Height, target.Top + target.Height);
            if (kind == ConstraintKind.CenterXAligned || kind == ConstraintKind.CentersCoincident)
                Near(follower.Left + follower.Width / 2, target.Left + target.Width / 2);
            if (kind == ConstraintKind.CenterYAligned || kind == ConstraintKind.CentersCoincident)
                Near(follower.Top + follower.Height / 2, target.Top + target.Height / 2);
            Near(LayoutSolver.MaxResidual(result.Values, new[] { relation }), 0);
        }
    }

    private static void FixedDimensions()
    {
        var initial = new[] { Shape(A, 10, 20, 40, 30), Shape(B, 100, 80, 90, 70) };
        var relations = new[] { Parameter(ConstraintKind.FixedWidth, 80, null, A),
            Parameter(ConstraintKind.FixedHeight, 45, null, A), Constraint(ConstraintKind.SameWidth, A, B),
            Parameter(ConstraintKind.AspectRatio, 2, null, B) };
        var result = Normalize(initial, relations);
        Near(Get(result, A).Width, 80); Near(Get(result, A).Height, 45);
        Near(Get(result, B).Width, 80); Near(Get(result, B).Height, 40);
        foreach (var old in initial) { Near(Get(result, old.Id).Left, old.Left); Near(Get(result, old.Id).Top, old.Top); }
        Near(LayoutSolver.MaxResidual(result.Values, relations), 0);
        Throws<ConstraintConflictException>(() => Edit(result.Values, relations, Shape(A, 10, 20, 90, 45), GeometryEditMode.ResizeE));
    }

    private static void FixedPosition()
    {
        var pin = Parameter(ConstraintKind.FixedPosition, 40, -25, A);
        var initial = new[] { Shape(A, 10, 20), Shape(B, 200, 70) };
        var result = Normalize(initial, pin, Constraint(ConstraintKind.SameX, A, B));
        Near(Get(result, A).Left, 40); Near(Get(result, A).Top, -25); Near(Get(result, B).Left, 40);
        Near(Get(result, B).Top, 70);
        Throws<ConstraintConflictException>(() => Edit(result.Values, new[] { pin }, Shape(A, 41, -25), GeometryEditMode.Move));
        var resized = Edit(result.Values, new[] { pin }, Shape(A, 40, -25, 60, 30), GeometryEditMode.ResizeE);
        Near(Get(resized, A).Width, 60); Near(Get(resized, A).Left, 40);
    }

    private static void RatioSides()
    {
        var initial = new[] { Shape(A, 10, 20, 100, 50), Shape(B, 200, 300, 100, 50) };
        var relations = new[] { Parameter(ConstraintKind.AspectRatio, 2, null, A), Constraint(ConstraintKind.SameWidth, A, B) };
        var targets = new[] { Shape(A, 10, 20, 160, 50), Shape(A, -50, 20, 160, 50),
            Shape(A, 10, 20, 100, 80), Shape(A, 10, -10, 100, 80) };
        var modes = new[] { GeometryEditMode.ResizeE, GeometryEditMode.ResizeW, GeometryEditMode.ResizeS, GeometryEditMode.ResizeN };
        var expected = new[] { Shape(A, 10, 5, 160, 80), Shape(A, -50, 5, 160, 80),
            Shape(A, -20, 20, 160, 80), Shape(A, -20, -10, 160, 80) };
        for (int i = 0; i < modes.Length; i++)
        {
            var result = Edit(initial, relations, targets[i], modes[i]);
            Same(Get(result, A), expected[i]);
            Same(Get(result, B), Shape(B, 200, 300, 160, 50));
            Near(LayoutSolver.MaxResidual(result.Values, relations), 0);
        }
    }

    private static void RatioCorners()
    {
        var initial = new[] { Shape(A, 10, 20, 100, 50) };
        var relation = Parameter(ConstraintKind.AspectRatio, 2, null, A);
        var modes = new[] { GeometryEditMode.ResizeSE, GeometryEditMode.ResizeNE, GeometryEditMode.ResizeSW, GeometryEditMode.ResizeNW };
        foreach (var mode in modes)
        {
            var result = Edit(initial, new[] { relation }, Shape(A, 0, 0, 140, 100), mode);
            var shape = Get(result, A);
            Near(shape.Width, 152); Near(shape.Height, 76);
            Near(shape.Left, mode == GeometryEditMode.ResizeSW || mode == GeometryEditMode.ResizeNW ? -42 : 10);
            Near(shape.Top, mode == GeometryEditMode.ResizeNE || mode == GeometryEditMode.ResizeNW ? -6 : 20);
            Near(LayoutSolver.MaxResidual(result.Values, new[] { relation }), 0);
        }
        // Without a ratio, an ordinary corner still reaches both pointer dimensions.
        var free = Edit(initial, new LayoutConstraint[0], Shape(A, 10, 20, 140, 100), GeometryEditMode.ResizeSE);
        Same(Get(free, A), Shape(A, 10, 20, 140, 100));
    }

    private static void RatioMinimum()
    {
        var relation = Parameter(ConstraintKind.AspectRatio, 10, null, A);
        var result = Normalize(new[] { Shape(A, 0, 0, 1, 1) }, relation);
        Near(Get(result, A).Width, 10); Near(Get(result, A).Height, 1);
        var corner = Edit(result.Values, new[] { relation }, Shape(A, 0, 0, 1, 1), GeometryEditMode.ResizeSE);
        Near(Get(corner, A).Width, 10); Near(Get(corner, A).Height, 1);
        Throws<ConstraintConflictException>(() => Edit(result.Values, new[] { relation }, Shape(A, 0, 0, 1, 1), GeometryEditMode.ResizeE));
    }

    private static void RatioChains()
    {
        var initial = new[] { Shape(A, 0, 0, 100, 50), Shape(B, 200, 0, 100, 25), Shape(C, 400, 0, 50, 25) };
        var relations = new[] { Parameter(ConstraintKind.AspectRatio, 2, null, A),
            Parameter(ConstraintKind.AspectRatio, 4, null, B), Constraint(ConstraintKind.SameWidth, A, B),
            Constraint(ConstraintKind.SameHeight, B, C) };
        var result = Edit(initial, relations, Shape(A, 0, 0, 160, 50), GeometryEditMode.ResizeE);
        Near(Get(result, A).Height, 80); Near(Get(result, B).Width, 160); Near(Get(result, B).Height, 40);
        Near(Get(result, C).Height, 40); Near(Get(result, C).Width, 50);
        Near(LayoutSolver.MaxResidual(result.Values, relations), 0);
        Throws<ConstraintConflictException>(() => Normalize(initial, relations.Concat(new[] { Constraint(ConstraintKind.SameHeight, A, B) }).ToArray()));
    }

    private static void FixedGaps()
    {
        foreach (bool horizontal in new[] { true, false })
        foreach (double distance in new[] { 17.5, -8.0 })
        {
            var relation = Parameter(horizontal ? ConstraintKind.FixedHorizontalGap : ConstraintKind.FixedVerticalGap, distance, null, A, B, C);
            var normalized = Normalize(Row(), relation);
            Near(LayoutSolver.MaxResidual(normalized.Values, new[] { relation }), 0);
            foreach (string driver in new[] { A, C })
            {
                ShapeGeometry old = Get(normalized, driver);
                var target = Shape(driver, old.Left + (horizontal ? 23 : 0), old.Top + (horizontal ? 0 : 23), old.Width, old.Height);
                var result = Edit(normalized.Values, new[] { relation }, target, GeometryEditMode.Move);
                foreach (var shape in normalized.Values)
                {
                    var actual = Get(result, shape.Id);
                    Near(horizontal ? actual.Left : actual.Top, (horizontal ? shape.Left : shape.Top) + 23);
                    Near(actual.Width, shape.Width); Near(actual.Height, shape.Height);
                }
                Near(LayoutSolver.MaxResidual(result.Values, new[] { relation }), 0);
            }
        }
    }

    private static void Offsets()
    {
        var relation = Parameter(ConstraintKind.Offset, 100, 30, A, B);
        var baseline = Normalize(new[] { Shape(A, 10, 10), Shape(B, 120, 40) }, relation);
        Near(Get(baseline, A).Left, 15); Near(Get(baseline, B).Left, 115);
        var fromB = Edit(baseline.Values, new[] { relation }, Shape(B, 150, 70), GeometryEditMode.Move);
        Same(Get(fromB, A), Shape(A, 50, 40));
        var fromA = Edit(baseline.Values, new[] { relation }, Shape(A, -30, -20), GeometryEditMode.Move);
        Same(Get(fromA, B), Shape(B, 70, 10));
        Near(LayoutSolver.MaxResidual(fromA.Values, new[] { relation }), 0);
    }

    private static void Symmetry()
    {
        foreach (bool vertical in new[] { true, false })
        {
            var relation = Parameter(vertical ? ConstraintKind.SymmetricVertical : ConstraintKind.SymmetricHorizontal, 150, null, A, B);
            var baseline = Normalize(new[] { Shape(A, 10, 20, 40, 20), Shape(B, 220, 200, 60, 40) }, relation);
            Near(Get(baseline, A).Width, 50); Near(Get(baseline, A).Height, 30);
            Near(LayoutSolver.MaxResidual(baseline.Values, new[] { relation }), 0);
            foreach (string driver in new[] { A, B })
            {
                var target = Shape(driver, 35, 55, 70, 40);
                var result = Solve(baseline.Values, new[] { relation }, target);
                Same(Get(result, driver), target);
                var follower = Get(result, driver == A ? B : A);
                Same(follower, Shape(follower.Id, vertical ? 195 : 35, vertical ? 55 : 205, 70, 40));
                Near(LayoutSolver.MaxResidual(result.Values, new[] { relation }), 0);
            }
        }
    }

    private static void DisabledRelations()
    {
        var source = Constraint(ConstraintKind.SameX, A, B);
        var disabled = new LayoutConstraint(source.Id, source.SlideId, source.Kind, source.ShapeIds, enabled: false);
        var state = new[] { Shape(A) };
        Same(Get(Edit(state, new[] { disabled }, Shape(A, 40), GeometryEditMode.Move), A), Shape(A, 40));
        Same(Get(Normalize(state, disabled), A), state[0]);
        Near(LayoutSolver.MaxResidual(state, new[] { disabled }), 0);
        Throws<ArgumentException>(() => Normalize(state, source));
        var document = new ConstraintDocument(new[] { disabled });
        Assert(!ConstraintDocument.Parse(document.ToXml()).Constraints[0].Enabled);
        Assert(document.ToXml() == ConstraintDocument.Parse(document.ToXml()).ToXml());
    }

    private static void ParameterValidation()
    {
        Throws<ArgumentException>(() => Constraint(ConstraintKind.FixedWidth, A));
        Throws<ArgumentException>(() => Parameter(ConstraintKind.FixedWidth, 20, null, A, B));
        Throws<ArgumentException>(() => Parameter(ConstraintKind.Offset, 20, 30, A, B, C));
        Throws<ArgumentException>(() => Parameter(ConstraintKind.SymmetricVertical, 10, null, A));
        Throws<ArgumentException>(() => Parameter(ConstraintKind.SameWidth, 20, null, A, B));
        Throws<ArgumentException>(() => Parameter(ConstraintKind.FixedPosition, 20, null, A));
        Throws<ArgumentOutOfRangeException>(() => Parameter(ConstraintKind.FixedHeight, 0.5, null, A));
        Throws<ArgumentOutOfRangeException>(() => Parameter(ConstraintKind.AspectRatio, 0, null, A));
        Throws<ArgumentOutOfRangeException>(() => Parameter(ConstraintKind.AspectRatio, Double.PositiveInfinity, null, A));
        string xml = new ConstraintDocument(new[] { Parameter(ConstraintKind.FixedWidth, 20.5, null, A) }).ToXml();
        foreach (string value in new[] { "NaN", "Infinity", "1,5", "", "0" })
            Throws<FormatException>(() => ConstraintDocument.Parse(xml.Replace("value=\"20.5\"", "value=\"" + value + "\"")));
        Throws<FormatException>(() => ConstraintDocument.Parse(xml.Replace("<constraint ", "<constraint enabled=\"False\" ")));
        Throws<FormatException>(() => ConstraintDocument.Parse(xml.Replace("<constraint ", "<constraint secondaryValue=\"5\" ")));
        var offset = Parameter(ConstraintKind.Offset, 20, 30, A, B);
        Assert(new ConstraintDocument(new[] { offset, Parameter(ConstraintKind.Offset, -20, -30, B, A) }).Constraints.Count == 2);
    }

    private static void NeutralNormalization()
    {
        var initial = new[] { Shape(A, 0, 20, 20, 30), Shape(B, 100, 80, 40, 50), Shape(C, 300, 400, 80, 90) };
        var result = Normalize(initial, Constraint(ConstraintKind.SameX, A, B), Constraint(ConstraintKind.SameWidth, A, B));
        Near(Get(result, A).Left, 50); Near(Get(result, B).Left, 50);
        Near(Get(result, A).Width, 30); Near(Get(result, B).Width, 30);
        Near(Get(result, A).Top, 20); Near(Get(result, B).Top, 80);
        Same(Get(result, C), initial[2]);
        Assert(LayoutSolver.Normalize(new ShapeGeometry[0], new LayoutConstraint[0]).Count == 0);
        var fixedOne = Parameter(ConstraintKind.FixedWidth, 20, null, A);
        var fixedTwo = Parameter(ConstraintKind.FixedWidth, 30, null, A);
        Throws<ConstraintConflictException>(() => Normalize(initial, fixedOne, fixedTwo));
    }

    private static void NewResiduals()
    {
        var initial = new[] { Shape(A, 10, 20, 40, 30), Shape(B, 100, 80, 70, 50) };
        var relations = new[] { Constraint(ConstraintKind.RightAligned, A, B), Constraint(ConstraintKind.BottomAligned, A, B),
            Constraint(ConstraintKind.CenterXAligned, A, B), Constraint(ConstraintKind.CenterYAligned, A, B),
            Constraint(ConstraintKind.CentersCoincident, A, B), Parameter(ConstraintKind.FixedPosition, 40, 60, A),
            Parameter(ConstraintKind.FixedWidth, 80, null, A), Parameter(ConstraintKind.FixedHeight, 60, null, A),
            Parameter(ConstraintKind.AspectRatio, 2, null, A), Parameter(ConstraintKind.FixedHorizontalGap, 20, null, A, B),
            Parameter(ConstraintKind.FixedVerticalGap, 20, null, A, B), Parameter(ConstraintKind.Offset, 120, 60, A, B),
            Parameter(ConstraintKind.SymmetricVertical, 200, null, A, B), Parameter(ConstraintKind.SymmetricHorizontal, 200, null, A, B) };
        foreach (var relation in relations)
        {
            var result = Normalize(initial, relation);
            Near(LayoutSolver.MaxResidual(result.Values, new[] { relation }), 0);
            var altered = result.Values.Select(s => s.Id == ShapeIdentity.Normalize(A)
                ? Shape(s.Id, s.Left + 2, s.Top + 3, s.Width + 5, s.Height + 7) : s).ToArray();
            Assert(LayoutSolver.MaxResidual(altered, new[] { relation }) > 0.1);
        }
    }
}
