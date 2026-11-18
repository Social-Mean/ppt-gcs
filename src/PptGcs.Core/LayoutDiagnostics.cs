using System;
using System.Collections.Generic;
using System.Linq;

namespace PptGcs.Core
{
    public enum SizeFreedom { Fixed, Increase, IncreaseAndDecrease }

    public sealed class GeometryFreedom
    {
        public bool Horizontal { get; }
        public bool Vertical { get; }
        public SizeFreedom Width { get; }
        public SizeFreedom Height { get; }
        public bool Proportional { get; }

        internal GeometryFreedom(bool horizontal, bool vertical, SizeFreedom width, SizeFreedom height, bool proportional)
        {
            Horizontal = horizontal;
            Vertical = vertical;
            Width = width;
            Height = height;
            Proportional = proportional;
        }
    }

    public static class LayoutDiagnostics
    {
        public static GeometryFreedom AnalyzeShape(IReadOnlyList<ShapeGeometry> geometry, IReadOnlyList<LayoutConstraint> relations, string shapeId)
        {
            string id = ShapeIdentity.Normalize(shapeId);
            var scope = LayoutSolver.ConnectedConstraints(relations, new[] { id });
            var ids = new HashSet<string>(scope.SelectMany(relation => relation.ShapeIds), StringComparer.Ordinal) { id };
            var component = geometry.Where(shape => ids.Contains(shape.Id)).ToList();
            if (scope.Count == 0)
            {
                var free = component.Single();
                return new GeometryFreedom(true, true, FreeSize(free.Width), FreeSize(free.Height), false);
            }
            var baseline = LayoutSolver.Normalize(component, scope);
            if (baseline.Zip(component, LayoutSolver.NeedsUpdate).Any(changed => changed)) return null;
            var driver = baseline.Single(shape => shape.Id == id);
            // Position equations are linear at fixed dimensions. Size edits use
            // the native left/top anchor and can move proportional followers.
            bool horizontal = TrySolve(baseline, scope, new[] { Edit(driver, 1, 0, 0, 0, GeometryEditMode.Move) }) != null;
            bool vertical = TrySolve(baseline, scope, new[] { Edit(driver, 0, 1, 0, 0, GeometryEditMode.Move) }) != null;
            var wider = TrySolve(baseline, scope, new[] { Edit(driver, 0, 0, 1, 0, GeometryEditMode.NativeWidth) });
            var taller = TrySolve(baseline, scope, new[] { Edit(driver, 0, 0, 0, 1, GeometryEditMode.NativeHeight) });
            bool proportional = wider != null && Math.Abs(wider.Single(shape => shape.Id == id).Height - driver.Height) > LayoutSolver.ResidualTolerance ||
                taller != null && Math.Abs(taller.Single(shape => shape.Id == id).Width - driver.Width) > LayoutSolver.ResidualTolerance;
            return new GeometryFreedom(horizontal, vertical, SizeRange(baseline, wider), SizeRange(baseline, taller), proportional);
        }

        private static GeometryEdit Edit(ShapeGeometry shape, double dx, double dy, double dw, double dh, GeometryEditMode mode)
        { return new GeometryEdit(new ShapeGeometry(shape.Id, shape.Left + dx, shape.Top + dy, shape.Width + dw, shape.Height + dh), mode); }

        private static SizeFreedom SizeRange(IReadOnlyList<ShapeGeometry> baseline, IReadOnlyList<ShapeGeometry> grown)
        {
            if (grown == null) return SizeFreedom.Fixed;
            // A feasible unit increase identifies the proportional component.
            // Shrinking reverses that vector until any linked size hits its minimum.
            for (int i = 0; i < baseline.Count; i++)
            {
                var before = baseline[i];
                var after = grown[i];
                if (AtMinimum(before.Width, after.Width) || AtMinimum(before.Height, after.Height)) return SizeFreedom.Increase;
            }
            return SizeFreedom.IncreaseAndDecrease;
        }

        private static bool AtMinimum(double before, double after)
        { return after - before > LayoutSolver.ResidualTolerance && before <= LayoutSolver.MinimumSize + LayoutSolver.ResidualTolerance; }

        private static SizeFreedom FreeSize(double value)
        { return value <= LayoutSolver.MinimumSize + LayoutSolver.ResidualTolerance ? SizeFreedom.Increase : SizeFreedom.IncreaseAndDecrease; }

        public static IReadOnlyList<LayoutConstraint> ExplainConflict(IReadOnlyList<ShapeGeometry> geometry,
            IReadOnlyList<LayoutConstraint> relations, LayoutConstraint attempted = null)
        {
            Func<IReadOnlyList<ShapeGeometry>, IReadOnlyList<LayoutConstraint>, bool> conflicts = (shapes, set) => TrySolve(shapes, set, null) == null;
            if (attempted != null)
            {
                var scope = LayoutSolver.ConnectedConstraints(relations, attempted.ShapeIds);
                var ids = new HashSet<string>(scope.SelectMany(relation => relation.ShapeIds), StringComparer.Ordinal);
                var shapes = geometry.Where(shape => ids.Contains(shape.Id)).ToList();
                if (conflicts(shapes, scope)) return Minimize(scope, set => conflicts(shapes, set));
            }
            return Minimize(relations.Where(relation => relation.Enabled).ToList(), set => conflicts(geometry, set));
        }

        public static IReadOnlyList<LayoutConstraint> ExplainEdit(IReadOnlyList<ShapeGeometry> geometry,
            IReadOnlyList<LayoutConstraint> relations, IReadOnlyList<GeometryEdit> edits, IReadOnlyList<LayoutConstraint> previous = null)
        {
            var scope = LayoutSolver.ConnectedConstraints(relations, edits.Select(edit => edit.Id));
            var ids = new HashSet<string>(scope.SelectMany(relation => relation.ShapeIds).Concat(edits.Select(edit => edit.Id)), StringComparer.Ordinal);
            var shapes = geometry.Where(shape => ids.Contains(shape.Id)).ToList();
            bool Restricted(IReadOnlyList<LayoutConstraint> set)
            {
                var result = TrySolve(shapes, set, edits);
                return result == null || edits.Any(edit => LayoutSolver.NeedsUpdate(edit.Target, result.Single(shape => shape.Id == edit.Id)));
            }
            if (previous != null && previous.Count > 0 && Restricted(previous)) return previous;
            return Minimize(scope, Restricted);
        }

        private static IReadOnlyList<LayoutConstraint> Minimize(IReadOnlyList<LayoutConstraint> startingSet,
            Func<IReadOnlyList<LayoutConstraint>, bool> fails)
        {
            if (!fails(startingSet) || fails(Array.Empty<LayoutConstraint>())) return Array.Empty<LayoutConstraint>();
            var witness = startingSet.ToList();
            // The dimension stage and position stage are deliberately separate.
            // Recheck earlier deletions if a smaller set used a different projection.
            bool reduced;
            do
            {
                reduced = false;
                for (int i = witness.Count - 1; i >= 0; i--)
                {
                    var candidate = witness.Where((relation, index) => index != i).ToList();
                    if (!fails(candidate)) continue;
                    witness = candidate;
                    reduced = true;
                }
            } while (reduced);
            return witness.AsReadOnly();
        }

        private static IReadOnlyList<ShapeGeometry> TrySolve(IReadOnlyList<ShapeGeometry> geometry,
            IReadOnlyList<LayoutConstraint> relations, IReadOnlyList<GeometryEdit> edits)
        {
            try { return edits == null ? LayoutSolver.Normalize(geometry, relations) : LayoutSolver.Solve(geometry, relations, edits); }
            catch (ConstraintConflictException) { return null; }
        }
    }
}
