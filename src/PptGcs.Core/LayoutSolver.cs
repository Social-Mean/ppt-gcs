using System;
using System.Collections.Generic;
using System.Linq;
using MathNet.Numerics.LinearAlgebra;

namespace PptGcs.Core
{
    public sealed class ConstraintConflictException : InvalidOperationException
    {
        public ConstraintConflictException(string message) : base(message) { }
    }

    // Equal dimensions propagate through their equality groups. With sizes known,
    // position constraints are linear. Orthogonal projection preserves all unused
    // degrees of freedom and minimizes the sum of squared position changes.
    public static class LayoutSolver
    {
        public const double PositionTolerance = 0.01;
        public const double MinimumSize = 1.0;
        public const double ResidualTolerance = 0.00001;

        public static IReadOnlyList<ShapeGeometry> Solve(IEnumerable<ShapeGeometry> initial,
            IEnumerable<LayoutConstraint> constraints, string driverId, ShapeGeometry target)
        {
            if (initial == null) throw new ArgumentNullException("initial");
            if (constraints == null) throw new ArgumentNullException("constraints");
            if (target == null) throw new ArgumentNullException("target");
            var baseline = initial.ToList();
            var byId = Index(baseline);
            var relations = constraints.ToList();
            ValidateReferences(relations, byId);
            string driver = ShapeIdentity.Normalize(driverId);
            if (!byId.ContainsKey(driver)) throw new ArgumentException("The driver is not in the geometry snapshot.", "driverId");
            if (target.Id != driver) throw new ArgumentException("Target and driver identities differ.", "target");

            var connected = new HashSet<string>(StringComparer.Ordinal) { driver };
            bool expanded;
            do
            {
                expanded = false;
                foreach (var relation in relations)
                    if (relation.ShapeIds.Any(connected.Contains))
                        foreach (var id in relation.ShapeIds) expanded |= connected.Add(id);
            } while (expanded);
            var component = baseline.Where(shape => connected.Contains(shape.Id)).ToList();
            var active = relations.Where(relation => connected.Contains(relation.ShapeIds[0])).ToList();
            var indices = component.Select((shape, index) => new { shape.Id, Index = index })
                .ToDictionary(item => item.Id, item => item.Index, StringComparer.Ordinal);
            var widths = ResolveSizes(component, active, driver, target.Width, ConstraintKind.SameWidth, shape => shape.Width);
            var heights = ResolveSizes(component, active, driver, target.Height, ConstraintKind.SameHeight, shape => shape.Height);
            var xs = ResolvePositions(component, active, indices, widths, driver, target.Left, true);
            var ys = ResolvePositions(component, active, indices, heights, driver, target.Top, false);
            var solved = new Dictionary<string, ShapeGeometry>(StringComparer.Ordinal);
            for (int i = 0; i < component.Count; i++)
                solved.Add(component[i].Id, new ShapeGeometry(component[i].Id, xs[i], ys[i], widths[i], heights[i]));
            if (MaxResidual(solved.Values, active) > ResidualTolerance)
                throw new ConstraintConflictException("The requested edit cannot satisfy the layout constraints.");
            return baseline.Select(shape => solved.ContainsKey(shape.Id) ? solved[shape.Id] : shape).ToList().AsReadOnly();
        }

        public static bool NeedsUpdate(ShapeGeometry current, ShapeGeometry target)
        {
            if (current == null) throw new ArgumentNullException("current");
            if (target == null) throw new ArgumentNullException("target");
            if (current.Id != target.Id) throw new ArgumentException("Geometry identities differ.");
            return Math.Abs(current.Left - target.Left) > PositionTolerance ||
                Math.Abs(current.Top - target.Top) > PositionTolerance ||
                Math.Abs(current.Width - target.Width) > PositionTolerance ||
                Math.Abs(current.Height - target.Height) > PositionTolerance;
        }

        public static double MaxResidual(IEnumerable<ShapeGeometry> geometries, IEnumerable<LayoutConstraint> constraints)
        {
            if (geometries == null) throw new ArgumentNullException("geometries");
            if (constraints == null) throw new ArgumentNullException("constraints");
            var byId = Index(geometries);
            var relations = constraints.ToList();
            ValidateReferences(relations, byId);
            double largest = 0;
            foreach (var relation in relations)
            {
                var shapes = relation.ShapeIds.Select(id => byId[id]).ToList();
                if (relation.Kind == ConstraintKind.EqualHorizontalGap || relation.Kind == ConstraintKind.EqualVerticalGap)
                {
                    bool horizontal = relation.Kind == ConstraintKind.EqualHorizontalGap;
                    double first = Gap(shapes[0], shapes[1], horizontal);
                    for (int i = 2; i < shapes.Count; i++)
                        largest = Math.Max(largest, Math.Abs(Gap(shapes[i - 1], shapes[i], horizontal) - first));
                }
                else
                {
                    double first = Attribute(shapes[0], relation.Kind);
                    for (int i = 1; i < shapes.Count; i++)
                        largest = Math.Max(largest, Math.Abs(Attribute(shapes[i], relation.Kind) - first));
                }
            }
            return largest;
        }

        private static Dictionary<string, ShapeGeometry> Index(IEnumerable<ShapeGeometry> shapes)
        {
            var result = new Dictionary<string, ShapeGeometry>(StringComparer.Ordinal);
            foreach (var shape in shapes)
            {
                if (shape == null) throw new ArgumentException("A shape cannot be null.");
                if (result.ContainsKey(shape.Id)) throw new ArgumentException("Duplicate geometry identity: " + shape.Id);
                result.Add(shape.Id, shape);
            }
            return result;
        }

        private static void ValidateReferences(IEnumerable<LayoutConstraint> relations, IDictionary<string, ShapeGeometry> shapes)
        {
            foreach (var relation in relations)
            {
                if (relation == null) throw new ArgumentException("A constraint cannot be null.");
                foreach (var id in relation.ShapeIds)
                    if (!shapes.ContainsKey(id)) throw new ArgumentException("A constraint references a missing shape: " + id);
            }
        }

        private static double[] ResolveSizes(IList<ShapeGeometry> component, IList<LayoutConstraint> relations,
            string driver, double requestedSize, ConstraintKind kind, Func<ShapeGeometry, double> size)
        {
            var groups = new DisjointSets(component.Select(shape => shape.Id));
            foreach (var relation in relations.Where(item => item.Kind == kind))
                for (int i = 1; i < relation.ShapeIds.Count; i++) groups.Join(relation.ShapeIds[0], relation.ShapeIds[i]);
            var chosen = new Dictionary<string, double>(StringComparer.Ordinal);
            foreach (var shape in component)
            {
                string root = groups.Find(shape.Id);
                if (!chosen.ContainsKey(root)) chosen.Add(root, size(shape));
            }
            chosen[groups.Find(driver)] = requestedSize;
            return component.Select(shape => chosen[groups.Find(shape.Id)]).ToArray();
        }

        private static double[] ResolvePositions(IList<ShapeGeometry> component, IList<LayoutConstraint> relations,
            IDictionary<string, int> indices, double[] sizes, string driver, double requestedPosition, bool horizontal)
        {
            int count = component.Count;
            var rows = new List<double[]>();
            var values = new List<double>();
            ConstraintKind alignment = horizontal ? ConstraintKind.SameX : ConstraintKind.SameY;
            ConstraintKind spacing = horizontal ? ConstraintKind.EqualHorizontalGap : ConstraintKind.EqualVerticalGap;
            foreach (var relation in relations)
            {
                if (relation.Kind == alignment)
                {
                    int first = indices[relation.ShapeIds[0]];
                    for (int i = 1; i < relation.ShapeIds.Count; i++)
                    {
                        var row = new double[count]; row[indices[relation.ShapeIds[i]]] = 1; row[first] = -1;
                        rows.Add(row); values.Add(0);
                    }
                }
                else if (relation.Kind == spacing)
                {
                    // p[i+1] - p[i] - size[i] = p[1] - p[0] - size[0].
                    int first = indices[relation.ShapeIds[0]], second = indices[relation.ShapeIds[1]];
                    for (int i = 1; i < relation.ShapeIds.Count - 1; i++)
                    {
                        int left = indices[relation.ShapeIds[i]], right = indices[relation.ShapeIds[i + 1]];
                        var row = new double[count];
                        row[right] += 1; row[left] -= 1; row[second] -= 1; row[first] += 1;
                        rows.Add(row); values.Add(sizes[left] - sizes[first]);
                    }
                }
            }

            // Eliminate the driven coordinate rather than making it a soft target.
            // The remaining solution is the least movement needed for this edit.
            int driverIndex = indices[driver];
            var baseline = component.Select(shape => horizontal ? shape.Left : shape.Top).ToArray();
            baseline[driverIndex] = requestedPosition;
            if (rows.Count == 0) return baseline;
            var free = Enumerable.Range(0, count).Where(index => index != driverIndex).ToArray();
            var a = Matrix<double>.Build.Dense(rows.Count, free.Length, (row, column) => rows[row][free[column]]);
            var rhs = Vector<double>.Build.Dense(rows.Count, row => values[row] - rows[row][driverIndex] * requestedPosition);
            var start = Vector<double>.Build.Dense(free.Length, column => baseline[free[column]]);
            Vector<double> result = start + a.PseudoInverse() * (rhs - a * start);
            if ((a * result - rhs).AbsoluteMaximum() > ResidualTolerance)
                throw new ConstraintConflictException("The requested edit conflicts with the " + (horizontal ? "horizontal" : "vertical") + " constraints.");
            for (int i = 0; i < free.Length; i++) baseline[free[i]] = result[i];
            return baseline;
        }

        private static double Attribute(ShapeGeometry shape, ConstraintKind kind)
        {
            switch (kind)
            {
                case ConstraintKind.SameX: return shape.Left;
                case ConstraintKind.SameY: return shape.Top;
                case ConstraintKind.SameWidth: return shape.Width;
                case ConstraintKind.SameHeight: return shape.Height;
                default: throw new ArgumentOutOfRangeException("kind");
            }
        }

        private static double Gap(ShapeGeometry left, ShapeGeometry right, bool horizontal)
        {
            return horizontal ? right.Left - left.Left - left.Width : right.Top - left.Top - left.Height;
        }

        private sealed class DisjointSets
        {
            private readonly Dictionary<string, string> parent;
            public DisjointSets(IEnumerable<string> ids) { parent = ids.ToDictionary(id => id, id => id, StringComparer.Ordinal); }
            public string Find(string id)
            {
                if (parent[id] != id) parent[id] = Find(parent[id]);
                return parent[id];
            }
            public void Join(string a, string b) { parent[Find(b)] = Find(a); }
        }
    }
}
