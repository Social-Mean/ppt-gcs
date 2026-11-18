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

    public enum GeometryEditMode { Exact, Move, ResizeNW, ResizeN, ResizeNE, ResizeE, ResizeSE, ResizeS, ResizeSW, ResizeW, NativeWidth, NativeHeight }

    // Dimensions are proportional equality components (same-size, ratio, fixed
    // values). Positions then form linear equations, solved by least movement.
    // Keeping these stages separate prevents alignment from resizing a follower.
    public static class LayoutSolver
    {
        public const double PositionTolerance = 0.01;
        public const double MinimumSize = 1.0;
        public const double ResidualTolerance = 0.00001;

        public static IReadOnlyList<ShapeGeometry> Normalize(IEnumerable<ShapeGeometry> initial,
            IEnumerable<LayoutConstraint> constraints)
        {
            return SolveCore(initial, constraints, new GeometryEdit[0]);
        }

        public static IReadOnlyList<ShapeGeometry> Solve(IEnumerable<ShapeGeometry> initial,
            IEnumerable<LayoutConstraint> constraints, string driverId, ShapeGeometry target,
            GeometryEditMode mode = GeometryEditMode.Exact)
        {
            if (target == null) throw new ArgumentNullException("target");
            string driver = ShapeIdentity.Normalize(driverId);
            if (target.Id != driver) throw new ArgumentException("Target and driver identities differ.", "target");
            return Solve(initial, constraints, new[] { new GeometryEdit(target, mode) });
        }

        public static IReadOnlyList<ShapeGeometry> Solve(IEnumerable<ShapeGeometry> initial,
            IEnumerable<LayoutConstraint> constraints, IEnumerable<GeometryEdit> edits)
        {
            if (edits == null) throw new ArgumentNullException("edits");
            var requested = edits.ToList();
            if (requested.Count == 0) throw new ArgumentException("An edit must contain at least one driver.", "edits");
            if (requested.Any(edit => edit == null)) throw new ArgumentException("An edit cannot be null.", "edits");
            if (requested.Select(edit => edit.Id).Distinct(StringComparer.Ordinal).Count() != requested.Count)
                throw new ArgumentException("A shape cannot drive the same edit more than once.", "edits");
            return SolveCore(initial, constraints, requested);
        }

        public static IReadOnlyList<LayoutConstraint> ConnectedConstraints(IEnumerable<LayoutConstraint> constraints,
            IEnumerable<string> driverIds)
        {
            if(constraints==null)throw new ArgumentNullException("constraints");
            if(driverIds==null)throw new ArgumentNullException("driverIds");
            var relations=Active(constraints);
            var connected=new HashSet<string>(driverIds.Select(ShapeIdentity.Normalize),StringComparer.Ordinal);
            bool expanded;
            do
            {
                expanded=false;
                foreach(var relation in relations)
                    if(relation.ShapeIds.Any(connected.Contains))
                        foreach(string id in relation.ShapeIds)expanded|=connected.Add(id);
            }while(expanded);
            return relations.Where(relation=>relation.ShapeIds.Any(connected.Contains)).ToList().AsReadOnly();
        }

        private static IReadOnlyList<ShapeGeometry> SolveCore(IEnumerable<ShapeGeometry> initial,
            IEnumerable<LayoutConstraint> constraints, IList<GeometryEdit> edits)
        {
            if (initial == null) throw new ArgumentNullException("initial");
            if (constraints == null) throw new ArgumentNullException("constraints");
            var baseline = initial.ToList();
            var byId = Index(baseline);
            var relations = Active(constraints);
            ValidateReferences(relations, byId);
            if (edits.Any(edit => !byId.ContainsKey(edit.Id)))
                throw new ArgumentException("A driver is not in the geometry snapshot.", "edits");
            var seeds=edits.Count==0 ? byId.Keys : edits.Select(edit=>edit.Id);
            var active=ConnectedConstraints(relations,seeds).ToList();
            var connected=new HashSet<string>(seeds,StringComparer.Ordinal);
            foreach(var relation in active)foreach(string id in relation.ShapeIds)connected.Add(id);
            var component = baseline.Where(shape => connected.Contains(shape.Id)).ToList();
            var indices = component.Select((shape, index) => new { shape.Id, Index = index })
                .ToDictionary(item => item.Id, item => item.Index, StringComparer.Ordinal);
            double[] dimensions = ResolveDimensions(component, active, indices, edits);
            double[] widths = Enumerable.Range(0, component.Count).Select(i => dimensions[2 * i]).ToArray();
            double[] heights = Enumerable.Range(0, component.Count).Select(i => dimensions[2 * i + 1]).ToArray();
            var xsRequested = new Dictionary<string, double>(StringComparer.Ordinal);
            var ysRequested = new Dictionary<string, double>(StringComparer.Ordinal);
            foreach (var edit in edits)
            {
                int index = indices[edit.Id];
                var anchor = DrivenPosition(component[index], edit.Target, widths[index], heights[index], edit.Mode);
                xsRequested.Add(edit.Id, anchor[0]); ysRequested.Add(edit.Id, anchor[1]);
            }
            var xs = ResolvePositions(component, active, indices, widths, xsRequested, true);
            var ys = ResolvePositions(component, active, indices, heights, ysRequested, false);
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
            var relations = Active(constraints);
            ValidateReferences(relations, byId);
            double largest = 0;
            foreach (var relation in relations)
            {
                var shapes = relation.ShapeIds.Select(id => byId[id]).ToList();
                ShapeGeometry first = shapes[0];
                switch (relation.Kind)
                {
                    case ConstraintKind.EqualHorizontalGap:
                    case ConstraintKind.EqualVerticalGap:
                    case ConstraintKind.FixedHorizontalGap:
                    case ConstraintKind.FixedVerticalGap:
                        bool horizontal = relation.Kind == ConstraintKind.EqualHorizontalGap || relation.Kind == ConstraintKind.FixedHorizontalGap;
                        double gap = relation.Value ?? Gap(first, shapes[1], horizontal);
                        for (int i = 1; i < shapes.Count; i++) largest = Max(largest, Gap(shapes[i - 1], shapes[i], horizontal) - gap);
                        break;
                    case ConstraintKind.FixedWidth: largest = Max(largest, first.Width - relation.Value.Value); break;
                    case ConstraintKind.FixedHeight: largest = Max(largest, first.Height - relation.Value.Value); break;
                    case ConstraintKind.FixedPosition:
                        largest = Max(largest, first.Left - relation.Value.Value, first.Top - relation.SecondaryValue.Value); break;
                    case ConstraintKind.AspectRatio:
                        largest = Max(largest, first.Width - relation.Value.Value * first.Height); break;
                    case ConstraintKind.Offset:
                        largest = Max(largest, shapes[1].Left - first.Left - relation.Value.Value,
                            shapes[1].Top - first.Top - relation.SecondaryValue.Value); break;
                    case ConstraintKind.SymmetricVertical:
                    case ConstraintKind.SymmetricHorizontal:
                        ShapeGeometry second = shapes[1];
                        largest = Max(largest, first.Width - second.Width, first.Height - second.Height);
                        largest = relation.Kind == ConstraintKind.SymmetricVertical
                            ? Max(largest, first.Top - second.Top, first.Left + first.Width / 2 + second.Left + second.Width / 2 - 2 * relation.Value.Value)
                            : Max(largest, first.Left - second.Left, first.Top + first.Height / 2 + second.Top + second.Height / 2 - 2 * relation.Value.Value);
                        break;
                    case ConstraintKind.CentersCoincident:
                        foreach (var shape in shapes.Skip(1))
                            largest = Max(largest, shape.Left + shape.Width / 2 - first.Left - first.Width / 2,
                                shape.Top + shape.Height / 2 - first.Top - first.Height / 2);
                        break;
                    default:
                        foreach (var shape in shapes.Skip(1)) largest = Max(largest, Attribute(shape, relation.Kind) - Attribute(first, relation.Kind));
                        break;
                }
            }
            return largest;
        }

        private static double Max(double initial, params double[] differences)
        { return differences.Aggregate(initial, (largest, difference) => Math.Max(largest, Math.Abs(difference))); }

        private static List<LayoutConstraint> Active(IEnumerable<LayoutConstraint> constraints)
        {
            var result = new List<LayoutConstraint>();
            foreach (var relation in constraints)
            {
                if (relation == null) throw new ArgumentException("A constraint cannot be null.");
                if (relation.Enabled) result.Add(relation);
            }
            return result;
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
                foreach (var id in relation.ShapeIds)
                    if (!shapes.ContainsKey(id)) throw new ArgumentException("A constraint references a missing shape: " + id);
        }

        private static double[] ResolveDimensions(IList<ShapeGeometry> shapes, IList<LayoutConstraint> relations,
            IDictionary<string, int> indices, IList<GeometryEdit> edits)
        {
            int count = shapes.Count * 2;
            var edges = Enumerable.Range(0, count).Select(i => (IList<SizeEdge>)new List<SizeEdge>()).ToArray();
            var fixedValues = new List<SizeValue>();
            var preferred = new List<SizeValue>();
            foreach (var relation in relations)
            {
                int first = indices[relation.ShapeIds[0]] * 2;
                bool symmetric = relation.Kind == ConstraintKind.SymmetricVertical || relation.Kind == ConstraintKind.SymmetricHorizontal;
                if (relation.Kind == ConstraintKind.SameWidth || relation.Kind == ConstraintKind.SameHeight || symmetric)
                    for (int i = 1; i < relation.ShapeIds.Count; i++)
                    {
                        int second = indices[relation.ShapeIds[i]] * 2;
                        if (relation.Kind != ConstraintKind.SameHeight) JoinSize(edges, first, second, 1);
                        if (relation.Kind != ConstraintKind.SameWidth) JoinSize(edges, first + 1, second + 1, 1);
                    }
                else if (relation.Kind == ConstraintKind.AspectRatio) JoinSize(edges, first + 1, first, relation.Value.Value);
                else if (relation.Kind == ConstraintKind.FixedWidth) fixedValues.Add(new SizeValue(first, relation.Value.Value));
                else if (relation.Kind == ConstraintKind.FixedHeight) fixedValues.Add(new SizeValue(first + 1, relation.Value.Value));
            }
            foreach (var edit in edits)
            {
                int at = indices[edit.Id] * 2;
                GeometryEditMode mode = edit.Mode;
                ShapeGeometry target = edit.Target;
                if (mode == GeometryEditMode.Exact || mode == GeometryEditMode.Move)
                {
                    ShapeGeometry requested = mode == GeometryEditMode.Move ? shapes[indices[edit.Id]] : target;
                    fixedValues.Add(new SizeValue(at, requested.Width));
                    fixedValues.Add(new SizeValue(at + 1, requested.Height));
                }
                else if (mode == GeometryEditMode.ResizeE || mode == GeometryEditMode.ResizeW || mode == GeometryEditMode.NativeWidth)
                    fixedValues.Add(new SizeValue(at, target.Width));
                else if (mode == GeometryEditMode.ResizeN || mode == GeometryEditMode.ResizeS || mode == GeometryEditMode.NativeHeight)
                    fixedValues.Add(new SizeValue(at + 1, target.Height));
                else
                {
                    // With a ratio there is only one dimension DOF. Project a
                    // corner's two requested dimensions onto that required line.
                    preferred.Add(new SizeValue(at, target.Width));
                    preferred.Add(new SizeValue(at + 1, target.Height));
                }
            }
            var baseline = shapes.SelectMany(shape => new[] { shape.Width, shape.Height }).ToArray();
            var factors = new double[count];
            var visited = new bool[count];
            var output = new double[count];
            for (int root = 0; root < count; root++)
            {
                if (visited[root]) continue;
                var members = new List<int>();
                var queue = new Queue<int>();
                queue.Enqueue(root); visited[root] = true; factors[root] = 1;
                while (queue.Count > 0)
                {
                    int current = queue.Dequeue(); members.Add(current);
                    foreach (var edge in edges[current])
                    {
                        double factor = factors[current] * edge.Factor;
                        if (!Finite(factor) || factor <= 0)
                            throw new ConstraintConflictException("Dimension ratios exceed the supported numerical range.");
                        if (!visited[edge.To])
                        { factors[edge.To] = factor; visited[edge.To] = true; queue.Enqueue(edge.To); }
                        else if (Math.Abs(factors[edge.To] - factor) > 1e-9 * Math.Max(factors[edge.To], factor))
                            throw new ConstraintConflictException("Dimension equality and ratio constraints are inconsistent.");
                    }
                }
                double scale = members.Max(i => factors[i]);
                foreach (int member in members) factors[member] /= scale;
                var memberSet = new HashSet<int>(members);
                var required = fixedValues.Where(value => memberSet.Contains(value.Index)).ToList();
                var requested = preferred.Where(value => memberSet.Contains(value.Index)).ToList();
                double lower = members.Max(i => MinimumSize / factors[i]);
                double chosen;
                if (required.Count > 0)
                {
                    chosen = required[0].Value / factors[required[0].Index];
                    if (required.Any(value => Math.Abs(factors[value.Index] * chosen - value.Value) > ResidualTolerance))
                        throw new ConstraintConflictException("The requested size conflicts with fixed dimensions, equal sizes or aspect ratios.");
                    if (chosen < lower - ResidualTolerance)
                        throw new ConstraintConflictException("The requested size would make a linked dimension smaller than one point.");
                }
                else if (requested.Count > 0)
                    chosen = requested.Sum(value => factors[value.Index] * value.Value) /
                        requested.Sum(value => factors[value.Index] * factors[value.Index]);
                else chosen = members.Sum(i => factors[i] * baseline[i]) / members.Sum(i => factors[i] * factors[i]);
                chosen = Math.Max(chosen, lower);
                if (!Finite(chosen)) throw new ConstraintConflictException("Dimension constraints exceed the supported numerical range.");
                foreach (int member in members)
                {
                    output[member] = Math.Max(MinimumSize, factors[member] * chosen);
                    if (!Finite(output[member])) throw new ConstraintConflictException("Dimension constraints exceed the supported numerical range.");
                }
            }
            return output;
        }

        private static void JoinSize(IList<SizeEdge>[] edges, int from, int to, double factor)
        { edges[from].Add(new SizeEdge(to, factor)); edges[to].Add(new SizeEdge(from, 1 / factor)); }

        private static double[] DrivenPosition(ShapeGeometry start, ShapeGeometry target, double width, double height, GeometryEditMode mode)
        {
            if (mode == GeometryEditMode.Exact || mode == GeometryEditMode.Move ||
                mode == GeometryEditMode.NativeWidth || mode == GeometryEditMode.NativeHeight) return new[] { target.Left, target.Top };
            bool west = mode == GeometryEditMode.ResizeW || mode == GeometryEditMode.ResizeNW || mode == GeometryEditMode.ResizeSW;
            bool north = mode == GeometryEditMode.ResizeN || mode == GeometryEditMode.ResizeNW || mode == GeometryEditMode.ResizeNE;
            bool verticalOnly = mode == GeometryEditMode.ResizeN || mode == GeometryEditMode.ResizeS;
            bool horizontalOnly = mode == GeometryEditMode.ResizeE || mode == GeometryEditMode.ResizeW;
            return new[] {
                verticalOnly ? start.Left + (start.Width - width) / 2 : west ? start.Left + start.Width - width : start.Left,
                horizontalOnly ? start.Top + (start.Height - height) / 2 : north ? start.Top + start.Height - height : start.Top
            };
        }

        private static double[] ResolvePositions(IList<ShapeGeometry> component, IList<LayoutConstraint> relations,
            IDictionary<string, int> indices, double[] sizes, IDictionary<string, double> requestedPositions, bool horizontal)
        {
            int count = component.Count;
            var rows = new List<double[]>();
            var values = new List<double>();
            foreach (var relation in relations)
            {
                int first = indices[relation.ShapeIds[0]];
                double? edge = AlignmentEdge(relation.Kind, horizontal);
                if (edge.HasValue)
                {
                    for (int i = 1; i < relation.ShapeIds.Count; i++)
                    {
                        int second = indices[relation.ShapeIds[i]];
                        var row = new double[count]; row[second] = 1; row[first] = -1;
                        rows.Add(row); values.Add(edge.Value * (sizes[first] - sizes[second]));
                    }
                }
                else if (relation.Kind == ConstraintKind.FixedPosition)
                {
                    var row = new double[count]; row[first] = 1;
                    rows.Add(row); values.Add(horizontal ? relation.Value.Value : relation.SecondaryValue.Value);
                }
                else if (relation.Kind == ConstraintKind.Offset)
                {
                    var row = new double[count]; row[indices[relation.ShapeIds[1]]] = 1; row[first] = -1;
                    rows.Add(row); values.Add(horizontal ? relation.Value.Value : relation.SecondaryValue.Value);
                }
                else if (relation.Kind == (horizontal ? ConstraintKind.SymmetricVertical : ConstraintKind.SymmetricHorizontal))
                {
                    int second = indices[relation.ShapeIds[1]];
                    var row = new double[count]; row[first] = 1; row[second] = 1;
                    rows.Add(row); values.Add(2 * relation.Value.Value - (sizes[first] + sizes[second]) / 2);
                }
                else if (relation.Kind == (horizontal ? ConstraintKind.FixedHorizontalGap : ConstraintKind.FixedVerticalGap))
                {
                    for (int i = 1; i < relation.ShapeIds.Count; i++)
                    {
                        int left = indices[relation.ShapeIds[i - 1]], right = indices[relation.ShapeIds[i]];
                        var row = new double[count]; row[right] = 1; row[left] = -1;
                        rows.Add(row); values.Add(sizes[left] + relation.Value.Value);
                    }
                }
                else if (relation.Kind == (horizontal ? ConstraintKind.EqualHorizontalGap : ConstraintKind.EqualVerticalGap))
                {
                    int second = indices[relation.ShapeIds[1]];
                    for (int i = 1; i < relation.ShapeIds.Count - 1; i++)
                    {
                        int left = indices[relation.ShapeIds[i]], right = indices[relation.ShapeIds[i + 1]];
                        var row = new double[count];
                        row[right] += 1; row[left] -= 1; row[second] -= 1; row[first] += 1;
                        rows.Add(row); values.Add(sizes[left] - sizes[first]);
                    }
                }
            }
            var driven = requestedPositions.ToDictionary(pair => indices[pair.Key], pair => pair.Value);
            var baseline = component.Select(shape => horizontal ? shape.Left : shape.Top).ToArray();
            foreach (var position in driven) baseline[position.Key] = position.Value;
            if (rows.Count == 0) return baseline;
            var free = Enumerable.Range(0, count).Where(index => !driven.ContainsKey(index)).ToArray();
            var rhs = Vector<double>.Build.Dense(rows.Count, row => values[row] -
                driven.Sum(position => rows[row][position.Key] * position.Value));
            if (free.Length == 0)
            {
                if (rhs.AbsoluteMaximum() > ResidualTolerance) throw PositionConflict(horizontal);
                return baseline;
            }
            var a = Matrix<double>.Build.Dense(rows.Count, free.Length, (row, column) => rows[row][free[column]]);
            var start = Vector<double>.Build.Dense(free.Length, column => baseline[free[column]]);
            Vector<double> result = start + a.PseudoInverse() * (rhs - a * start);
            if ((a * result - rhs).AbsoluteMaximum() > ResidualTolerance) throw PositionConflict(horizontal);
            for (int i = 0; i < free.Length; i++) baseline[free[i]] = result[i];
            return baseline;
        }

        private static ConstraintConflictException PositionConflict(bool horizontal)
        { return new ConstraintConflictException("The requested edit conflicts with the " + (horizontal ? "horizontal" : "vertical") + " constraints."); }

        private static double? AlignmentEdge(ConstraintKind kind, bool horizontal)
        {
            if (kind == ConstraintKind.CentersCoincident) return 0.5;
            if (kind == (horizontal ? ConstraintKind.SameX : ConstraintKind.SameY) ||
                kind == (horizontal ? ConstraintKind.SymmetricHorizontal : ConstraintKind.SymmetricVertical)) return 0;
            if (kind == (horizontal ? ConstraintKind.RightAligned : ConstraintKind.BottomAligned)) return 1;
            if (kind == (horizontal ? ConstraintKind.CenterXAligned : ConstraintKind.CenterYAligned)) return 0.5;
            return null;
        }

        private static double Attribute(ShapeGeometry shape, ConstraintKind kind)
        {
            switch (kind)
            {
                case ConstraintKind.SameX: return shape.Left;
                case ConstraintKind.SameY: return shape.Top;
                case ConstraintKind.SameWidth: return shape.Width;
                case ConstraintKind.SameHeight: return shape.Height;
                case ConstraintKind.RightAligned: return shape.Left + shape.Width;
                case ConstraintKind.BottomAligned: return shape.Top + shape.Height;
                case ConstraintKind.CenterXAligned: return shape.Left + shape.Width / 2;
                case ConstraintKind.CenterYAligned: return shape.Top + shape.Height / 2;
                default: throw new ArgumentOutOfRangeException("kind");
            }
        }

        private static double Gap(ShapeGeometry left, ShapeGeometry right, bool horizontal)
        { return horizontal ? right.Left - left.Left - left.Width : right.Top - left.Top - left.Height; }
        private static bool Finite(double value) { return !Double.IsNaN(value) && !Double.IsInfinity(value); }
        private sealed class SizeEdge
        {
            public readonly int To; public readonly double Factor;
            public SizeEdge(int to, double factor) { To = to; Factor = factor; }
        }
        private sealed class SizeValue
        {
            public readonly int Index; public readonly double Value;
            public SizeValue(int index, double value) { Index = index; Value = value; }
        }
    }
}
