using System;
using System.Collections.Generic;
using System.Linq;

namespace PptGcs.Core
{
    public sealed class ConstraintSuggestion
    {
        public ConstraintKind Kind { get; }
        public string AnchorId { get; }
        public IReadOnlyList<string> MemberIds { get; }
        internal double Score { get; }

        internal ConstraintSuggestion(ConstraintKind kind, string anchorId, IEnumerable<ShapeGeometry> members, double score)
        { Kind = kind; AnchorId = anchorId; MemberIds = members.Select(shape => shape.Id).ToList().AsReadOnly(); Score = score; }
    }

    public static class ConstraintSuggestions
    {
        public const double Tolerance = 4;
        private const double NeighborDistance = 320;
        private const int MaximumOffers = 3;
        private static readonly ConstraintKind[] PairKinds = {
            ConstraintKind.SameY, ConstraintKind.SameX, ConstraintKind.CenterYAligned, ConstraintKind.CenterXAligned,
            ConstraintKind.BottomAligned, ConstraintKind.RightAligned, ConstraintKind.SameWidth, ConstraintKind.SameHeight
        };

        public static IReadOnlyList<ConstraintSuggestion> Find(int slideId, IReadOnlyList<ShapeGeometry> shapes,
            IEnumerable<string> editedIds, IReadOnlyList<LayoutConstraint> existing)
        {
            var edited = new HashSet<string>(editedIds, StringComparer.Ordinal);
            var viable = shapes.Where(shape => shape.Width >= LayoutSolver.MinimumSize && shape.Height >= LayoutSolver.MinimumSize).ToList();
            var proposals = new List<ConstraintSuggestion>();
            foreach (var driver in viable.Where(shape => edited.Contains(shape.Id)))
            {
                var local = viable.Where(shape => Distance(driver, shape) <= NeighborDistance).ToList();
                foreach (var peer in local.Where(shape => shape.Id != driver.Id))
                {
                    var pair = new[] { driver, peer };
                    foreach (var kind in PairKinds)
                    {
                        double error = Error(kind, pair);
                        if (error <= Tolerance)
                            proposals.Add(new ConstraintSuggestion(kind, driver.Id, pair, error + Distance(driver, peer) / 200 + 1));
                    }
                }
                AddSpacings(proposals, local, driver, true);
                AddSpacings(proposals, local, driver, false);
            }

            var byId = shapes.ToDictionary(shape => shape.Id, StringComparer.Ordinal);
            var current = existing.Where(item => item.SlideId == slideId).ToList();
            var active = current.Where(item => item.Enabled).ToList();
            if (active.SelectMany(item => item.ShapeIds).Any(id => !byId.ContainsKey(id))) return Array.Empty<ConstraintSuggestion>();
            var tested = new HashSet<string>(StringComparer.Ordinal);
            var shown = new HashSet<string>(StringComparer.Ordinal);
            var offers = new List<ConstraintSuggestion>();
            int attempts = 0;
            foreach (var item in proposals.OrderBy(item => item.Score))
            {
                string members = String.Join(",", item.MemberIds.OrderBy(id => id, StringComparer.Ordinal));
                if (!tested.Add(item.Kind + "/" + members) || shown.Contains(Family(item.Kind) + "/" + members) ||
                    AlreadyRelated(current, item.Kind, item.MemberIds)) continue;
                if (attempts++ == 18) break;
                var candidate = new LayoutConstraint(Guid.NewGuid().ToString("D"), slideId, item.Kind, item.MemberIds);
                try
                {
                    var result = LayoutSolver.Normalize(shapes, active.Concat(new[] { candidate }));
                    if (result.Any(shape => Change(shape, byId[shape.Id]) > 2 * Tolerance)) continue;
                }
                catch (ConstraintConflictException) { continue; }
                shown.Add(Family(item.Kind) + "/" + members);
                offers.Add(item);
                if (offers.Count == MaximumOffers) break;
            }
            return offers.AsReadOnly();
        }

        public static bool AlreadyRelated(IEnumerable<LayoutConstraint> existing, ConstraintKind kind, IReadOnlyList<string> members)
        {
            bool gap = kind == ConstraintKind.EqualHorizontalGap || kind == ConstraintKind.EqualVerticalGap;
            var alike = existing.Where(item => item.Kind == kind).ToList();
            if (gap)
            {
                foreach (var item in alike)
                {
                    var positions = members.Select(item.ShapeIds.IndexOf).ToList();
                    if (positions.Min() >= 0 && positions.Max() - positions.Min() == members.Count - 1) return true;
                }
                return false;
            }
            if (alike.Any(item => members.All(item.ShapeIds.Contains))) return true;
            // Equality through a chain already supplies the same pair relation.
            var reached = new HashSet<string>(StringComparer.Ordinal) { members[0] };
            bool changed;
            do
            {
                changed = false;
                foreach (var link in alike.Where(item => item.Enabled && item.ShapeIds.Any(reached.Contains)))
                    foreach (string id in link.ShapeIds) changed |= reached.Add(id);
            } while (changed && !reached.Contains(members[1]));
            return reached.Contains(members[1]);
        }

        public static double Error(ConstraintKind kind, IReadOnlyList<ShapeGeometry> members)
        {
            bool horizontalGap = kind == ConstraintKind.EqualHorizontalGap;
            if (horizontalGap || kind == ConstraintKind.EqualVerticalGap)
            {
                if (members.Count != 3 || !SameLane(members[0], members[1], horizontalGap) ||
                    !SameLane(members[1], members[2], horizontalGap) || !SameLane(members[0], members[2], horizontalGap)) return Double.PositiveInfinity;
                var ordered = members.OrderBy(shape => Start(shape, horizontalGap)).ToArray();
                double first = Start(ordered[1], horizontalGap) - End(ordered[0], horizontalGap);
                double second = Start(ordered[2], horizontalGap) - End(ordered[1], horizontalGap);
                if (first < 0 || second < 0) return Double.PositiveInfinity;
                return Math.Abs(second - first);
            }
            if (members.Count != 2) return Double.PositiveInfinity;
            return Math.Abs(Coordinate(kind, members[0]) - Coordinate(kind, members[1]));
        }

        private static void AddSpacings(List<ConstraintSuggestion> proposals, List<ShapeGeometry> shapes, ShapeGeometry driver, bool horizontal)
        {
            var lane = shapes.Where(shape => SameLane(shape, driver, horizontal)).OrderBy(shape => Start(shape, horizontal)).ToList();
            int at = lane.IndexOf(driver);
            var kind = horizontal ? ConstraintKind.EqualHorizontalGap : ConstraintKind.EqualVerticalGap;
            for (int index = Math.Max(0, at - 2); index <= at && index + 2 < lane.Count; index++)
            {
                var trio = lane.GetRange(index, 3);
                double error = Error(kind, trio);
                if (error <= Tolerance)
                    proposals.Add(new ConstraintSuggestion(kind, driver.Id, trio, error * 0.5 + Distance(trio[0], trio[2]) / 250));
            }
        }

        private static bool SameLane(ShapeGeometry a, ShapeGeometry b, bool horizontal)
        {
            double overlap = Math.Min(End(a, !horizontal), End(b, !horizontal)) - Math.Max(Start(a, !horizontal), Start(b, !horizontal));
            return overlap >= 0.25 * Math.Min(End(a, !horizontal) - Start(a, !horizontal), End(b, !horizontal) - Start(b, !horizontal));
        }

        private static double Coordinate(ConstraintKind kind, ShapeGeometry shape)
        {
            switch (kind)
            {
                case ConstraintKind.SameX: return shape.Left;
                case ConstraintKind.SameY: return shape.Top;
                case ConstraintKind.RightAligned: return shape.Left + shape.Width;
                case ConstraintKind.BottomAligned: return shape.Top + shape.Height;
                case ConstraintKind.CenterXAligned: return shape.Left + shape.Width / 2;
                case ConstraintKind.CenterYAligned: return shape.Top + shape.Height / 2;
                case ConstraintKind.SameWidth: return shape.Width;
                case ConstraintKind.SameHeight: return shape.Height;
                default: throw new ArgumentOutOfRangeException(nameof(kind));
            }
        }

        private static string Family(ConstraintKind kind)
        {
            switch (kind)
            {
                case ConstraintKind.SameX: case ConstraintKind.RightAligned: case ConstraintKind.CenterXAligned: return "x";
                case ConstraintKind.SameY: case ConstraintKind.BottomAligned: case ConstraintKind.CenterYAligned: return "y";
                default: return kind.ToString();
            }
        }

        private static double Start(ShapeGeometry shape, bool horizontal) { return horizontal ? shape.Left : shape.Top; }
        private static double End(ShapeGeometry shape, bool horizontal) { return horizontal ? shape.Left + shape.Width : shape.Top + shape.Height; }
        private static double Distance(ShapeGeometry a, ShapeGeometry b)
        {
            double x = Math.Max(0, Math.Max(a.Left - b.Left - b.Width, b.Left - a.Left - a.Width));
            double y = Math.Max(0, Math.Max(a.Top - b.Top - b.Height, b.Top - a.Top - a.Height));
            return Math.Sqrt(x * x + y * y);
        }
        private static double Change(ShapeGeometry a, ShapeGeometry b)
        { return Math.Max(Math.Max(Math.Abs(a.Left - b.Left), Math.Abs(a.Top - b.Top)), Math.Max(Math.Abs(a.Width - b.Width), Math.Abs(a.Height - b.Height))); }
    }
}
