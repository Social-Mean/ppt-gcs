using System;

namespace PptGcs.Core
{
    public sealed class ShapeTopUpdate
    {
        public string ShapeId { get; private set; }
        public double Top { get; private set; }

        internal ShapeTopUpdate(string shapeId, double top)
        {
            ShapeId = shapeId;
            Top = top;
        }
    }

    // Exact solution of yA = yB with either variable supplied as the edit target.
    // Office sampling, undo boundaries, and document persistence belong to the host.
    public static class SameTopSolver
    {
        public const double PositionTolerance = 0.01; // PowerPoint points.

        public static ShapeTopUpdate Solve(SameTopConstraint constraint, string driverShapeId, double driverTop)
        {
            if (constraint == null) throw new ArgumentNullException("constraint");
            ValidateCoordinate(driverTop, "driverTop");
            string driver = ShapeIdentity.Normalize(driverShapeId);
            if (driver == constraint.ShapeA) return new ShapeTopUpdate(constraint.ShapeB, driverTop);
            if (driver == constraint.ShapeB) return new ShapeTopUpdate(constraint.ShapeA, driverTop);
            throw new ArgumentException("The driver shape is not a member of this constraint.", "driverShapeId");
        }

        public static bool NeedsUpdate(double currentTop, double targetTop)
        {
            ValidateCoordinate(currentTop, "currentTop");
            ValidateCoordinate(targetTop, "targetTop");
            return Math.Abs(currentTop - targetTop) > PositionTolerance;
        }

        private static void ValidateCoordinate(double value, string parameter)
        {
            if (Double.IsNaN(value) || Double.IsInfinity(value))
                throw new ArgumentOutOfRangeException(parameter, "A finite coordinate is required.");
        }
    }
}
