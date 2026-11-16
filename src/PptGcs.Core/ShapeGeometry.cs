using System;

namespace PptGcs.Core
{
    public sealed class ShapeGeometry
    {
        public string Id { get; private set; }
        public double Left { get; private set; }
        public double Top { get; private set; }
        public double Width { get; private set; }
        public double Height { get; private set; }

        public ShapeGeometry(string id, double left, double top, double width, double height)
        {
            Id = ShapeIdentity.Normalize(id);
            ValidateFinite(left, "left"); ValidateFinite(top, "top");
            ValidateFinite(width, "width"); ValidateFinite(height, "height");
            if (width < LayoutSolver.MinimumSize || height < LayoutSolver.MinimumSize)
                throw new ArgumentOutOfRangeException("width/height", "Shape dimensions must be at least 1 point.");
            Left = left; Top = top; Width = width; Height = height;
        }

        private static void ValidateFinite(double value, string parameter)
        {
            if (Double.IsNaN(value) || Double.IsInfinity(value))
                throw new ArgumentOutOfRangeException(parameter, "A finite coordinate is required.");
        }
    }
}
