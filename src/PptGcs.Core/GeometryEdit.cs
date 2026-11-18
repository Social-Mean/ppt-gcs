using System;

namespace PptGcs.Core
{
    // One member of an atomic layout edit. Exact pins the full native target;
    // Move and resize modes preserve the same anchors as a pointer gesture.
    public sealed class GeometryEdit
    {
        public ShapeGeometry Target { get; private set; }
        public GeometryEditMode Mode { get; private set; }
        public string Id { get { return Target.Id; } }

        public GeometryEdit(ShapeGeometry target, GeometryEditMode mode = GeometryEditMode.Exact)
        {
            if (target == null) throw new ArgumentNullException("target");
            if (!Enum.IsDefined(typeof(GeometryEditMode), mode)) throw new ArgumentOutOfRangeException("mode");
            Target = target;
            Mode = mode;
        }
    }
}
