using System;
using System.Collections.Generic;
using System.Drawing;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    // Plain UI data: the task pane never reads or writes PowerPoint objects.
    public sealed class ConstraintPaneModel
    {
        public string PageLabel { get; set; }
        public string Message { get; set; }
        public IList<ConstraintPaneObject> Objects { get; set; } = new List<ConstraintPaneObject>();
        public IList<ConstraintPaneRelation> Relations { get; set; } = new List<ConstraintPaneRelation>();
        public IList<ConstraintPaneOption> Options { get; set; } = new List<ConstraintPaneOption>();
    }

    public sealed class ConstraintPaneObject
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public bool Selected { get; set; }
        public int RelationCount { get; set; }
        public string UnsupportedReason { get; set; }
        public ShapeGeometry Geometry { get; set; }
        public GeometryFreedom Freedom { get; set; }
        internal RectangleF? CanvasBounds { get; set; }
    }

    public sealed class ConstraintPaneRelation
    {
        public string Id { get; set; }
        public string Kind { get; set; }
        public string Label { get; set; }
        public IList<string> MemberIds { get; set; } = new List<string>();
        // Keep member order for directional and equal-gap relations.
        public IList<string> MemberNames { get; set; } = new List<string>();
        public double? Value { get; set; }
        public double? Second { get; set; }
        public bool Enabled { get; set; } = true;
        public string Status { get; set; }
    }

    public sealed class ConstraintPaneOption
    {
        public string Kind { get; set; }
        public string Label { get; set; }
        public string Description { get; set; }
        public int MinimumMembers { get; set; } = 2;
        // Zero means no upper limit.
        public int MaximumMembers { get; set; }
        public string ValueLabel { get; set; }
        public string SecondLabel { get; set; }
        public double? DefaultValue { get; set; }
        public double? DefaultSecond { get; set; }
        public double MinimumValue { get; set; } = -1000000;
        public double MaximumValue { get; set; } = 1000000;
        public double MinimumSecond { get; set; } = -1000000;
        public double MaximumSecond { get; set; } = 1000000;
        public override string ToString() { return Label ?? Kind; }
    }

    public sealed class ConstraintPaneActions
    {
        public Action<IReadOnlyList<string>> SelectShapes { get; set; }
        public Action<string, double?, double?> Add { get; set; }
        public Action<string, double?, double?> Update { get; set; }
        public Action<string, bool> SetEnabled { get; set; }
        public Action<string> Delete { get; set; }
        public Action Refresh { get; set; }
        public Action<ConstraintCommands.DraftConflict,IReadOnlyList<string>> ResolveConflict { get; set; }
    }
}
