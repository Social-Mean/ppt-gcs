using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using PptGcs.Core;

namespace PptGcs.AddIn
{
    // Native-only selection members need no GCS dimensions. Lines and other
    // objects with a zero-width/height bounding box still translate rigidly.
    public sealed class SelectionMovePosition
    {
        public int NativeId { get; private set; }
        public double Left { get; private set; }
        public double Top { get; private set; }
        public double Width { get; private set; }
        public double Height { get; private set; }

        public SelectionMovePosition(int nativeId,double left,double top,double width,double height)
        {
            if(nativeId<=0)throw new ArgumentOutOfRangeException("nativeId");
            if(!Finite(left)||!Finite(top)||!Finite(width)||!Finite(height)||width<0||height<0)
                throw new ArgumentOutOfRangeException("geometry","Native selection bounds must be finite and nonnegative in size.");
            NativeId=nativeId;Left=left;Top=top;Width=width;Height=height;
        }
        public SelectionMovePosition Translate(double dx,double dy)
        {return new SelectionMovePosition(NativeId,Left+dx,Top+dy,Width,Height);}
        private static bool Finite(double value){return !Double.IsNaN(value)&&!Double.IsInfinity(value);}
    }

    public sealed class SelectionMoveFrame
    {
        public ReadOnlyCollection<ShapeGeometry> Geometry { get; private set; }
        public ReadOnlyCollection<SelectionMovePosition> Extras { get; private set; }
        public SelectionMoveFrame(IEnumerable<ShapeGeometry> geometry,IEnumerable<SelectionMovePosition> extras)
        {
            if(geometry==null)throw new ArgumentNullException("geometry");
            if(extras==null)throw new ArgumentNullException("extras");
            var shapes=geometry.ToList();var positions=extras.ToList();
            if(shapes.Any(shape=>shape==null)||positions.Any(position=>position==null))throw new ArgumentException("A frame member cannot be null.");
            if(shapes.Select(shape=>shape.Id).Distinct(StringComparer.Ordinal).Count()!=shapes.Count ||
                positions.Select(position=>position.NativeId).Distinct().Count()!=positions.Count)
                throw new ArgumentException("A frame cannot contain duplicate identities.");
            Geometry=shapes.AsReadOnly();Extras=positions.AsReadOnly();
        }
        public bool Matches(SelectionMoveFrame other)
        {
            if(other==null || Geometry.Count!=other.Geometry.Count || Extras.Count!=other.Extras.Count)return false;
            var shapes=other.Geometry.ToDictionary(shape=>shape.Id,StringComparer.Ordinal);
            foreach(var shape in Geometry)
            {
                ShapeGeometry expected;
                if(!shapes.TryGetValue(shape.Id,out expected)||LayoutSolver.NeedsUpdate(shape,expected))return false;
            }
            var positions=other.Extras.ToDictionary(position=>position.NativeId);
            foreach(var position in Extras)
            {
                SelectionMovePosition expected;
                if(!positions.TryGetValue(position.NativeId,out expected)||
                    Math.Abs(position.Left-expected.Left)>LayoutSolver.PositionTolerance ||
                    Math.Abs(position.Top-expected.Top)>LayoutSolver.PositionTolerance ||
                    Math.Abs(position.Width-expected.Width)>LayoutSolver.PositionTolerance ||
                    Math.Abs(position.Height-expected.Height)>LayoutSolver.PositionTolerance)return false;
            }
            return true;
        }
    }

    // An Office write was rejected, but the previous accepted frame has been
    // restored. The pointer can continue without partially moving the selection.
    public sealed class SelectionMoveRejectedException : InvalidOperationException
    {
        public SelectionMoveRejectedException(string message,Exception inner):base(message,inner){}
    }

    public sealed class SelectionMovePlan
    {
        private readonly List<LayoutConstraint> constraints;
        private readonly List<ShapeGeometry> drivers;
        public SelectionMoveFrame Initial { get; private set; }
        public ReadOnlyCollection<string> DriverIds { get; private set; }

        public SelectionMovePlan(IEnumerable<ShapeGeometry> baseline,IEnumerable<LayoutConstraint> relations,
            IEnumerable<string> driverIds,IEnumerable<SelectionMovePosition> extras)
        {
            if(relations==null)throw new ArgumentNullException("relations");
            if(driverIds==null)throw new ArgumentNullException("driverIds");
            Initial=new SelectionMoveFrame(baseline,extras);
            var ids=driverIds.Select(ShapeIdentity.Normalize).ToList();
            if(ids.Count==0 || ids.Distinct(StringComparer.Ordinal).Count()!=ids.Count)
                throw new ArgumentException("A selected constrained driver is required and cannot be repeated.","driverIds");
            var geometry=Initial.Geometry.ToDictionary(shape=>shape.Id,StringComparer.Ordinal);
            if(ids.Any(id=>!geometry.ContainsKey(id)))throw new ArgumentException("A selected driver is absent from the layout.","driverIds");
            constraints=LayoutSolver.ConnectedConstraints(relations,ids).ToList();
            drivers=ids.Select(id=>geometry[id]).ToList();
            DriverIds=ids.AsReadOnly();
        }

        public SelectionMoveFrame Solve(double dx,double dy)
        {
            return new SelectionMoveFrame(LayoutSolver.Solve(Initial.Geometry,constraints,RequestedEdits(dx,dy)),
                Initial.Extras.Select(position=>position.Translate(dx,dy)));
        }

        internal IReadOnlyList<GeometryEdit> RequestedEdits(double dx,double dy)
        {
            if(Double.IsNaN(dx)||Double.IsInfinity(dx)||Double.IsNaN(dy)||Double.IsInfinity(dy))
                throw new ArgumentOutOfRangeException("delta");
            return drivers.Select(shape=>new GeometryEdit(new ShapeGeometry(shape.Id,
                shape.Left+dx,shape.Top+dy,shape.Width,shape.Height),GeometryEditMode.Move)).ToList();
        }

        public SelectionMoveFrame ApplyAtomic(SelectionMoveFrame previous,SelectionMoveFrame requested,
            Action<SelectionMoveFrame> write,Func<SelectionMoveFrame> read)
        {
            if(previous==null)throw new ArgumentNullException("previous");
            if(requested==null)throw new ArgumentNullException("requested");
            if(write==null)throw new ArgumentNullException("write");
            if(read==null)throw new ArgumentNullException("read");
            try
            {
                write(requested);
                var actual=read();
                if(!actual.Matches(requested))throw new InvalidOperationException("PowerPoint did not retain the complete selection translation.");
                if(LayoutSolver.MaxResidual(actual.Geometry,constraints)>LayoutSolver.PositionTolerance)
                    throw new InvalidOperationException("PowerPoint geometry violates a constraint after selection writeback.");
                return actual;
            }
            catch(Exception error)
            {
                try
                {
                    write(previous);
                    if(!read().Matches(previous))throw new InvalidOperationException("PowerPoint did not restore the previous selection frame.");
                }
                catch(Exception rollback)
                {throw new InvalidOperationException("Cannot restore a rejected selection move.",new AggregateException(error,rollback));}
                throw new SelectionMoveRejectedException("The requested selection move was rejected and the previous frame was restored.",error);
            }
        }
    }
}
