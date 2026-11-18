using System;

namespace PptGcs.Core
{
    public sealed class RelativeAnchor
    {
        private static readonly RelativeAnchor[] PageAnchors=Build(true);
        private static readonly RelativeAnchor[] ContainerAnchors=Build(false);

        public bool OnPage { get; }
        public bool Horizontal { get; }
        public double Fraction { get; }
        public double MarginSign { get { return Fraction==1 ? -1 : 1; } }

        private RelativeAnchor(bool onPage,bool horizontal,double fraction)
        { OnPage=onPage;Horizontal=horizontal;Fraction=fraction; }

        private static RelativeAnchor[] Build(bool onPage)
        {
            return new[] {
                new RelativeAnchor(onPage,true,0), new RelativeAnchor(onPage,true,0.5), new RelativeAnchor(onPage,true,1),
                new RelativeAnchor(onPage,false,0), new RelativeAnchor(onPage,false,0.5), new RelativeAnchor(onPage,false,1)
            };
        }

        public static RelativeAnchor For(ConstraintKind kind)
        {
            if(kind>=ConstraintKind.PageLeft && kind<=ConstraintKind.PageBottom)
                return PageAnchors[(int)kind-(int)ConstraintKind.PageLeft];
            if(kind>=ConstraintKind.ContainerLeft && kind<=ConstraintKind.ContainerBottom)
                return ContainerAnchors[(int)kind-(int)ConstraintKind.ContainerLeft];
            return null;
        }

        public static bool UsesContainer(ConstraintKind kind)
        { var anchor=For(kind);return anchor!=null && !anchor.OnPage; }
    }
}
