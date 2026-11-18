using System;

namespace PptGcs.AddIn
{
    public static class KeyboardNudgePolicy
    {
        public static bool IsArrow(int key) { return key>=0x25 && key<=0x28; }

        public static bool CanOwn(int key,int selectionType,bool shift,bool alt,bool canvasFocus,bool activeWindow)
        { return IsArrow(key) && selectionType==2 && !shift && !alt && canvasFocus && activeWindow; }

        public static double PixelsPerPoint(double atZero,double atTenThousand)
        {
            double scale=(atTenThousand-atZero)/10000.0;
            if(!Finite(scale) || scale<=0)throw new InvalidOperationException("PowerPoint returned an invalid screen scale.");
            return scale;
        }

        public static double StepPoints(bool snapToGrid,double gridDistance,bool control,double pixelsPerPoint)
        {
            if(!Finite(pixelsPerPoint) || pixelsPerPoint<=0)throw new ArgumentOutOfRangeException("pixelsPerPoint");
            if(snapToGrid && !control)
            {
                if(!Finite(gridDistance) || gridDistance<=0)throw new InvalidOperationException("PowerPoint returned an invalid grid spacing.");
                return gridDistance;
            }
            return 1.0/pixelsPerPoint;
        }

        public static Tuple<double,double> Delta(int key,double horizontalStep,double verticalStep)
        {
            if(!Finite(horizontalStep)||horizontalStep<=0||!Finite(verticalStep)||verticalStep<=0)
                throw new ArgumentOutOfRangeException("step");
            switch(key)
            {
                case 0x25:return Tuple.Create(-horizontalStep,0.0);
                case 0x26:return Tuple.Create(0.0,-verticalStep);
                case 0x27:return Tuple.Create(horizontalStep,0.0);
                case 0x28:return Tuple.Create(0.0,verticalStep);
                default:throw new ArgumentOutOfRangeException("key");
            }
        }

        private static bool Finite(double value) { return !Double.IsNaN(value)&&!Double.IsInfinity(value); }
    }
}
