using System;
using PptGcs.AddIn;

public static class NativeFieldTargetTests
{
    private static int count;
    public static int Main()
    {
        try
        {
            var before=new SelectionMovePosition(1,20,30,120,80);
            Matches(NativeFieldTarget.Create(before,NativeFieldProperty.Left,300,false),300,30,120,80);
            Matches(NativeFieldTarget.Create(before,NativeFieldProperty.Left,-17.5,true),-17.5,30,120,80);
            Matches(NativeFieldTarget.Create(before,NativeFieldProperty.Top,-25,true),20,-25,120,80);
            Matches(NativeFieldTarget.Create(before,NativeFieldProperty.Top,0,false),20,0,120,80);
            Matches(NativeFieldTarget.Create(before,NativeFieldProperty.Width,240,false),20,30,240,80);
            Matches(NativeFieldTarget.Create(before,NativeFieldProperty.Width,240,true),20,30,240,160);
            Matches(NativeFieldTarget.Create(before,NativeFieldProperty.Height,40,false),20,30,120,40);
            Matches(NativeFieldTarget.Create(before,NativeFieldProperty.Height,40,true),20,30,60,40);
            var line=new SelectionMovePosition(2,15,40,0,80);
            Matches(NativeFieldTarget.Create(line,NativeFieldProperty.Left,-40,true),-40,40,0,80);
            Matches(NativeFieldTarget.Create(line,NativeFieldProperty.Width,20,false),15,40,20,80);
            Rejects(()=>NativeFieldTarget.Create(line,NativeFieldProperty.Width,20,true));
            Rejects(()=>NativeFieldTarget.Create(before,NativeFieldProperty.Width,0.99,false));
            Rejects(()=>NativeFieldTarget.Create(before,NativeFieldProperty.Height,-1,false));
            Rejects(()=>NativeFieldTarget.Create(before,NativeFieldProperty.Left,Double.NaN,false));
            Rejects(()=>NativeFieldTarget.Create(before,NativeFieldProperty.Top,Double.PositiveInfinity,false));
            Rejects(()=>NativeFieldTarget.Create(before,NativeFieldProperty.Left,-1e40,false));
            Rejects(()=>NativeFieldTarget.Create(before,(NativeFieldProperty)99,40,false));
            // A shared native width retains each selected object's own ratio.
            Matches(NativeFieldTarget.Create(new SelectionMovePosition(3,10,20,40,80),NativeFieldProperty.Width,96,true),10,20,96,192);
            Matches(NativeFieldTarget.Create(new SelectionMovePosition(4,50,60,160,40),NativeFieldProperty.Width,96,true),50,60,96,24);
            double center=NativeFieldTarget.AbsolutePosition(NativeFieldProperty.Left,NativePositionOrigin.SlideCenter,0,960,540);
            Matches(NativeFieldTarget.Create(new SelectionMovePosition(5,100,70,170.078735,100),NativeFieldProperty.Left,center,true),480,70,170.078735,100);
            center=NativeFieldTarget.AbsolutePosition(NativeFieldProperty.Left,NativePositionOrigin.SlideCenter,-72,960,540);
            Matches(NativeFieldTarget.Create(before,NativeFieldProperty.Left,center,false),408,30,120,80);
            center=NativeFieldTarget.AbsolutePosition(NativeFieldProperty.Left,NativePositionOrigin.SlideCenter,28.5,1000,700);
            Matches(NativeFieldTarget.Create(before,NativeFieldProperty.Left,center,false),528.5,30,120,80);
            center=NativeFieldTarget.AbsolutePosition(NativeFieldProperty.Top,NativePositionOrigin.SlideCenter,0,960,540);
            Matches(NativeFieldTarget.Create(before,NativeFieldProperty.Top,center,true),20,270,120,80);
            center=NativeFieldTarget.AbsolutePosition(NativeFieldProperty.Top,NativePositionOrigin.SlideCenter,-100,960,540);
            Matches(NativeFieldTarget.Create(before,NativeFieldProperty.Top,center,false),20,170,120,80);
            Console.WriteLine("Native field target tests passed: "+count);return 0;
        }
        catch(Exception error){Console.Error.WriteLine(error);return 1;}
    }
    private static void Matches(SelectionMovePosition actual,double left,double top,double width,double height)
    {if(Math.Abs(actual.Left-left)>1e-8||Math.Abs(actual.Top-top)>1e-8||Math.Abs(actual.Width-width)>1e-8||Math.Abs(actual.Height-height)>1e-8)throw new Exception("Unexpected native target.");count++;}
    private static void Rejects(Action action)
    {try{action();}catch(ArgumentException){count++;return;}catch(InvalidOperationException){count++;return;}catch(NotSupportedException){count++;return;}throw new Exception("Invalid native target was accepted.");}
}
