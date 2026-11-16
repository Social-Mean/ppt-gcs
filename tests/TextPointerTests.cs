using System;
using System.Drawing;
using PptGcs.AddIn;

internal static class TextPointerTests
{
    private static int passed;
    private static readonly RectangleF Box = new RectangleF(100, 100, 200, 80);
    private static void Main()
    {
        Check(Native(17,false,160,130), "first click in a text-box interior remains native");
        Check(Native(17,false,280,150), "blank text-box interior accepts a native click");
        Check(Native(17,true,160,130), "caret and text selection remain native");
        Check(Move(false,true,17,false,Box,new PointF(101,130),6), "an unselected text-box border can begin a constrained drag immediately");
        Check(!Native(17,false,102,130), "selected left border is not claimed as text input");
        Check(!Native(17,true,298,130), "right border remains movable while editing text");
        Check(!Native(17,true,180,102), "top border remains movable while editing text");
        Check(!Native(17,true,180,178), "bottom border remains movable while editing text");
        Check(!Native(17,true,80,130), "outside clicks are not claimed as text input");
        Check(!Native(1,false,160,130), "ordinary shape-body drag remains constrained");
        Check(Native(1,true,160,130), "already editing shape text retains native caret handling");
        Check(!Native(13,false,160,130), "picture-body drag remains constrained");
        Check(TextPointerPolicy.UseNativeTextInput(17,false,new RectangleF(200,200,400,160),new PointF(213,250),12), "text hit testing follows device-pixel scale");
        Check(Move(true,null,17,true,new RectangleF(200,200,400,160),new PointF(210,250),12), "scaled border preserves constrained movement while editing text");

        // Actual 2026-11-16 physical pointer routes at screen scale 1.85 and 96 DPI.
        // The former strict body test missed the selection outline just outside
        // these bounds, allowing native movement without its constrained peer.
        var actualA=new RectangleF(590f,620.5f,481f,112.08393f);
        var actualOutside=new PointF(1006,736);
        Check(Move(false,true,17,false,actualA,actualOutside,6), "recorded bottom outline 3.416 px outside starts an unselected constrained drag");
        Check(Move(true,null,17,true,actualA,actualOutside,6), "recorded selected outline works when native RangeFromPoint returns null");
        Check(!Move(true,false,17,true,actualA,actualOutside,6), "an overlapping object's native hit wins over the selected outline");
        var movedA=new RectangleF(585.99394f,672.46576f,481f,112.08393f);
        var actualInside=new PointF(1041,781);
        Check(Move(false,true,17,false,movedA,actualInside,6), "recorded unselected bottom band 3.55 px inside remains constrained");
        Check(!TextPointerPolicy.UseNativeTextInput(17,false,movedA,actualInside,6), "recorded unselected border no longer routes to native text input");
        var actualB=new RectangleF(1367f,620.5f,481f,112.08393f);
        Check(Move(false,true,17,false,actualB,new PointF(1496,733),6), "recorded B outline 0.416 px outside moves B and can drive its peer");

        foreach(var point in new[]{new PointF(97,140),new PointF(303,140),new PointF(180,97),new PointF(180,183)})
        {
            Check(Move(true,null,17,true,Box,point,6), "selected outer border supports all four sides at "+point);
            Check(!TextPointerPolicy.UseNativeTextInput(17,true,Box,point,6), "outer border is never text input at "+point);
        }
        Check(TextPointerPolicy.HitMoveBorder(17,false,Box,new PointF(94,140),6), "exact outer tolerance is included");
        Check(!TextPointerPolicy.HitMoveBorder(17,false,Box,new PointF(93,140),6), "blank canvas beyond the outer tolerance is excluded");
        Check(!Move(true,true,17,false,Box,new PointF(80,140),6), "a distant native hit cannot expand the allowed border area");
        Check(!Native(17,false,106,140) && TextPointerPolicy.HitMoveBorder(17,false,Box,new PointF(106,140),6), "exact inner tolerance belongs only to border movement");
        Check(Native(17,false,106.1f,140) && !TextPointerPolicy.HitMoveBorder(17,false,Box,new PointF(106.1f,140),6), "the text interior and border do not overlap");

        Check(TextPointerPolicy.CanUseBorderHit(true,true), "selected native self-hit is eligible");
        Check(TextPointerPolicy.CanUseBorderHit(false,true), "unselected native self-hit is eligible");
        Check(TextPointerPolicy.CanUseBorderHit(true,null), "selected outline can fill a native blank hit");
        Check(!TextPointerPolicy.CanUseBorderHit(false,null), "unselected objects never guess a native blank hit");
        Check(!TextPointerPolicy.CanUseBorderHit(true,false), "selected objects cannot steal another object's hit");
        Check(!TextPointerPolicy.CanUseBorderHit(false,false), "unselected objects cannot steal another object's hit");

        var dpi125=new RectangleF(125,125,250,100);
        Check(Move(true,null,17,true,dpi125,new PointF(118,170),7.5f), "125 percent DPI scales the external selection outline tolerance");
        Check(!Move(true,null,17,true,dpi125,new PointF(117,170),7.5f), "125 percent DPI does not absorb distant blank canvas");
        Check(TextPointerPolicy.UseNativeTextInput(17,true,dpi125,new PointF(133,170),7.5f), "125 percent DPI keeps interior text editing native");
        var dpi150=new RectangleF(150,150,300,120);
        Check(Move(true,null,17,true,dpi150,new PointF(141.5f,200),9), "150 percent DPI scales the external selection outline tolerance");
        Check(TextPointerPolicy.UseNativeTextInput(17,true,dpi150,new PointF(159.5f,200),9), "150 percent DPI keeps interior text editing native");
        Check(!TextPointerPolicy.HitMoveBorder(1,false,Box,new PointF(97,140),6), "ordinary shape outlines do not gain text-specific hit tolerance");
        Check(TextPointerPolicy.HitMoveBorder(1,true,Box,new PointF(97,140),6), "an actively edited shape text box uses the same border policy");
        Check(TextPointerPolicy.UseNativeTextInput(17,false,new RectangleF(100,100,8,8),new PointF(104,104),6), "a small text box still retains an editable interior");
        Console.WriteLine("PASS: " + passed + " text pointer tests");
    }
    private static bool Native(int type,bool editing,float x,float y)
    { return TextPointerPolicy.UseNativeTextInput(type,editing,Box,new PointF(x,y),6); }
    private static bool Move(bool selected,bool? nativeHitMatches,int type,bool editing,RectangleF bounds,PointF point,float border)
    { return TextPointerPolicy.HitMoveBorder(type,editing,bounds,point,border) && TextPointerPolicy.CanUseBorderHit(selected,nativeHitMatches); }
    private static void Check(bool result,string description)
    { if(!result)throw new Exception(description);passed++;Console.WriteLine("PASS "+description); }
}
