using System;
using System.Drawing;
using PptGcs.AddIn;

internal static class TextPointerTests
{
    private static int passed;
    private static readonly RectangleF Box = new RectangleF(100, 100, 200, 80);
    private static void Main()
    {
        Check(Native(17,false,false,160,130), "first click in a text box remains native");
        Check(Native(17,true,false,160,130), "selected text box accepts the next native click");
        Check(Native(17,true,true,160,130), "caret and text selection remain native");
        Check(Native(17,false,false,101,101), "unselected text-box edge selects natively first");
        Check(!Native(17,true,false,102,130), "selected left border moves with constraints");
        Check(!Native(17,true,true,298,130), "right border moves with constraints while editing text");
        Check(!Native(17,true,true,180,102), "top border moves with constraints while editing text");
        Check(!Native(17,true,true,180,178), "bottom border moves with constraints while editing text");
        Check(!Native(17,true,true,80,130), "outside clicks are not claimed as text input");
        Check(!Native(1,true,false,160,130), "ordinary shape-body drag remains constrained");
        Check(Native(1,true,true,160,130), "already editing shape text retains native caret handling");
        Check(!Native(13,true,false,160,130), "picture-body drag remains constrained");
        Check(TextPointerPolicy.UseNativeTextInput(17,true,false,new RectangleF(200,200,400,160),new PointF(213,250),12), "text hit testing follows device-pixel scale");
        Check(!TextPointerPolicy.UseNativeTextInput(17,true,true,new RectangleF(200,200,400,160),new PointF(210,250),12), "scaled border preserves constrained movement");
        Console.WriteLine("PASS: " + passed + " text pointer tests");
    }
    private static bool Native(int type,bool selected,bool editing,float x,float y)
    { return TextPointerPolicy.UseNativeTextInput(type,selected,editing,Box,new PointF(x,y),6); }
    private static void Check(bool result,string description)
    { if(!result)throw new Exception(description);passed++;Console.WriteLine("PASS "+description); }
}
