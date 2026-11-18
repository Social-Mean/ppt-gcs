using System;
using PptGcs.AddIn;

public static class NativeFieldDirtyTests
{
    private static int count;
    public static int Main()
    {
        try
        {
            // Native baseline: 140pt displays 4.94cm; focus and leave retains
            // exactly140pt. The unchanged display must not become140.031494pt.
            Scenario("4.94 厘米","4.94 厘米",0,false);
            // Native baseline: retyping the displayed number sets raw=-1 and
            // intentionally commits the rounded value, even if texts match.
            Scenario("4.94 厘米","4.94 厘米",-1,true);
            Scenario("4.94 厘米","4.94",-1,true);
            Scenario("4.94 厘米","4.94 厘米",1,true);
            Scenario("4.94 厘米","6 cm",-1,true);
            // Formatting can change the display after a native commit while its
            // modify flag is clear. Text differences do not override the flag.
            Scenario("6 cm","60 mm",0,false);
            Scenario("","8",-1,true);
            Scenario("","",0,false);
            Unknown("4.94 厘米","4.94 厘米");
            Unknown("4.94 厘米","8");
            Console.WriteLine("Native field dirty tests passed: "+count);return 0;
        }
        catch(Exception error){Console.Error.WriteLine(error);return 1;}
    }
    private static void Scenario(string initial,string current,long raw,bool expected)
    {
        if(NativeFieldDirtyPolicy.RequiresPreapply(raw)!=expected)
            throw new Exception("Wrong edit decision for raw="+raw+", initial="+initial+", current="+current);
        count++;
    }
    private static void Unknown(string initial,string current)
    {
        try{NativeFieldDirtyPolicy.RequiresPreapply(null);}
        catch(InvalidOperationException){count++;return;}
        throw new Exception("Unknown flag was treated as unmodified: "+initial+" -> "+current);
    }
}
