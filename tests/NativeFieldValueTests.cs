using System;
using System.Globalization;
using PptGcs.AddIn;

public static class NativeFieldValueTests
{
    private static int count;
    public static int Main()
    {
        try
        {
            var english=CultureInfo.GetCultureInfo("en-US");var german=CultureInfo.GetCultureInfo("de-DE");
            Equal("cm",NativeFieldValue.CaptureUnit("8 厘米","in"));
            Equal("mm",NativeFieldValue.CaptureUnit("80 mm","in"));
            Equal("in",NativeFieldValue.CaptureUnit("3","in"));
            Equal("cm",NativeFieldValue.CaptureUnit("","cm"));
            Equal("in",NativeFieldValue.CaptureUnit(" ","in"));
            Equal(null,NativeFieldValue.CaptureUnit("混合","cm"));
            Parses("2.54 cm",null,english,72);
            Parses("25.4毫米",null,english,72);
            Parses("1英寸",null,english,72);
            Parses("1\"",null,english,72);
            Parses("72磅",null,english,72);
            Parses("72 points",null,english,72);
            Parses("2.54","cm",english,72);
            Parses("8",NativeFieldValue.CaptureUnit("","cm"),english,8*72/2.54);
            Parses("2",NativeFieldValue.CaptureUnit("","in"),english,144);
            Parses("2,54 cm",null,german,72);
            Parses("2.54 cm",null,german,72);
            Parses("1e2 pt",null,english,100);
            Rejects("8",null,english);
            Rejects("8px",null,english);
            Rejects("NaN pt",null,english);
            Rejects("Infinity in",null,english);
            Rejects("1e40 pt",null,english);
            Rejects("0 pt",null,english);
            Rejects("-1 cm",null,english);
            Rejects("0.999 pt",null,english);
            Rejects("1,000 pt",null,english);
            Parses("1 pt",null,english,1);
            string restored=NativeFieldValue.FormatPoints(226.77165222167969,"cm",english);
            Parses(restored,null,english,226.77165222167969);
            double coordinate;
            if(!NativeFieldValue.TryParsePoints("-2.54 cm",null,english,Double.NegativeInfinity,out coordinate)||Math.Abs(coordinate+72)>0.0000001)
                throw new Exception("Negative position was rejected or changed.");count++;
            if(!NativeFieldValue.TryParsePoints("0 pt",null,english,Double.NegativeInfinity,out coordinate)||coordinate!=0)
                throw new Exception("Zero position was rejected.");count++;
            if(NativeFieldValue.TryParsePoints("-Infinity pt",null,english,Double.NegativeInfinity,out coordinate))
                throw new Exception("Infinite position was accepted.");count++;
            if(NativeFieldValue.TryParsePoints("-1e40 pt",null,english,Double.NegativeInfinity,out coordinate))
                throw new Exception("Position outside Office Single range was accepted.");count++;
            Console.WriteLine("Native field value tests passed: "+count);return 0;
        }
        catch(Exception error){Console.Error.WriteLine(error);return 1;}
    }
    private static void Equal(string expected,string actual)
    {if(expected!=actual)throw new Exception("Expected unit "+expected+", got "+actual);count++;}
    private static void Parses(string text,string unit,CultureInfo culture,double expected)
    {double value;if(!NativeFieldValue.TryParsePoints(text,unit,culture,1,out value)||Math.Abs(value-expected)>0.0000001)throw new Exception("Unexpected points for "+text+": "+value);count++;}
    private static void Rejects(string text,string unit,CultureInfo culture)
    {double value;if(NativeFieldValue.TryParsePoints(text,unit,culture,1,out value))throw new Exception("Accepted invalid or ambiguous field: "+text);count++;}
}
