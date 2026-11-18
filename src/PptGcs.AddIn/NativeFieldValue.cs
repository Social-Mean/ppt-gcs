using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace PptGcs.AddIn
{
    // PowerPoint fields retain the user's display unit when an edit omits one.
    // Native mixed blank fields use the user's regional measurement unit.
    public static class NativeFieldValue
    {
        private static readonly Regex Literal=new Regex(@"^\s*(?<number>[+-]?(?:\d+(?:[\.,]\d*)?|[\.,]\d+)(?:[eE][+-]?\d+)?)\s*(?<unit>[^\d\s].*?)?\s*$",RegexOptions.CultureInvariant);
        public static string CaptureUnit(string displayed,string regionalUnit)
        {
            if(String.IsNullOrWhiteSpace(displayed))return NormalizeUnit(regionalUnit);
            Match match=Literal.Match(displayed);
            if(!match.Success)return null;
            string suffix=match.Groups["unit"].Value.Trim();
            return suffix.Length==0 ? NormalizeUnit(regionalUnit) : NormalizeUnit(suffix);
        }
        public static bool TryParsePoints(string text,string bareUnit,CultureInfo culture,double minimumPoints,out double points)
        {
            points=0;Match match=Literal.Match(text??String.Empty);if(!match.Success)return false;
            string suffix=match.Groups["unit"].Value.Trim();string unit=NormalizeUnit(suffix.Length==0?bareUnit:suffix);
            if(unit==null)return false;
            string number=match.Groups["number"].Value;double value;
            if(!Double.TryParse(number,NumberStyles.Float,culture??CultureInfo.CurrentCulture,out value) &&
                !Double.TryParse(number,NumberStyles.Float,CultureInfo.InvariantCulture,out value))return false;
            points=value*PointsPerUnit(unit);
            return !Double.IsNaN(points)&&!Double.IsInfinity(points)&&points>=minimumPoints&&Math.Abs(points)<=Single.MaxValue;
        }
        public static string FormatPoints(double points,string unit,CultureInfo culture)
        {
            string normalized=NormalizeUnit(unit);if(normalized==null)throw new ArgumentException("A known native display unit is required.","unit");
            return (points/PointsPerUnit(normalized)).ToString("0.###############",culture??CultureInfo.CurrentCulture)+" "+normalized;
        }
        public static string NormalizeUnit(string unit)
        {
            switch((unit??String.Empty).Trim().ToLowerInvariant())
            {
                case "cm":case "centimeter":case "centimeters":case "centimetre":case "centimetres":case "厘米":return "cm";
                case "mm":case "millimeter":case "millimeters":case "millimetre":case "millimetres":case "毫米":return "mm";
                case "in":case "inch":case "inches":case "\"":case "英寸":case "吋":return "in";
                case "pt":case "point":case "points":case "磅":case "点":return "pt";
                default:return null;
            }
        }
        private static double PointsPerUnit(string unit)
        {return unit=="cm"?72.0/2.54:unit=="mm"?72.0/25.4:unit=="in"?72.0:1.0;}
    }
}
