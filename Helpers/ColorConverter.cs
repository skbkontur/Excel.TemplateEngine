using System;
using System.Diagnostics.CodeAnalysis;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Helpers
{
    [SuppressMessage("ReSharper", "CompareOfFloatsByEqualityOperator")]
    public static class ColorConverter
    {
        public struct HslColor
        {
            public double A { get; set; }
            public double H { get; set; }
            public double L { get; set; }
            public double S { get; set; }
        }

        public static HslColor RgbToHls(ExcelColor rgbColor)
        {
            var hslColor = new HslColor();
            var r = (double)rgbColor.Red / 255;
            var g = (double)rgbColor.Green / 255;
            var b = (double)rgbColor.Blue / 255;
            var a = (double)rgbColor.Alpha / 255;
            var min = Math.Min(r, Math.Min(g, b));
            var max = Math.Max(r, Math.Max(g, b));
            var delta = max - min;
            if(max == min)
            {
                hslColor.A = a;
                hslColor.H = 0;
                hslColor.S = 0;
                hslColor.L = max;
                return hslColor;
            }
            hslColor.L = (min + max) / 2;
            if(hslColor.L < 0.5)
            {
                hslColor.S = delta / (max + min);
            }
            else
            {
                hslColor.S = delta / (2.0 - max - min);
            }
            if(r == max) hslColor.H = (g - b) / delta;
            if(g == max) hslColor.H = 2.0 + (b - r) / delta;
            if(b == max) hslColor.H = 4.0 + (r - g) / delta;
            hslColor.H *= 60;
            if(hslColor.H < 0) hslColor.H += 360;
            hslColor.A = a;
            return hslColor;
        }

        public static ExcelColor HslToRgb(HslColor hslColor)
        {
            if(hslColor.S == 0)
            {
                return new ExcelColor
                    {
                        Alpha = (int)(hslColor.A * 255),
                        Red = (int)(hslColor.L * 255),
                        Green = (int)(hslColor.L * 255),
                        Blue = (int)(hslColor.L * 255),
                    };
            }
            double t1;
            if(hslColor.L < 0.5)
            {
                t1 = hslColor.L * (1.0 + hslColor.S);
            }
            else
            {
                t1 = hslColor.L + hslColor.S - (hslColor.L * hslColor.S);
            }
            var t2 = 2.0 * hslColor.L - t1;
            var h = hslColor.H / 360;
            var tR = h + (1.0 / 3.0);
            var r = SetColor(t1, t2, tR);
            var tG = h;
            var g = SetColor(t1, t2, tG);
            var tB = h - (1.0 / 3.0);
            var b = SetColor(t1, t2, tB);
            return new ExcelColor
                {
                    Alpha = (int)(hslColor.A * 255),
                    Red = (int)(r * 255),
                    Green = (int)(g * 255),
                    Blue = (int)(b * 255),
                };
        }

        private static double SetColor(double t1, double t2, double t3)
        {
            if(t3 < 0) t3 += 1.0;
            if(t3 > 1) t3 -= 1.0;
            double color;
            if(6.0 * t3 < 1)
            {
                color = t2 + (t1 - t2) * 6.0 * t3;
            }
            else if(2.0 * t3 < 1)
            {
                color = t1;
            }
            else if(3.0 * t3 < 2)
            {
                color = t2 + (t1 - t2) * ((2.0 / 3.0) - t3) * 6.0;
            }
            else
            {
                color = t2;
            }
            return color;
        }
    }
}