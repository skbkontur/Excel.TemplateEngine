namespace SKBKontur.Catalogue.ExcelFileGenerator
{
    public class ExcelCellStyle
    {
        public ExcelCellNumberingFormat NumberingFormat { get; set; }
        public ExcelCellFillStyle FillStyle { get; set; }
        public ExcelCellFontStyle FontStyle { get; set; }
        public ExcelCellBordersStyle BordersStyle { get; set; }
        public ExcelCellAlignment Alignment { get; set; }
    }

    public class ExcelCellFontStyle
    {
        public int? Size { get; set; }
        public ExcelColor Color { get; set; }
    }

    public class ExcelCellAlignment
    {
        public ExcelHorizontalAlignment HorizontalAlignment { get; set; }
        public ExcelVerticalAlignment VerticalAlignment { get; set; }
    }

    public enum ExcelVerticalAlignment
    {
        Default,
        Top,
        Center,
        Bottom
    }

    public enum ExcelHorizontalAlignment
    {
        Default,
        Left,
        Center,
        Right
    }

    public class ExcelCellBordersStyle
    {
        public ExcelCellBorderStyle LeftBorder { get; set; }
        public ExcelCellBorderStyle RightBorder { get; set; }
        public ExcelCellBorderStyle TopBorder { get; set; }
        public ExcelCellBorderStyle BottomBorder { get; set; }
    }

    public class ExcelCellBorderStyle
    {
        public ExcelColor Color { get; set; }
        public ExcelBorderType BorderType { get; set; }
    }

    public enum ExcelBorderType
    {
        None,
        Single,
        Double,
        Bold
    }

    public class ExcelCellFillStyle
    {
        public ExcelColor Color { get; set; }
    }

    public class ExcelColor
    {
        public int Red { get; set; }
        public int Green { get; set; }
        public int Blue { get; set; }
        public int Alpha { get; set; }
    }

    public static class ExcelColors
    {
        public static ExcelColor Black = new ExcelColor {Red = 0, Green = 0, Blue = 0, Alpha = 0};
        public static ExcelColor White = new ExcelColor {Red = 255, Green = 255, Blue = 255, Alpha = 0};
    }

    public class ExcelCellNumberingFormat
    {
        public int Precision { get; set; }
    }
}