namespace SKBKontur.Catalogue.ExcelFileGenerator.DataTypes
{
    public class ExcelCellStyle
    {
        public ExcelCellNumberingFormat NumberingFormat { get; set; }
        public ExcelCellFillStyle FillStyle { get; set; }
        public ExcelCellFontStyle FontStyle { get; set; }
        public ExcelCellBordersStyle BordersStyle { get; set; }
        public ExcelCellAlignment Alignment { get; set; }
    }
}