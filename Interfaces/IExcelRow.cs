namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelRow
    {
        IExcelCell CreateCell(int index);
    }
}