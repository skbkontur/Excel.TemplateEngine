namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelCheckBoxControlInfo : IExcelFormControlInfo
    {
        bool IsChecked { get; set; }
    }
}