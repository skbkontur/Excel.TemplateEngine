namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelDropDownControlInfo : IExcelFormControlInfo
    {
        string SelectedValue { get; set; }
    }
}