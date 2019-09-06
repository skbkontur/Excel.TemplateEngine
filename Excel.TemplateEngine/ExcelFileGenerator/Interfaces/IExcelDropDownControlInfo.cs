using JetBrains.Annotations;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelDropDownControlInfo : IExcelFormControlInfo
    {
        [CanBeNull]
        string SelectedValue { get; set; }
    }
}