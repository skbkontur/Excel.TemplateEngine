using JetBrains.Annotations;

namespace Excel.TemplateEngine.FileGenerating.Interfaces
{
    public interface IExcelDropDownControlInfo : IExcelFormControlInfo
    {
        [CanBeNull]
        string SelectedValue { get; set; }
    }
}