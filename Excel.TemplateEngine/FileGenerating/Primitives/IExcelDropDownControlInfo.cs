using JetBrains.Annotations;

namespace Excel.TemplateEngine.FileGenerating.Primitives
{
    public interface IExcelDropDownControlInfo : IExcelFormControlInfo
    {
        [CanBeNull]
        string SelectedValue { get; set; }
    }
}