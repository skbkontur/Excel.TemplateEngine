using JetBrains.Annotations;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.Primitives
{
    public interface IExcelDropDownControlInfo : IExcelFormControlInfo
    {
        [CanBeNull]
        string SelectedValue { get; set; }
    }
}