namespace Excel.TemplateEngine.FileGenerating.Primitives
{
    public interface IExcelCheckBoxControlInfo : IExcelFormControlInfo
    {
        bool IsChecked { get; set; }
    }
}