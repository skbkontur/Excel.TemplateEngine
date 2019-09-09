namespace Excel.TemplateEngine.FileGenerating.Interfaces
{
    public interface IExcelCheckBoxControlInfo : IExcelFormControlInfo
    {
        bool IsChecked { get; set; }
    }
}