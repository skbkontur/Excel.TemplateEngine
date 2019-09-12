namespace Excel.TemplateEngine.ObjectPrinting.RenderingTemplates
{
    internal interface ITemplateCollection
    {
        RenderingTemplate GetTemplate(string templateName);
    }
}