using JetBrains.Annotations;

namespace SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes
{
    public class ExcelCellNumberingFormat
    {
        public ExcelCellNumberingFormat(uint id, [CanBeNull] string code = null)
        {
            Id = id;
            Code = code;
        }

        public override string ToString() => $"Id = {Id}, Code = {Code}";

        public uint Id { get; }

        [CanBeNull]
        public string Code { get; }
    }
}