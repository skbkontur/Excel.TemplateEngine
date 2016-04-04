namespace SKBKontur.Catalogue.ExcelFileGenerator.DataTypes
{
    public class ExcelColor
    {
        public int Red { get; set; }
        public int Green { get; set; }
        public int Blue { get; set; }
        public int Alpha { get; set; }

        public override string ToString()
        {
            return string.Format("R = {0}, G = {1}, B = {2}, A = {3}", Red, Green, Blue, Alpha);
        }
    }
}