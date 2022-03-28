namespace Benchmark.IhclmeModel
{
    public sealed class AcceptanceCertificateInfoExcelModel
    {
        public decimal? OrganizationReward { get; set; }
        public decimal? Total { get; set; }
        public decimal? TotalWithVat { get; set; }
        public decimal? TotalVat { get; set; }
        public string Comment { get; set; }
        public TotalModifierExcelModel Modifier1 { get; set; }
        public TotalModifierExcelModel Modifier2 { get; set; }
        public TotalModifierExcelModel Modifier3 { get; set; }
    }
}