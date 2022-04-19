namespace Benchmark.IhclmeModel
{
    public sealed class ProformaInvoiceSumInfoExcelModel
    {
        public decimal? TotalSum { get; set; }
        public string PaymentPurpose { get; set; }

        public TotalModifierExcelModel Modifier1 { get; set; }
        public TotalModifierExcelModel Modifier2 { get; set; }
        public TotalModifierExcelModel Modifier3 { get; set; }
    }
}