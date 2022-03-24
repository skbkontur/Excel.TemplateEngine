namespace Benchmark.IhclmeModel
{
    public class BankAccountExcelModel
    {
        public string BankAccountNumber { get; set; }
        public string CorrespondentAccountNumber { get; set; }
        public string BankId { get; set; }
        public string BankName { get; set; }
        public string RecipientInn { get; set; }
        public string RecipientKpp { get; set; }
        public string RecipientName { get; set; }
        public string RecipientId { get; set; }

        // kbk is used to be printed only in proforma invoice
        public string Kbk { get; set; }
    }
}