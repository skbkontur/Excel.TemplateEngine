namespace Benchmark.IhclmeModel
{
    public class PartyInfoExcelModel
    {
        public string Name { get; set; }
        public string Inn { get; set; }
        public string Kpp { get; set; }
        public string MediId { get; set; }
        public string FullAddress { get; set; }

        public BankAccountExcelModel BankAccount { get; set; }

        public ContactInformationExcelModel Chief { get; set; }
        public ContactInformationExcelModel Accountant { get; set; }
    }
}