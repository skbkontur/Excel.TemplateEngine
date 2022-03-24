namespace Benchmark.IhclmeModel
{
    public class ProvidedServiceExcelModel
    {
        public int NumberInList { get; set; }
        public string PriceListCode { get; set; }

        public InsuredPersonExcelModel InsuredPerson { get; set; }
        public TreatmentExcelModel Treatment { get; set; }
        public PartyInfoExcelModel HealthCareProviderDepartment { get; set; }
    }
}