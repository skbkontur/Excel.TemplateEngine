namespace Benchmark.IhclmeModel
{
    public class LineItemExcelModel
    {
        public string Name { get; set; }
        public string PayerServiceCode { get; set; }
        public string SupplierServiceCode { get; set; }
        public string PriceListCode { get; set; }
        public string ServiceDate { get; set; }
        public string ServiceEndDate { get; set; }
        public string AdmissionDate { get; set; }
        public string DischargeDate { get; set; }
        public string HealthCareResponseNumber { get; set; }
        public string HealthCareResponseDate { get; set; }
        public string ServiceCategory { get; set; }
        public string InsuranceServiceStatus { get; set; }
        public string OmsCode { get; set; }
        public string Comment { get; set; }
        public string ReportId { get; set; }

        public DentalInformationExcelModel DentalInformation { get; set; }
        public string ReasonToExclude { get; set; }
        public decimal ConfirmedQuantity { get; set; }
        public decimal ProvidedQuantity { get; set; }
        public decimal Price { get; set; }
        public string AllowanceCode { get; set; }
        public decimal TotalSumModifier { get; set; }
        public decimal ContractDiscount { get; set; }
        public decimal TotalWithoutDiscountCharge { get; set; }
        public decimal Total { get; set; }
        public decimal TotalInCurrency { get; set; }
        public decimal TotalExcluded { get; set; }
        public decimal TotalExcludedInCurrency { get; set; }
        public decimal TotalPaidByCustomer { get; set; }
    }
}