namespace Benchmark.IhclmeModel
{
    class IhclmeExcelModel
    {
        public PartyInfoExcelModel HealthCareProvider { get; set; }

        public PartyInfoExcelModel InsuranceCompany { get; set; }

        public string ContractNumber { get; set; }
        public string ContractDate { get; set; }
        public string HealthCareClaimNumber { get; set; }
        public string HealthCareClaimDate { get; set; }
        public string ClaimPeriodFrom { get; set; }
        public string ClaimPeriodTo { get; set; }

        public string SupplementaryAgreementNumber { get; set; }
        public string SupplementaryAgreementDate { get; set; }

        public string AcceptanceCertificateNumber { get; set; }
        public string AcceptanceCertificateDate { get; set; }

        public string ProformaInvoiceNumber { get; set; }
        public string ProformaInvoiceDate { get; set; }

        public string InvoiceNumber { get; set; }
        public string InvoiceDate { get; set; }

        public string ProvidedServicePaymentType { get; set; }

        public string CurrencyCode { get; set; }
        public decimal? CurrencyExchangeRate { get; set; }

        public decimal? ClaimTotalSum { get; set; }
        public decimal? ClaimTotalVat { get; set; }
        public decimal? ClaimTotalSumWithVat { get; set; }
        public decimal? ClaimTotalWithoutDiscountCharge { get; set; }
        public decimal? ClaimTotalDiscount { get; set; }
        public decimal? ClaimTotalSumWithVatInCurrency { get; set; }

        public VATDetailsExcelModel VATDetailsWithoutVAT { get; set; }
        public VATDetailsExcelModel VATDetails20VAT { get; set; }

        public string ProvidedServicesName { get; set; }
        public ProvidedServiceExcelModel[] ProvidedServices { get; set; }
        public ProformaInvoiceSumInfoExcelModel ProformaInvoiceSumInfo { get; set; }

        public AcceptanceCertificateInfoExcelModel AcceptanceCertificateInfo { get; set; }
    }
}