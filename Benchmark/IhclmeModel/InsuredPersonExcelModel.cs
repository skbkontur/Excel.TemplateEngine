namespace Benchmark.IhclmeModel
{
    public class InsuredPersonExcelModel
    {
        public int? NumberInList { get; set; }

        public string FirstName { get; set; }
        public string MiddleName { get; set; }
        public string LastName { get; set; }
        public string FullName { get; set; }

        public string Gender { get; set; }
        public string MedicalRecord { get; set; }
        public string DateOfBirth { get; set; }

        public PartyInfoExcelModel Employer { get; set; }

        public InsurancePolicyExcelModel InsurancePolicy { get; set; }
        public InsuranceProgramExcelModel InsuranceProgram { get; set; }
        public ContactsExcelModel Contacts { get; set; }
        public string FullAddress { get; set; }

        public string InsurantStatus { get; set; }
    }
}