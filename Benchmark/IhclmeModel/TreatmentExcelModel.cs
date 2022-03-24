namespace Benchmark.IhclmeModel
{
    public class TreatmentExcelModel
    {
        public DiagnosisExcelModel Diagnosis { get; set; }
        public string DoctorFullName { get; set; }
        public string DoctorCode { get; set; }
        public string DoctorSpecialization { get; set; }
        public LineItemExcelModel LineItem { get; set; }
    }
}