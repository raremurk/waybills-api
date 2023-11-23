namespace WaybillsAPI.CreationModels
{
    public class OperationCreation
    {
        public int Id { get; set; }
        public string ProductionCostCode { get; set; }
        public int NumberOfRuns { get; set; }
        public double TotalMileage { get; set; }
        public double TotalMileageWithLoad { get; set; }
        public double TransportedLoad { get; set; }
        public double Norm { get; set; }
        public double Fact { get; set; }
        public double FuelConsumptionPerUnit { get; set; }
    }
}
