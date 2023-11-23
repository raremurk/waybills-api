using WaybillsAPI.CreationModels;

namespace WaybillsAPI.Models
{
    public class Operation
    {
        public int Id { get; init; }

        public string ProductionCostCode { get; init; }

        public int NumberOfRuns { get; init; }
        public double TotalMileage { get; init; }
        public double TotalMileageWithLoad { get; init; }
        public double TransportedLoad { get; init; }

        public double Norm { get; init; }
        public double Fact { get; init; }
        public double MileageWithLoad { get; init; }

        public double NormShift { get; init; }
        public double ConditionalReferenceHectares { get; init; }

        public double FuelConsumptionPerUnit { get; init; }
        public double TotalFuelConsumption { get; init; }

        public int WaybillId { get; init; }
        public Waybill? Waybill { get; init; }

        private Operation() { }
        public Operation(OperationCreation operationCreation, double transportCoefficient)
        {
            Id = operationCreation.Id;
            ProductionCostCode = operationCreation.ProductionCostCode;
            NumberOfRuns = operationCreation.NumberOfRuns;
            TotalMileage = operationCreation.TotalMileage;
            TotalMileageWithLoad = operationCreation.TotalMileageWithLoad;
            TransportedLoad = operationCreation.TransportedLoad;
            Norm = operationCreation.Norm;
            Fact = operationCreation.Fact;
            MileageWithLoad = NumberOfRuns != 0 ? Math.Round(TotalMileageWithLoad / NumberOfRuns, 1) : 0;
            NormShift = Norm != 0 ? Math.Round(Fact / Norm, 2) : 0;
            ConditionalReferenceHectares = Math.Round(NormShift * transportCoefficient, 2);
            FuelConsumptionPerUnit = operationCreation.FuelConsumptionPerUnit;
            TotalFuelConsumption = Math.Round(FuelConsumptionPerUnit * Fact, 1);
        }
    }
}
