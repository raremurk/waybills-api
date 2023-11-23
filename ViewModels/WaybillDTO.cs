namespace WaybillsAPI.ViewModels
{
    public class WaybillDTO
    {
        public int Id { get; set; }

        public int Number { get; set; }
        public DateOnly Date { get; set; }
        public bool TwoDaysWaybill { get; set; }
        public int Days { get; set; }
        public double Hours { get; set; }

        public int StartFuel { get; set; }
        public int FuelTopUp { get; set; }
        public int EndFuel { get; set; }
        public int FactFuelConsumption { get; set; }
        public int NormalFuelConsumption { get; set; }

        public double ConditionalReferenceHectares { get; set; }

        public DriverDTO? Driver { get; set; }
        public TransportDTO? Transport { get; set; }

        public double Earnings { get; set; }
        public double Weekend { get; set; }
        public double Bonus { get; set; }

        public List<OperationDTO> Operations { get; set; }
        public List<CalculationDTO> Calculations { get; set; }
    }
}
