namespace WaybillsAPI.CreationModels
{
    public class WaybillCreation
    {
        public int Id { get; set; }

        public int Number { get; set; }
        public DateOnly Date { get; set; }
        public int Days { get; set; }
        public double Hours { get; set; }

        public int StartFuel { get; set; }
        public int FuelTopUp { get; set; }
        public int EndFuel { get; set; }

        public int DriverId { get; set; }
        public int TransportId { get; set; }

        public double Bonus { get; set; }
        public double Weekend { get; set; }

        public double BonusSizeInPercentages { get; set; }
        public bool WeekendEqualsEarnings { get; set; }

        public List<OperationCreation> Operations { get; set; } = [];
        public List<CalculationCreation> Calculations { get; set; } = [];
    }
}
