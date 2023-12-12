namespace WaybillsAPI.ViewModels
{
    public class ShortWaybillDTO
    {
        public int Id { get; set; }

        public int Number { get; set; }
        public DateOnly Date { get; set; }
        public int Days { get; set; }
        public double Hours { get; set; }

        public int FactFuelConsumption { get; set; }
        public int NormalFuelConsumption { get; set; }

        public double ConditionalReferenceHectares { get; set; }

        public string DriverShortFullName { get; set; }
        public string TransportName { get; set; }

        public double Earnings { get; set; }
        public double Weekend { get; set; }
        public double Bonus { get; set; }
    }
}
