using WaybillsAPI.CreationModels;

namespace WaybillsAPI.Models
{
    public class Waybill
    {
        public int Id { get; init; }

        public int Number { get; init; }
        public DateOnly Date { get; init; }
        public bool TwoDaysWaybill { get; init; }
        public int Days { get; init; }
        public double Hours { get; init; }

        public int StartFuel { get; init; }
        public int FuelTopUp { get; init; }
        public int EndFuel { get; init; }
        public int FactFuelConsumption { get; init; }
        public int NormalFuelConsumption { get; init; }

        public double ConditionalReferenceHectares { get; init; }

        public int DriverId { get; init; }
        public Driver? Driver { get; init; }

        public int TransportId { get; init; }
        public Transport? Transport { get; init; }

        public double Earnings { get; init; }
        public double Weekend { get; init; }
        public double Bonus { get; init; }

        public List<Operation> Operations { get; init; } = [];
        public List<Calculation> Calculations { get; init; } = [];

        private Waybill() { }
        public Waybill(WaybillCreation creationModel, double transportCoefficient)
        {
            Id = creationModel.Id;
            Number = creationModel.Number;
            Date = creationModel.Date;
            TwoDaysWaybill = creationModel.TwoDaysWaybill;
            Days = creationModel.Days;
            Hours = creationModel.Hours;
            StartFuel = creationModel.StartFuel;
            FuelTopUp = creationModel.FuelTopUp;
            EndFuel = creationModel.EndFuel;
            DriverId = creationModel.DriverId;
            TransportId = creationModel.TransportId;
            Weekend = creationModel.Weekend;
            Bonus = creationModel.Bonus;

            Operations = creationModel.Operations.Select(x => new Operation(x, transportCoefficient)).ToList();
            Calculations = creationModel.Calculations.Select(x => new Calculation(x)).ToList();

            FactFuelConsumption = StartFuel + FuelTopUp - EndFuel;
            NormalFuelConsumption = (int)Math.Round(Operations.Sum(x => x.TotalFuelConsumption));
            ConditionalReferenceHectares = Math.Round(Operations.Sum(x => x.ConditionalReferenceHectares), 2);
            Earnings = Math.Round(Calculations.Sum(x => x.Sum), 2);
        }

        public string FullDate => TwoDaysWaybill ? Date.ToString($"d—{Date.Day + 1} MMMM yyyy") : Date.ToString($"d MMMM yyyy");
        public string DriverShortFullName => Driver is null ? "" : Driver.ShortFullName();
        public string TransportName => Transport is null ? "" : Transport.Name;
    }
}
