using Microsoft.EntityFrameworkCore;
using WaybillsAPI.CreationModels;

namespace WaybillsAPI.Models
{
    [Index(nameof(Number), nameof(SalaryYear), IsUnique = true)]
    public class Waybill
    {
        public int Id { get; private set; }

        public int Number { get; private set; }

        public int SalaryYear { get; private set; }
        public int SalaryMonth { get; private set; }
        public DateOnly Date { get; private set; }

        public int Days { get; private set; }
        public double Hours { get; private set; }

        public int StartFuel { get; private set; }
        public int FuelTopUp { get; private set; }
        public int EndFuel { get; private set; }
        public int FactFuelConsumption { get; private set; }
        public int NormalFuelConsumption { get; private set; }

        public double ConditionalReferenceHectares { get; private set; }

        public int DriverId { get; private set; }
        public Driver? Driver { get; private set; }

        public int TransportId { get; private set; }
        public Transport? Transport { get; private set; }

        public double Earnings { get; private set; }
        public double Weekend { get; private set; }
        public double Bonus { get; private set; }

        public List<Operation> Operations { get; private set; } = [];
        public List<Calculation> Calculations { get; private set; } = [];

        private Waybill() { }
        public Waybill(WaybillCreation creationModel, DateOnly currentDate, double transportCoefficient)
        {
            Id = creationModel.Id;
            Number = creationModel.Number;
            SetSalaryDate(currentDate);
            Date = creationModel.Date;
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

        public string FullDate => Days == 2 ? Date.ToString($"d—{Date.Day + 1} MMMM yyyy") : Date.ToString($"d MMMM yyyy");
        public string DriverShortFullName => Driver is null ? "" : Driver.ShortFullName();
        public string TransportName => Transport is null ? "" : Transport.Name;

        private void SetSalaryDate(DateOnly currentDate)
        {
            if (currentDate.Day < 16)
            {
                currentDate = currentDate.AddMonths(-1);
            }

            SalaryMonth = currentDate.Month;
            SalaryYear = currentDate.Year;
        }
    }
}
