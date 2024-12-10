using Microsoft.EntityFrameworkCore;
using WaybillsAPI.Context;
using WaybillsAPI.CreationModels;
using WaybillsAPI.Interfaces;

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
        public double Bonus { get; private set; }
        public double Weekend { get; private set; }

        public List<Operation> Operations { get; private set; } = [];
        public List<Calculation> Calculations { get; private set; } = [];

        private Waybill() { }
        public Waybill(WaybillsContext context, ISalaryPeriodService salaryPeriodService, WaybillCreation creationModel)
        {
            var (Year, Month) = salaryPeriodService.GetSalaryPeriod();
            if (context.Waybills.Any(x => x.Number == creationModel.Number && x.SalaryYear == Year))
            {
                throw new ArgumentException($"Путевой лист №{creationModel.Number} уже существует в {Year} году.");
            }

            SalaryYear = Year;
            SalaryMonth = Month;
            SetProperties(context, salaryPeriodService, creationModel);
        }

        public void Edit(WaybillsContext context, ISalaryPeriodService salaryPeriodService, WaybillCreation creationModel)
        {
            if (context.Waybills.Any(x => x.Id != Id && x.Number == creationModel.Number && x.SalaryYear == SalaryYear))
            {
                throw new ArgumentException($"Путевой лист №{creationModel.Number} уже существует в {SalaryYear} году.");
            }

            SetProperties(context, salaryPeriodService, creationModel);
        }

        private void SetProperties(WaybillsContext context, ISalaryPeriodService salaryPeriodService, WaybillCreation creationModel)
        {
            var maxWaybillDate = salaryPeriodService.GetMaxWaybillDate();
            Date = creationModel.Date <= maxWaybillDate ? creationModel.Date : throw new ArgumentException("Дата путевого листа больше допустимого.");
            Driver = context.Drivers.Find(creationModel.DriverId) ?? throw new ArgumentException("Водителя с заданным ID не существует.");
            Transport = context.Transport.Find(creationModel.TransportId) ?? throw new ArgumentException("Транспорта с заданным ID не существует.");

            Number = creationModel.Number;
            Days = creationModel.Days;
            Hours = creationModel.Hours;
            StartFuel = creationModel.StartFuel;
            FuelTopUp = creationModel.FuelTopUp;
            EndFuel = creationModel.EndFuel;
            DriverId = creationModel.DriverId;
            TransportId = creationModel.TransportId;

            Operations = creationModel.Operations.Select(x => new Operation(x, Transport.Coefficient)).ToList();
            Calculations = creationModel.Calculations.Select(x => new Calculation(x)).ToList();

            FactFuelConsumption = StartFuel + FuelTopUp - EndFuel;
            NormalFuelConsumption = (int)Math.Round(Operations.Sum(x => x.TotalFuelConsumption));
            ConditionalReferenceHectares = Math.Round(Operations.Sum(x => x.ConditionalReferenceHectares), 2);
            Earnings = Math.Round(Calculations.Sum(x => x.Sum), 2);
            Bonus = creationModel.BonusSizeInPercentages > 0 ? Math.Round(Earnings * creationModel.BonusSizeInPercentages / 100, 2) : creationModel.Bonus;
            Weekend = creationModel.WeekendEqualsEarnings ? Earnings : creationModel.Weekend;
        }

        public string FullDate => Days == 2 ? Date.ToString($"d—{Date.Day + 1} MMMM yyyy") : Date.ToString($"d MMMM yyyy");
        public string DriverShortFullName => Driver is null ? "" : Driver.ShortFullName();
        public string TransportName => Transport is null ? "" : Transport.Name;
    }
}
