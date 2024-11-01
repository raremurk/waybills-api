using WaybillsAPI.Models;
using WaybillsAPI.ViewModels;

namespace WaybillsAPI.ReportsModels.Fuel.Drivers
{
    public class DetailedDriverFuelMonthTotal : DriverFuelMonthTotal
    {
        public DriverDTO Driver { get; init; }
        public IEnumerable<DriverFuelSubTotal> SubTotals { get; init; }

        public DetailedDriverFuelMonthTotal(IEnumerable<Waybill> waybills)
        {
            if (waybills is null || !waybills.Any()) throw new ArgumentException("The collection of waybills is null or empty.");

            var driver = waybills.First().Driver ?? throw new ArgumentException("Waybills must include driver.");

            if (waybills.Any(x => x.DriverId != driver.Id)) throw new ArgumentException("Waybills must belong to one driver.");

            if (waybills.Any(x => x.Transport is null)) throw new ArgumentException("Waybills must include transport.");

            Driver = new DriverDTO(driver);
            SubTotals = waybills.GroupBy(x => x.TransportId).Select(x => new SubTotal(new TransportDTO(x.First().Transport!), x));
            InitializeFromSubTotals();
        }

        private void InitializeFromSubTotals()
        {
            FuelTopUp = SubTotals.Sum(x => x.FuelTopUp);
            FactFuelConsumption = SubTotals.Sum(x => x.FactFuelConsumption);
            NormalFuelConsumption = SubTotals.Sum(x => x.NormalFuelConsumption);
            FuelEconomy = NormalFuelConsumption - FactFuelConsumption;
        }


        private class SubTotal : DriverFuelSubTotal
        {
            public SubTotal(TransportDTO transport, IEnumerable<Waybill> waybills) : base(transport)
            {
                FuelTopUp = waybills.Sum(x => x.FuelTopUp);
                FactFuelConsumption = waybills.Sum(x => x.FactFuelConsumption);
                NormalFuelConsumption = waybills.Sum(x => x.NormalFuelConsumption);
                FuelEconomy = NormalFuelConsumption - FactFuelConsumption;
            }
        }
    }
}
