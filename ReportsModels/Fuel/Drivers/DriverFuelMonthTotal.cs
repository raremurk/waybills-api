using WaybillsAPI.Models;

namespace WaybillsAPI.ReportsModels.Fuel.Drivers
{
    public class DriverFuelMonthTotal
    {
        public int FuelTopUp { get; private set; }
        public int FactFuelConsumption { get; private set; }
        public int NormalFuelConsumption { get; private set; }
        public int FuelEconomy { get; private set; }

        public DriverFuelMonthTotal(IEnumerable<Waybill> waybills)
        {
            FuelTopUp = waybills.Sum(x => x.FuelTopUp);
            FactFuelConsumption = waybills.Sum(x => x.FactFuelConsumption);
            NormalFuelConsumption = waybills.Sum(x => x.NormalFuelConsumption);
            FuelEconomy = NormalFuelConsumption - FactFuelConsumption;
        }

        public DriverFuelMonthTotal(IEnumerable<DriverFuelMonthTotal> totals)
        {
            FuelTopUp = totals.Sum(x => x.FuelTopUp);
            FactFuelConsumption = totals.Sum(x => x.FactFuelConsumption);
            NormalFuelConsumption = totals.Sum(x => x.NormalFuelConsumption);
            FuelEconomy = NormalFuelConsumption - FactFuelConsumption;
        }
    }
}
